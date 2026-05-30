import torch
import numpy as np
import cv2
from PIL import Image
import albumentations as A
from albumentations.pytorch import ToTensorV2
from typing import Union
from config import *
from dataset import adaptive_resize, clahe_enhance
from model import build_model
from evaluate import beam_search_decode, decode_ctc

INFER_TRANSFORM = A.Compose([
    A.Normalize(mean=(0.5,), std=(0.5,)),
    ToTensorV2(),
])


def preprocess_image(img_input: Union[str, np.ndarray, Image.Image]) -> torch.Tensor:
    if isinstance(img_input, str):
        img = cv2.imread(img_input, cv2.IMREAD_GRAYSCALE)
        if img is None:
            raise FileNotFoundError(f"Не могу загрузить: {img_input}")
    elif isinstance(img_input, np.ndarray):
        if img_input.ndim == 3:
            img = cv2.cvtColor(img_input, cv2.COLOR_BGR2GRAY)
        else:
            img = img_input.copy()
    elif isinstance(img_input, Image.Image):
        img = np.array(img_input.convert("L"))
    else:
        raise TypeError(f"Неподдерживаемый тип: {type(img_input)}")

    if img.mean() < 127:
        img = 255 - img

    img = clahe_enhance(img)
    img = adaptive_resize(img, IMG_HEIGHT, IMG_WIDTH)

    img = np.expand_dims(img, -1)
    tensor = INFER_TRANSFORM(image=img)["image"]

    return tensor.unsqueeze(0)


class RussianHTRPredictor:
    def __init__(
            self,
            checkpoint_path: str = BEST_MODEL,
            device: str = "auto",
            use_beam: bool = True,
            beam_width: int = BEAM_WIDTH,
    ):
        if device == "auto":
            self.device = DEVICE
        else:
            self.device = device

        self.use_beam = use_beam
        self.beam_width = beam_width

        self.model = build_model()
        ckpt = torch.load(checkpoint_path, map_location=self.device)
        state = ckpt.get("model_state", ckpt)
        self.model.load_state_dict(state)
        self.model.eval()

        print(f"[Predictor] Модель загружена: {checkpoint_path}")
        print(f"[Predictor] Устройство: {self.device}, Beam={use_beam}")

    @torch.no_grad()
    def predict(
            self,
            img_input: Union[str, np.ndarray, Image.Image],
            return_confidence: bool = False
    ) -> Union[str, tuple[str, float]]:
        tensor = preprocess_image(img_input).to(self.device)
        logits = self.model(tensor)
        log_probs = logits.log_softmax(2)[:, 0, :].cpu().numpy()

        if self.use_beam:
            text = beam_search_decode(log_probs, self.beam_width)
        else:
            pred_ids = log_probs.argmax(1).tolist()
            text = decode_ctc(pred_ids)

        if return_confidence:
            probs = np.exp(log_probs)
            confidence = float(probs.max(axis=1).mean())
            return text, confidence

        return text

    @torch.no_grad()
    def predict_batch(
            self,
            img_inputs: list[Union[str, np.ndarray, Image.Image]],
            batch_size: int = 16
    ) -> list[str]:
        results = []
        for i in range(0, len(img_inputs), batch_size):
            batch_imgs = img_inputs[i:i + batch_size]
            tensors = [preprocess_image(img) for img in batch_imgs]
            batch = torch.cat(tensors, dim=0).to(self.device)

            logits = self.model(batch)
            log_probs = logits.log_softmax(2).permute(1, 0, 2).cpu().numpy()

            for lp in log_probs:
                if self.use_beam:
                    text = beam_search_decode(lp, self.beam_width)
                else:
                    text = decode_ctc(lp.argmax(1).tolist())
                results.append(text)

        return results

    def warmup(self):
        dummy = torch.zeros(1, IMG_CHANNELS, IMG_HEIGHT, IMG_WIDTH).to(self.device)
        _ = self.model(dummy)
        print("[Predictor] Warmup завершён")


if __name__ == "__main__":
    import argparse, glob, os

    parser = argparse.ArgumentParser(description="Russian Handwriting OCR — Inference")
    parser.add_argument("--input", required=True,
                        help="Путь к изображению или директории с изображениями")
    parser.add_argument("--checkpoint", default=BEST_MODEL)
    parser.add_argument("--beam", action="store_true", default=True)
    parser.add_argument("--output", default=None,
                        help="Файл для сохранения результатов (.txt)")
    args = parser.parse_args()

    predictor = RussianHTRPredictor(args.checkpoint, use_beam=args.beam)
    predictor.warmup()

    if os.path.isdir(args.input):
        paths = glob.glob(os.path.join(args.input, "*.jpg")) + \
                glob.glob(os.path.join(args.input, "*.png")) + \
                glob.glob(os.path.join(args.input, "*.tiff"))
        texts = predictor.predict_batch(paths)
        output_lines = []
        for path, text in zip(paths, texts):
            print(f"{os.path.basename(path)}: {text}")
            output_lines.append(f"{path}\t{text}")
    else:
        text, conf = predictor.predict(args.input, return_confidence=True)
        print(f"Результат: {text}")
        print(f"Уверенность: {conf:.3f}")
        output_lines = [f"{args.input}\t{text}\t{conf:.3f}"]

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write("\n".join(output_lines))
        print(f"Результаты сохранены в: {args.output}")
