import torch
import torch.nn as nn
import numpy as np
from torch.utils.data import DataLoader
import editdistance
from tqdm import tqdm
import argparse, json
from config import *
from dataset import HandwritingDataset, collate_fn, decode_ctc, ALPHABET
from model import build_model


def compute_cer_wer(
    preds: list[str],
    targets: list[str]
) -> tuple[float, float]:
    cer_total = wer_total = 0
    cer_denom = wer_denom = 0

    for pred_str, tgt_str in zip(preds, targets):
        cer_total += editdistance.eval(list(pred_str), list(tgt_str))
        cer_denom += max(len(tgt_str), 1)

        pred_words = pred_str.split()
        tgt_words  = tgt_str.split()
        wer_total += editdistance.eval(pred_words, tgt_words)
        wer_denom += max(len(tgt_words), 1)

    return cer_total / cer_denom, wer_total / wer_denom


def beam_search_decode(
    log_probs: np.ndarray,
    beam_width: int = BEAM_WIDTH
) -> str:
    T, C = log_probs.shape
    beams = [("", BLANK_IDX, 0.0)]

    for t in range(T):
        new_beams = {}
        probs = log_probs[t]
        top_k = np.argsort(probs)[-beam_width:]

        for text, last_c, score in beams:
            for c in top_k:
                p = probs[c]
                if c == BLANK_IDX:
                    key = (text, BLANK_IDX)
                    new_beams[key] = max(
                        new_beams.get(key, -np.inf), score + p
                    )
                elif c == last_c:
                    key = (text, c)
                    new_beams[key] = max(
                        new_beams.get(key, -np.inf), score + p
                    )
                else:
                    new_text = text + ALPHABET[c - 1]
                    key = (new_text, c)
                    new_beams[key] = max(
                        new_beams.get(key, -np.inf), score + p
                    )

        sorted_beams = sorted(new_beams.items(), key=lambda x: x[1], reverse=True)
        beams = [(k[0], k[1], v) for k, v in sorted_beams[:beam_width]]

    return beams[0][0]


def evaluate_model(checkpoint_path: str, use_beam: bool = True):
    model = build_model()
    ckpt  = torch.load(checkpoint_path, map_location=DEVICE)
    model.load_state_dict(ckpt["model_state"])
    model.eval()

    test_ds = HandwritingDataset(TEST_CSV, is_train=False)
    loader  = DataLoader(test_ds, batch_size=16, shuffle=False,
                         num_workers=4, collate_fn=collate_fn)

    all_preds: list[str] = []
    all_targets: list[str] = []
    results = []

    with torch.no_grad():
        for imgs, labels, _, label_lens in tqdm(loader, desc="Evaluating"):
            imgs   = imgs.to(DEVICE)
            logits = model(imgs)
            log_probs = logits.log_softmax(2).permute(1, 0, 2)
            log_probs_np = log_probs.cpu().numpy()

            offset = 0
            for i in range(imgs.size(0)):
                lp = log_probs_np[i]
                if use_beam:
                    pred_str = beam_search_decode(lp)
                else:
                    pred_ids = logits[:, i, :].argmax(1).cpu().tolist()
                    pred_str = decode_ctc(pred_ids)

                length = label_lens[i].item()
                tgt_ids = labels[offset:offset+length].tolist()
                tgt_str = "".join(ALPHABET[j - 1] for j in tgt_ids if j > 0)
                offset += length

                all_preds.append(pred_str)
                all_targets.append(tgt_str)
                results.append({"pred": pred_str, "target": tgt_str,
                                 "correct": pred_str == tgt_str})

    cer, wer = compute_cer_wer(all_preds, all_targets)
    accuracy = sum(r["correct"] for r in results) / len(results)

    print("\n" + "="*60)
    print("  РЕЗУЛЬТАТЫ ОЦЕНКИ МОДЕЛИ")
    print("="*60)
    print(f"  CER (Character Error Rate):  {cer:.4f}  ({cer*100:.2f}%)")
    print(f"  WER (Word Error Rate):       {wer:.4f}  ({wer*100:.2f}%)")
    print(f"  Accuracy (exact match):      {accuracy:.4f}  ({accuracy*100:.2f}%)")
    print(f"  Всего примеров:              {len(results)}")
    print("="*60)
    print("\n  Примеры предсказаний:")
    for r in results[:5]:
        status = "Да" if r["correct"] else "Нет"
        print(f"  {status} Pred:   '{r['pred']}'")
        print(f"     Target: '{r['target']}'\n")

    with open("evaluation_results.json", "w", encoding="utf-8") as f:
        json.dump({"cer": cer, "wer": wer, "accuracy": accuracy,
                   "samples": results[:50]}, f, ensure_ascii=False, indent=2)
    print("  Результаты сохранены в evaluation_results.json")

    return cer, wer, accuracy


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--checkpoint", default=BEST_MODEL)
    parser.add_argument("--beam", action="store_true", default=True)
    args = parser.parse_args()
    evaluate_model(args.checkpoint, args.beam)
