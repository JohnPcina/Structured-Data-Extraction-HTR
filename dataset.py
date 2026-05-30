import cv2
import numpy as np
import torch
from torch.utils.data import Dataset
import albumentations as A
from albumentations.pytorch import ToTensorV2
import pandas as pd
from config import *


def encode_text(text: str) -> list[int]:
    return [ALPHABET.index(c) + 1 for c in text if c in ALPHABET]


def decode_ctc(indices: list[int]) -> str:
    result = []
    prev = -1
    for idx in indices:
        if idx != prev and idx != BLANK_IDX:
            result.append(ALPHABET[idx - 1])
        prev = idx
    return "".join(result)


def adaptive_resize(image: np.ndarray, target_h: int, max_w: int) -> np.ndarray:
    h, w = image.shape[:2]
    ratio = target_h / h
    new_w = min(int(w * ratio), max_w)
    resized = cv2.resize(image, (new_w, target_h), interpolation=cv2.INTER_CUBIC)
    if new_w < max_w:
        pad = np.ones((target_h, max_w - new_w), dtype=np.uint8) * 255
        resized = np.hstack([resized, pad])
    return resized


def deslant(image: np.ndarray) -> np.ndarray:
    img = (255 - image) if image.mean() > 127 else image.copy()
    best_angle = 0
    best_score = -1
    for angle in np.arange(-0.5, 0.5, 0.05):
        M = np.float32([[1, angle, 0], [0, 1, 0]])
        sheared = cv2.warpAffine(
            img, M, (image.shape[1], image.shape[0]),
            borderValue=0
        )
        proj = sheared.sum(axis=1).astype(float)
        score = np.sum((proj[1:] - proj[:-1]) ** 2)
        if score > best_score:
            best_score = score
            best_angle = angle
    M = np.float32([[1, best_angle, 0], [0, 1, 0]])
    return cv2.warpAffine(
        image, M, (image.shape[1], image.shape[0]),
        borderValue=255
    )


def clahe_enhance(image: np.ndarray) -> np.ndarray:
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    return clahe.apply(image)


def binarize_adaptive(image: np.ndarray) -> np.ndarray:
    _, binary = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary


TRAIN_AUGMENT = A.Compose([
    A.OneOf([
        A.GaussianBlur(blur_limit=BLUR_LIMIT, p=0.5),
        A.MotionBlur(blur_limit=5, p=0.3),
        A.MedianBlur(blur_limit=3, p=0.2),
    ], p=AUG_PROB),

    A.RandomBrightnessContrast(
        brightness_limit=BRIGHTNESS_LIMIT,
        contrast_limit=CONTRAST_LIMIT,
        p=AUG_PROB
    ),
    A.GaussNoise(var_limit=(0, NOISE_VAR * 255 ** 2), p=0.3),

    A.OneOf([
        A.ElasticTransform(
            alpha=ELASTIC_ALPHA, sigma=ELASTIC_SIGMA,
            p=0.5
        ),
        A.GridDistortion(num_steps=5, distort_limit=0.3, p=0.3),
        A.OpticalDistortion(distort_limit=0.5, shift_limit=0.05, p=0.2),
    ], p=AUG_PROB),

    A.Rotate(limit=ROTATE_LIMIT, border_mode=cv2.BORDER_CONSTANT,
             value=255, p=0.5),

    A.OneOf([
        A.Morphological(scale=(2, 3), operation="dilation", p=0.5),
        A.Morphological(scale=(2, 3), operation="erosion", p=0.5),
    ], p=0.3),

    A.CoarseDropout(max_holes=4, max_height=8, max_width=20,
                    fill=255, p=0.2),

    A.Normalize(mean=(0.5,), std=(0.5,)),
    ToTensorV2(),
])

VAL_AUGMENT = A.Compose([
    A.Normalize(mean=(0.5,), std=(0.5,)),
    ToTensorV2(),
])


class HandwritingDataset(Dataset):
    def __init__(self, csv_path: str, augment=None, is_train: bool = True):
        self.df = pd.read_csv(csv_path)
        self.augment = augment if augment else (TRAIN_AUGMENT if is_train else VAL_AUGMENT)
        self.is_train = is_train

    def __len__(self):
        return len(self.df)

    def _preprocess(self, image_path: str) -> np.ndarray:
        img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            raise FileNotFoundError(f"Не могу прочитать: {image_path}")
        if img.mean() < 127:
            img = 255 - img
        img = clahe_enhance(img)
        if self.is_train:
            img = deslant(img)
        img = adaptive_resize(img, IMG_HEIGHT, IMG_WIDTH)
        return img

    def __getitem__(self, idx):
        row = self.df.iloc[idx]
        img = self._preprocess(row["image_path"])
        img = np.expand_dims(img, axis=-1)
        augmented = self.augment(image=img)
        img_tensor = augmented["image"]
        label = row["label"]
        encoded = torch.tensor(encode_text(label), dtype=torch.long)
        return img_tensor, encoded, len(img_tensor[0]), len(encoded)


def collate_fn(batch):
    imgs, labels, img_lens, label_lens = zip(*batch)
    imgs = torch.stack(imgs, 0)
    labels_concat = torch.cat(labels)
    img_lens    = torch.tensor(img_lens, dtype=torch.long)
    label_lens  = torch.tensor(label_lens, dtype=torch.long)
    return imgs, labels_concat, img_lens, label_lens
