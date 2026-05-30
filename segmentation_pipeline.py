import os
import re
import glob
import shutil
from collections import defaultdict

import cv2
import numpy as np

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_INPUT_DIR  = os.path.join(_BASE_DIR, "TempImagesCrop")
DEFAULT_OUTPUT_DIR = os.path.join(_BASE_DIR, "TempOutput")

_S1_KERNEL_W    = 20
_S1_KERNEL_H    = 1
_S1_MIN_BOX_W   = 40
_S1_MIN_BOX_H   = 10
_S1_SPLIT_RATIO = 1.2


def _s1_detect_lines(img_bgr: np.ndarray) -> list:
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)

    _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (int(_S1_KERNEL_W), int(_S1_KERNEL_H))
    )
    dilated = cv2.dilate(bw, kernel, iterations=1)
    closed  = cv2.morphologyEx(dilated, cv2.MORPH_CLOSE, kernel, iterations=1)

    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    boxes = [cv2.boundingRect(c) for c in contours]

    boxes = [
        (x, y, w, h) for x, y, w, h in boxes
        if w >= _S1_MIN_BOX_W and h >= _S1_MIN_BOX_H
    ]

    if not boxes:
        return []

    boxes.sort(key=lambda b: b[1])

    heights  = [h for _, _, _, h in boxes]
    median_h = float(np.median(heights))

    split_boxes = []
    for x, y, w, h in boxes:
        if median_h > 0 and h > median_h * _S1_SPLIT_RATIO:
            n_parts = max(2, round(h / median_h))
            part_h  = h / n_parts
            for i in range(n_parts):
                new_y = int(y + i * part_h)
                new_h = int(part_h)
                if new_y + new_h > y + h:
                    new_h = (y + h) - new_y
                split_boxes.append((x, new_y, w, new_h))
        else:
            split_boxes.append((x, y, w, h))

    split_boxes.sort(key=lambda b: b[1])
    return split_boxes


def _s1_process(img_bgr: np.ndarray, lines_dir: str, ext: str) -> int:
    boxes = _s1_detect_lines(img_bgr)
    if not boxes:
        print("    [segm1] Строки не найдены — пропускаем файл.")
        return 0

    os.makedirs(lines_dir, exist_ok=True)

    for idx, (x, y, w, h) in enumerate(boxes, start=1):
        crop = img_bgr[y: y + h, x: x + w]
        out_path = os.path.join(lines_dir, f"line_{idx:03d}{ext}")
        cv2.imwrite(out_path, crop)

    print(f"    [segm1] Найдено строк: {len(boxes)}")
    return len(boxes)


def _s2_process(lines_dir: str) -> None:
    pattern = os.path.join(lines_dir, "line_*.*")
    image_files = [
        f for f in glob.glob(pattern)
        if f.lower().endswith(('.png', '.jpg', '.jpeg'))
    ]
    image_files.sort()

    print(f"    [segm2] Обрабатываем строк: {len(image_files)}")

    for img_path in image_files:
        filename         = os.path.basename(img_path)
        name_without_ext = os.path.splitext(filename)[0]

        if not name_without_ext.startswith("line_"):
            continue

        suffix = name_without_ext[5:]

        img = cv2.imread(img_path)
        if img is None:
            print(f"      [segm2] Не удалось прочитать {img_path}")
            continue

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        kernel  = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 3))
        dilated = cv2.dilate(thresh, kernel, iterations=1)

        contours, _ = cv2.findContours(
            dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        bboxes = [cv2.boundingRect(c) for c in contours]
        if not bboxes:
            continue

        bboxes = sorted(bboxes, key=lambda b: b[0])

        out_dir = os.path.join(lines_dir, f"linewords_{suffix}")
        os.makedirs(out_dir, exist_ok=True)

        word_count = 1
        for (x, y, w, h) in bboxes:
            pixels = w * h
            if (w < 30 and h < 30) or (pixels < 610):
                continue

            word_img  = img[y: y + h, x: x + w]
            word_path = os.path.join(out_dir, f"word_{word_count:03d}.jpg")
            cv2.imwrite(word_path, word_img)
            word_count += 1

        print(f"      [segm2] {filename}: сохранено {word_count - 1} слов → {out_dir}")


def _s3_process(lines_dir: str) -> None:
    folders = glob.glob(os.path.join(lines_dir, "linewords_*"))
    folders.sort()

    print(f"    [segm3] Папок linewords: {len(folders)}")

    for folder in folders:
        folder_name    = os.path.basename(folder)
        new_folder     = folder_name.replace("linewords_", "enhanced_words_")
        save_dir       = os.path.join(lines_dir, new_folder)
        os.makedirs(save_dir, exist_ok=True)

        images = [
            f for f in glob.glob(os.path.join(folder, "*.*"))
            if f.lower().endswith(('.png', '.jpg', '.jpeg'))
        ]

        for img_path in images:
            img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
            if img is None:
                continue

            img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
            img = cv2.convertScaleAbs(img, alpha=1.02, beta=0)
            _, img_bin = cv2.threshold(
                img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
            )

            save_path = os.path.join(save_dir, os.path.basename(img_path))
            cv2.imwrite(save_path, img_bin)

        print(f"      [segm3] {folder_name} → {new_folder}")


_FILENAME_RE = re.compile(
    r"^(list_(\d+))_part_(\d+)(\.[^.]+)$",
    re.IGNORECASE,
)


def _collect_images(input_dir: str) -> dict:
    result = defaultdict(list)

    for fname in os.listdir(input_dir):
        m = _FILENAME_RE.match(fname)
        if not m:
            continue
        list_key  = m.group(1).lower()
        part_num  = int(m.group(3))
        full_path = os.path.join(input_dir, fname)
        result[list_key].append((part_num, full_path))

    for key in result:
        result[key].sort(key=lambda t: t[0])

    return dict(result)


def run_pipeline(
    input_dir:  str = DEFAULT_INPUT_DIR,
    output_dir: str = DEFAULT_OUTPUT_DIR,
) -> None:
    if not os.path.isdir(input_dir):
        raise FileNotFoundError(f"Входная папка не найдена: {input_dir}")

    os.makedirs(output_dir, exist_ok=True)

    groups = _collect_images(input_dir)

    if not groups:
        print("Файлы с именами list_XXX_part_XXX не найдены.")
        return

    print(f"Найдено листов: {len(groups)}")
    print(f"  {', '.join(sorted(groups.keys()))}\n")

    for list_key in sorted(groups.keys()):
        parts = groups[list_key]
        list_out_dir = os.path.join(output_dir, list_key)
        os.makedirs(list_out_dir, exist_ok=True)

        print(f"{'═'*60}")
        print(f"  Лист: {list_key}  ({len(parts)} частей)")
        print(f"  Выходная папка: {list_out_dir}")
        print(f"{'═'*60}")

        for part_num, img_path in parts:
            part_label = f"part_{part_num:03d}"
            print(f"\n  ── {part_label}  ({os.path.basename(img_path)})")

            part_out_dir = os.path.join(list_out_dir, part_label)

            if os.path.exists(part_out_dir):
                shutil.rmtree(part_out_dir)
            os.makedirs(part_out_dir)

            img_bgr = cv2.imread(img_path)
            if img_bgr is None:
                print(f"    Не удалось прочитать файл — пропускаем.")
                continue

            ext = os.path.splitext(img_path)[1].lower() or ".png"

            lines_dir = os.path.join(part_out_dir, "lines")
            n_lines = _s1_process(img_bgr, lines_dir, ext)

            if n_lines == 0:
                print("    Строки не найдены — шаги 2 и 3 пропущены.")
                continue

            _s2_process(lines_dir)
            _s3_process(lines_dir)

        print(f"\n  [{list_key}] Готово!\n")

    print("Весь пайплайн завершён.")


if __name__ == "__main__":
    run_pipeline()
