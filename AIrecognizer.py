import os
import re
import sys
import logging
from pathlib import Path
from typing import Optional

BASE_DIR = Path(__file__).resolve().parent

TEMP_OUTPUT_DIR     = BASE_DIR / "TempOutput"
TEMP_RECOGNIZED_DIR = BASE_DIR / "TempRecognized"

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tiff", ".bmp"}

logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("recognizer")


def _natural_sort_key(path: Path) -> list:
    parts = re.split(r"(\d+)", path.stem)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def _get_word_images(line_dir: Path) -> list[Path]:
    images = [
        f for f in line_dir.iterdir()
        if f.is_file()
        and f.suffix.lower() in IMAGE_EXTENSIONS
        and f.stem.lower().startswith("word_")
    ]
    images.sort(key=_natural_sort_key)
    return images


def _get_line_dirs(part_dir: Path) -> list[Path]:
    lines_dir = part_dir / "lines"
    if not lines_dir.exists():
        log.warning(f"Папка lines не найдена: {lines_dir}")
        return []

    dirs = [
        d for d in lines_dir.iterdir()
        if d.is_dir() and d.name.lower().startswith("linewords_")
    ]
    dirs.sort(key=_natural_sort_key)
    return dirs


def recognize_part(
    list_name: str,
    part_name: str,
    part_dir: Path,
    predictor,
    output_dir: Path,
    batch_size: int = 16,
) -> Path:
    log.info(f"=== Обработка: {list_name} / {part_name} ===")

    line_dirs = _get_line_dirs(part_dir)
    if not line_dirs:
        log.warning(f"Нет строк в {part_dir}")

    recognized_lines: list[str] = []

    for line_dir in line_dirs:
        word_images = _get_word_images(line_dir)

        if not word_images:
            log.debug(f"  Строка {line_dir.name}: изображений не найдено, пропускаем")
            recognized_lines.append("")
            continue

        log.info(f"  Строка {line_dir.name}: {len(word_images)} слов(о)")

        image_paths = [str(p) for p in word_images]
        try:
            words = predictor.predict_batch(image_paths, batch_size=batch_size)
        except Exception as e:
            log.error(f"  Ошибка при распознавании строки {line_dir.name}: {e}")
            words = ["[ОШИБКА]"] * len(word_images)

        line_text = " ".join(w.strip() for w in words if w.strip())
        recognized_lines.append(line_text)

        log.info(f"  → {line_text}")

    output_dir.mkdir(parents=True, exist_ok=True)

    list_idx = re.search(r"\d+", list_name)
    part_idx = re.search(r"\d+", part_name)
    list_suffix = list_idx.group() if list_idx else list_name
    part_suffix = part_idx.group() if part_idx else part_name

    output_filename = f"list_{list_suffix}_part_{part_suffix}.txt"
    output_path = output_dir / output_filename

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(recognized_lines))

    log.info(f"  Сохранено: {output_path}")
    return output_path


def recognize_all(
    predictor=None,
    checkpoint_path: Optional[str] = None,
    temp_output_dir: Path = TEMP_OUTPUT_DIR,
    temp_recognized_dir: Path = TEMP_RECOGNIZED_DIR,
    batch_size: int = 16,
    use_beam: bool = True,
) -> list[Path]:
    if predictor is None:
        if checkpoint_path is None:
            raise ValueError(
                "Укажи predictor= или checkpoint_path= при вызове recognize_all()"
            )
        log.info(f"Загрузка модели из: {checkpoint_path}")
        from inference import RussianHTRPredictor
        predictor = RussianHTRPredictor(
            checkpoint_path=checkpoint_path,
            use_beam=use_beam,
        )
        predictor.warmup()

    if not temp_output_dir.exists():
        raise FileNotFoundError(f"TempOutput не найден: {temp_output_dir}")

    list_dirs = sorted(
        [d for d in temp_output_dir.iterdir()
         if d.is_dir() and d.name.lower().startswith("list_")],
        key=_natural_sort_key,
    )

    if not list_dirs:
        log.warning(f"В {temp_output_dir} не найдено папок list_xxx")
        return []

    created_files: list[Path] = []

    for list_dir in list_dirs:
        list_name = list_dir.name

        part_dirs = sorted(
            [d for d in list_dir.iterdir()
             if d.is_dir() and d.name.lower().startswith("part_")],
            key=_natural_sort_key,
        )

        if not part_dirs:
            log.warning(f"В {list_dir} не найдено папок part_xxx")
            continue

        for part_dir in part_dirs:
            part_name = part_dir.name
            try:
                out_path = recognize_part(
                    list_name=list_name,
                    part_name=part_name,
                    part_dir=part_dir,
                    predictor=predictor,
                    output_dir=temp_recognized_dir,
                    batch_size=batch_size,
                )
                created_files.append(out_path)
            except Exception as e:
                log.error(
                    f"Критическая ошибка при обработке "
                    f"{list_name}/{part_name}: {e}",
                    exc_info=True,
                )

    log.info(f"\nГотово. Создано файлов: {len(created_files)}")
    for p in created_files:
        log.info(f"  {p}")

    return created_files


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Пакетное распознавание рукописного текста"
    )
    parser.add_argument(
        "--checkpoint",
        required=True,
        help="Путь к .pth файлу с весами модели",
    )
    parser.add_argument(
        "--input",
        default=str(TEMP_OUTPUT_DIR),
        help=f"Путь к TempOutput (по умолчанию: {TEMP_OUTPUT_DIR})",
    )
    parser.add_argument(
        "--output",
        default=str(TEMP_RECOGNIZED_DIR),
        help=f"Путь к TempRecognized (по умолчанию: {TEMP_RECOGNIZED_DIR})",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=16,
        help="Размер батча (по умолчанию: 16)",
    )
    parser.add_argument(
        "--greedy",
        action="store_true",
        help="Использовать Greedy вместо Beam Search",
    )
    args = parser.parse_args()

    recognize_all(
        checkpoint_path=args.checkpoint,
        temp_output_dir=Path(args.input),
        temp_recognized_dir=Path(args.output),
        batch_size=args.batch_size,
        use_beam=not args.greedy,
    )
