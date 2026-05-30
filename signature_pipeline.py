import os
import re

from SignatureCleaner import process_signature

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_INPUT_FOLDER  = os.path.join(_BASE_DIR, "TempImagesCrop")
DEFAULT_OUTPUT_FOLDER = os.path.join(_BASE_DIR, "TempOutputSignature")

SIGNATURE_PATTERN = re.compile(
    r"^list_(.+)_signature_(.+)\.(png|jpg|jpeg|bmp|tiff|tif|webp)$",
    re.IGNORECASE
)


def get_signature_files(input_folder: str) -> list[str]:
    if not os.path.isdir(input_folder):
        print(f"[ERROR] Папка не найдена: {input_folder}")
        return []

    matched = [
        f for f in os.listdir(input_folder)
        if SIGNATURE_PATTERN.match(f)
    ]

    matched.sort()
    return matched


def process_all_signatures(
    input_folder:  str = DEFAULT_INPUT_FOLDER,
    output_folder: str = DEFAULT_OUTPUT_FOLDER
) -> dict[str, str]:
    os.makedirs(output_folder, exist_ok=True)

    files = get_signature_files(input_folder)

    if not files:
        print("[INFO] Файлов с шаблоном 'list_*_signature_*' не найдено.")
        return {}

    print(f"[INFO] Найдено файлов для обработки: {len(files)}")
    print(f"[INFO] Входная папка : {input_folder}")
    print(f"[INFO] Выходная папка: {output_folder}")
    print("-" * 60)

    results: dict[str, str] = {}

    for idx, filename in enumerate(files, start=1):
        input_path = os.path.join(input_folder, filename)

        base_name, _ = os.path.splitext(filename)
        output_filename = base_name + ".png"
        output_path = os.path.join(output_folder, output_filename)

        print(f"[{idx}/{len(files)}] Обработка: {filename}")

        try:
            process_signature(input_path, output_path)
            results[filename] = output_path
        except Exception as e:
            print(f"  [ERROR] Не удалось обработать '{filename}': {e}")

    print("-" * 60)
    print(f"[DONE] Успешно обработано: {len(results)}/{len(files)} файлов.")

    return results


if __name__ == "__main__":
    process_all_signatures()
