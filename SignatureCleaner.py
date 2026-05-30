import sys
import os
from PIL import Image, ImageOps, ImageFilter, ImageMath


def process_signature(input_path, output_path=None):
    if not os.path.exists(input_path):
        print(f"Error: File not found: {input_path}")
        return

    try:
        print(f"Processing: {input_path}...")

        img = Image.open(input_path).convert("RGB")

        r, g, b = img.split()

        background = r.filter(ImageFilter.GaussianBlur(radius=50))

        normalized = ImageMath.eval("convert((float(a) / float(b)) * 255, 'L')", a=r, b=background)

        THRESHOLD = 205
        mask = normalized.point(lambda p: 255 if p < THRESHOLD else 0)

        final_img = Image.new("RGBA", img.size, (0, 0, 0, 0))
        black_ink = Image.new("RGBA", img.size, (0, 0, 0, 255))
        final_img.paste(black_ink, (0, 0), mask)

        if output_path is None:
            folder, filename = os.path.split(input_path)
            name, ext = os.path.splitext(filename)
            output_path = os.path.join(folder, f"{name}_clean.png")

        final_img.save(output_path, "PNG")
        print(f"Success! Saved to: {output_path}")

    except Exception as e:
        print(f"Error processing file: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py  [output_image_path]")
    else:
        in_file = sys.argv[1]
        out_file = sys.argv[2] if len(sys.argv) > 2 else None
        process_signature(in_file, out_file)
