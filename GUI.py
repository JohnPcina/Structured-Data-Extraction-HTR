import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import json
import os
import datetime
import sys
import glob
import threading
from pathlib import Path

try:
    from PIL import Image, ImageTk, ImageDraw
except ImportError:
    print("=" * 50)
    print("Необходимо установить Pillow:")
    print("  pip install Pillow")
    print("=" * 50)
    sys.exit(1)

BASE_DIR = Path(__file__).resolve().parent

COLORS = {
    "text":      {"outline": "#3B82F6", "fill": "#3B82F620", "tag": "  Текст"},
    "signature": {"outline": "#F59E0B", "fill": "#F59E0B20", "tag": "  Подпись"},
}

BG_DARK   = "#1E1E2E"
BG_PANEL  = "#2A2A3C"
BG_INPUT  = "#3A3A4C"
FG_TEXT   = "#E0E0F0"
FG_DIM    = "#9090A0"
ACCENT    = "#7C3AED"
ACCENT_HV = "#6D28D9"
DANGER    = "#EF4444"
CANVAS_BG = "#111120"

CROP_OUTPUT_DIR = str(BASE_DIR / "TempImagesCrop")

DEFAULT_RECOGNIZED_DIR  = str(BASE_DIR / "TempRecognized")
DEFAULT_SIGNATURE_DIR   = str(BASE_DIR / "TempOutputSignature")
DEFAULT_TEMP_OUTPUT_DIR = str(BASE_DIR / "TempOutput")
DEFAULT_WORD_OUTPUT_DIR = str(BASE_DIR / "WordOutput")

DEFAULT_TEMPLATES_DIR = str(BASE_DIR / "DefaultTemplates")

TEMP_FOLDERS_TO_CLEAR = [
    CROP_OUTPUT_DIR,
    DEFAULT_TEMP_OUTPUT_DIR,
    DEFAULT_SIGNATURE_DIR,
    DEFAULT_RECOGNIZED_DIR,
    DEFAULT_WORD_OUTPUT_DIR,
]

PAGE_W_CM   = 21.0
PAGE_H_CM   = 29.7
MARGIN_CM   = 2.54
CM_PER_PT   = 0.035278
EMU_PER_CM  = 360000
EMU_PER_PT  = 12700


def clear_temp_folders():
    for folder in TEMP_FOLDERS_TO_CLEAR:
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            print(f"[Startup] Не удалось создать папку {folder}: {e}")
            continue

        removed = 0
        failed  = 0
        for entry in os.scandir(folder):
            if entry.is_file(follow_symlinks=False):
                try:
                    os.remove(entry.path)
                    removed += 1
                except Exception as e:
                    print(f"[Startup] Не удалось удалить файл {entry.path}: {e}")
                    failed += 1

        print(f"[Startup] Папка очищена: {folder} "
              f"(удалено файлов: {removed}, ошибок: {failed})")


def create_btn(parent, text, cmd, bg=ACCENT, size=10):
    return tk.Button(
        parent, text=text, command=cmd, bg=bg, fg="white",
        activebackground=ACCENT_HV, activeforeground="white",
        relief="flat", bd=0, font=("Segoe UI", size, "bold"),
        padx=10, pady=4, cursor="hand2"
    )


def px_to_emu(px, img_dim_px, page_dim_cm=PAGE_W_CM, margin_cm=MARGIN_CM):
    work_cm   = page_dim_cm - 2 * margin_cm
    work_emu  = work_cm * EMU_PER_CM
    return int(px / img_dim_px * work_emu)


def box_to_word_emu(box_px, img_w, img_h, margin_cm=MARGIN_CM):
    margin_emu = int(margin_cm * EMU_PER_CM)

    left_offset_emu  = px_to_emu(box_px["x1"], img_w, PAGE_W_CM, margin_cm)
    top_offset_emu   = px_to_emu(box_px["y1"], img_h, PAGE_H_CM, margin_cm)
    right_offset_emu = px_to_emu(box_px["x2"], img_w, PAGE_W_CM, margin_cm)
    bot_offset_emu   = px_to_emu(box_px["y2"], img_h, PAGE_H_CM, margin_cm)

    return {
        "left":   left_offset_emu,
        "top":    top_offset_emu,
        "width":  right_offset_emu - left_offset_emu,
        "height": bot_offset_emu  - top_offset_emu,
    }


def insert_into_word(temp_json_path: str,
                     recognized_dir: str,
                     signature_dir: str,
                     output_dir: str,
                     progress_cb=None):
    print("[Word] Начало insert_into_word")
    print(f"[Word] temp_json_path = {temp_json_path}")
    print(f"[Word] recognized_dir = {recognized_dir}")
    print(f"[Word] signature_dir  = {signature_dir}")
    print(f"[Word] output_dir     = {output_dir}")

    try:
        import win32com.client as win32
        import pythoncom
        print("[Word] win32com и pythoncom импортированы успешно")
    except ImportError as e:
        print(f"[Word] ОШИБКА импорта pywin32: {e}")
        raise RuntimeError(
            "Библиотека pywin32 не установлена.\n"
            "Выполните: pip install pywin32"
        )

    print(f"[Word] Чтение JSON: {temp_json_path}")
    with open(temp_json_path, "r", encoding="utf-8") as f:
        temp_data = json.load(f)
    pages = temp_data.get("pages", [])
    print(f"[Word] JSON прочитан. Страниц: {len(pages)}")

    if not pages:
        return [], ["Нет страниц в Temp.json"]

    pythoncom.CoInitialize()
    print("[Word] CoInitialize выполнен")

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    print("[Word] Word.Application запущен (Visible=False, DisplayAlerts=0)")

    import time
    time.sleep(1.5)
    print("[Word] Пауза после запуска Word завершена")

    errors = []
    created_docs = []

    try:
        print("[Word] Создание единого документа...")
        doc = None
        for attempt in range(1, 6):
            try:
                doc = word.Documents.Add()
                print(f"[Word] Документ создан (попытка {attempt})")
                time.sleep(0.3)
                break
            except Exception as e_add:
                print(f"[Word] Попытка {attempt}/5 — Documents.Add() не удалась: "
                      f"{type(e_add).__name__}: {e_add}")
                time.sleep(1.0 * attempt)

        if doc is None:
            msg = "Не удалось создать документ после 5 попыток"
            print(f"[Word] ОШИБКА: {msg}")
            return [], [msg]

        print("[Word] Настройка PageSetup через Sections(1)...")
        _setup_page(doc, word, errors, page_label="документ")

        for page_num, entry in enumerate(pages):
            page_idx   = entry["pageIndex"]
            img_path   = entry["imagePath"]
            tpl_name   = entry["templateName"]
            img_w      = entry["imageWidth"]
            img_h      = entry["imageHeight"]
            boxes      = entry["boxes"]

            print(f"\n[Word] === Обработка листа {page_idx} ===")
            print(f"[Word] imagePath = {img_path}")
            print(f"[Word] templateName = {tpl_name}")
            print(f"[Word] imageWidth = {img_w}, imageHeight = {img_h}")
            print(f"[Word] boxes count = {len(boxes)}")

            if progress_cb:
                progress_cb(f"Обработка листа {page_idx}: {os.path.basename(img_path)}")

            if page_num > 0:
                print(f"[Word] Вставка разрыва страницы перед листом {page_idx}...")
                try:
                    rng = doc.Content
                    rng.Collapse(0)
                    rng.InsertBreak(7)
                    print(f"[Word] Разрыв страницы вставлен")
                except Exception as e_br:
                    print(f"[Word] ПРЕДУПРЕЖДЕНИЕ: не удалось вставить разрыв страницы: {e_br}")
                    errors.append(f"Лист {page_idx}: ошибка разрыва страницы — {e_br}")

            for box_idx, box in enumerate(boxes, start=1):
                box_type = box["type"]
                emu      = box["wordEmu"]

                print(f"[Word] Box {box_idx}: type={box_type}, emu={emu}")

                try:
                    margin_emu = int(word.InchesToPoints(MARGIN_CM / 2.54) * EMU_PER_PT)
                except Exception:
                    margin_emu = int(MARGIN_CM * EMU_PER_CM)

                final_left_emu   = emu["left"]   + margin_emu
                final_top_emu    = emu["top"]     + margin_emu
                final_width_emu  = emu["width"]
                final_height_emu = emu["height"]

                left_pt   = final_left_emu   / EMU_PER_PT
                top_pt    = final_top_emu    / EMU_PER_PT
                width_pt  = final_width_emu  / EMU_PER_PT
                height_pt = final_height_emu / EMU_PER_PT

                print(f"[Word] Points: left={left_pt:.1f}, top={top_pt:.1f}, "
                      f"width={width_pt:.1f}, height={height_pt:.1f}")

                if width_pt <= 0 or height_pt <= 0:
                    msg = (f"Лист {page_idx}, бокс {box_idx}: "
                           f"нулевой или отрицательный размер — пропущен")
                    print(f"[Word] ПРЕДУПРЕЖДЕНИЕ: {msg}")
                    errors.append(msg)
                    continue

                anchor_range = _get_page_anchor(doc, page_num + 1)

                if box_type == "text":
                    part_stem = f"list_{page_idx:03d}_part_{box_idx:03d}"
                    txt_content = _read_text_file(recognized_dir, part_stem, page_idx, box_idx)

                    print(f"[Word] Добавляем TextBox на страницу {page_num + 1}...")
                    try:
                        tb = doc.Shapes.AddTextbox(
                            1,
                            left_pt,
                            top_pt,
                            width_pt,
                            height_pt,
                            anchor_range
                        )
                        print(f"[Word] TextBox добавлен — OK")

                        tb.TextFrame.TextRange.Text = txt_content
                        tb.TextFrame.WordWrap       = True
                        tb.Line.Visible             = False
                        tb.Fill.Visible             = False
                        tf_range = tb.TextFrame.TextRange

                        tf_range.Font.Size = 14
                        tf_range.Font.Name = "Times New Roman"

                        tb.RelativeHorizontalPosition = 0
                        tb.RelativeVerticalPosition   = 0
                        tb.Left = left_pt
                        tb.Top  = top_pt
                        print(f"[Word] TextBox настроен")
                    except Exception as e_tb:
                        msg = f"Лист {page_idx}, бокс {box_idx}: ошибка TextBox — {e_tb}"
                        print(f"[Word] ОШИБКА TextBox: {type(e_tb).__name__}: {e_tb}")
                        errors.append(msg)

                elif box_type == "signature":
                    sig_stem = f"list_{page_idx:03d}_signature_{box_idx:03d}"
                    sig_candidates = glob.glob(os.path.join(signature_dir, sig_stem + ".*"))
                    print(f"[Word] Подпись: ищем '{sig_stem}' в '{signature_dir}', "
                          f"найдено: {len(sig_candidates)}")

                    if not sig_candidates:
                        msg = f"Подпись не найдена: {sig_stem} (лист {page_idx})"
                        print(f"[Word] ПРЕДУПРЕЖДЕНИЕ: {msg}")
                        errors.append(msg)
                        continue

                    sig_path = sig_candidates[0]
                    print(f"[Word] Вставляем изображение подписи: {sig_path}")
                    try:
                        pic = doc.Shapes.AddPicture(
                            FileName=sig_path,
                            LinkToFile=False,
                            SaveWithDocument=True,
                            Left=left_pt,
                            Top=top_pt,
                            Width=width_pt,
                            Height=height_pt,
                            Anchor=anchor_range
                        )
                        print(f"[Word] Изображение вставлено — OK")

                        pic.LockAspectRatio = False
                        pic.Width  = width_pt
                        pic.Height = height_pt

                        pic.WrapFormat.Type = 3

                        pic.RelativeHorizontalPosition = 0
                        pic.RelativeVerticalPosition   = 0
                        pic.Left = left_pt
                        pic.Top  = top_pt

                        pic.ZOrder(0)

                        print(f"[Word] Изображение подписи настроено")
                    except Exception as e_pic:
                        msg = f"Лист {page_idx}, бокс {box_idx}: ошибка вставки подписи — {e_pic}"
                        print(f"[Word] ОШИБКА вставки подписи: {type(e_pic).__name__}: {e_pic}")
                        errors.append(msg)
                else:
                    print(f"[Word] Неизвестный тип бокса: {box_type} — пропущен")

            if progress_cb:
                progress_cb(f"Лист {page_idx} обработан")

        out_name = "result_all_lists.docx"
        out_path = os.path.join(output_dir, out_name)
        print(f"\n[Word] Сохраняем документ: {out_path}")
        try:
            doc.SaveAs2(out_path, FileFormat=16)
            print(f"[Word] Документ сохранён: {out_path}")
            doc.Close(False)
            print(f"[Word] Документ закрыт")
            created_docs.append(out_path)
        except Exception as e_save:
            print(f"[Word] ОШИБКА сохранения: {type(e_save).__name__}: {e_save}")
            errors.append(f"Ошибка сохранения: {e_save}")
            try:
                doc.Close(False)
            except Exception:
                pass

    except Exception as e_global:
        print(f"[Word] ГЛОБАЛЬНАЯ ОШИБКА: {type(e_global).__name__}: {e_global}")
        import traceback
        traceback.print_exc()
        errors.append(f"Глобальная ошибка: {e_global}")
    finally:
        print("[Word] Завершение: word.Quit()...")
        try:
            word.Quit()
            print("[Word] word.Quit() выполнен")
        except Exception as e_quit:
            print(f"[Word] Ошибка при word.Quit(): {e_quit}")
        pythoncom.CoUninitialize()
        print("[Word] CoUninitialize выполнен")

    print(f"\n[Word] Готово. Создано документов: {len(created_docs)}, ошибок: {len(errors)}")
    return created_docs, errors


def _setup_page(doc, word, errors, page_label="документ"):
    import time
    print(f"[Word] Настройка PageSetup для: {page_label}")

    for ps_attempt in range(1, 4):
        try:
            ps = doc.Sections(1).PageSetup

            margin_pts = word.InchesToPoints(MARGIN_CM / 2.54)
            ps.TopMargin    = margin_pts
            ps.BottomMargin = margin_pts
            ps.LeftMargin   = margin_pts
            ps.RightMargin  = margin_pts
            print(f"[Word] Поля установлены ({margin_pts:.1f} pt) — OK")

            try:
                current_paper = ps.PaperSize
                current_orient = ps.Orientation
                if current_paper != 9:
                    ps.PaperSize = 9
                    print(f"[Word] PaperSize установлен в A4")
                if current_orient != 0:
                    ps.Orientation = 0
                    print(f"[Word] Orientation установлен в Portrait")
            except Exception:
                pass

            print(f"[Word] PageSetup настроен успешно")
            break

        except Exception as e_ps:
            print(f"[Word] PageSetup попытка {ps_attempt}/3 — "
                  f"{type(e_ps).__name__}: {e_ps}")
            if ps_attempt < 3:
                time.sleep(1.0)
            else:
                errors.append(f"{page_label}: ошибка PageSetup — {e_ps}")
                print(f"[Word] PageSetup не удался, продолжаем без настройки полей")


def _get_page_anchor(doc, page_number: int):
    try:
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            try:
                para_page = para.Range.Information(3)
                if para_page == page_number:
                    return para.Range
            except Exception:
                continue
    except Exception as e:
        print(f"[Word] _get_page_anchor: ошибка поиска якоря для стр.{page_number}: {e}")

    try:
        rng = doc.Content
        rng.Collapse(0)
        return rng
    except Exception:
        return None


def _read_text_file(recognized_dir: str, part_stem: str, page_idx: int, box_idx: int) -> str:
    candidates = (
        glob.glob(os.path.join(recognized_dir, part_stem + ".*")) +
        glob.glob(os.path.join(recognized_dir, part_stem))
    )
    print(f"[Word] Текст: ищем '{part_stem}' в '{recognized_dir}', "
          f"найдено кандидатов: {len(candidates)}")

    if candidates:
        print(f"[Word] Читаем: {candidates[0]}")
        try:
            with open(candidates[0], "r", encoding="utf-8") as tf:
                content = tf.read().strip()
            print(f"[Word] Текст прочитан ({len(content)} символов)")
            return content
        except Exception as e:
            print(f"[Word] ОШИБКА чтения файла: {e}")
            return f"[Ошибка чтения: {e}]"
    else:
        print(f"[Word] Файл не найден: {part_stem}")
        return f"[Файл не найден: {part_stem}]"


class Box:
    _counter = 0

    def __init__(self, x1, y1, x2, y2, box_type="text", label=None, id=None):
        Box._counter += 1
        self.id = id if id else Box._counter
        self.x1, self.y1 = min(x1, x2), min(y1, y2)
        self.x2, self.y2 = max(x1, x2), max(y1, y2)
        self.box_type = box_type
        self.label = label or f"{COLORS[box_type]['tag']} {self.id}"
        self.canvas_rect  = None
        self.canvas_label = None

    def real_corners(self, scale, offset_x, offset_y):
        def tr(cx, cy):
            return round((cx - offset_x) / scale, 1), round((cy - offset_y) / scale, 1)
        tl, tr_ = tr(self.x1, self.y1), tr(self.x2, self.y1)
        bl, br  = tr(self.x1, self.y2), tr(self.x2, self.y2)
        return {
            "topLeft":     {"x": tl[0],  "y": tl[1]},
            "topRight":    {"x": tr_[0], "y": tr_[1]},
            "bottomLeft":  {"x": bl[0],  "y": bl[1]},
            "bottomRight": {"x": br[0],  "y": br[1]},
        }

    def width(self):  return abs(self.x2 - self.x1)
    def height(self): return abs(self.y2 - self.y1)


class DefaultTemplateDialog(tk.Toplevel):
    def __init__(self, parent, json_files: list):
        super().__init__(parent)
        self.title("Импорт готового шаблона")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)
        self.geometry("620x480")
        self.grab_set()
        self.focus_set()

        self.selected_path    = None
        self.json_files       = json_files
        self._templates_cache = {}

        self._build_ui()
        self._load_list()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=BG_PANEL, height=50)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="Выберите шаблон для импорта",
            font=("Segoe UI", 12, "bold"), bg=BG_PANEL, fg=FG_TEXT
        ).pack(side="left", padx=14, pady=10)

        body = tk.Frame(self, bg=BG_DARK)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        list_frame = tk.Frame(body, bg=BG_DARK)
        list_frame.pack(side="left", fill="both", expand=True)

        tk.Label(
            list_frame, text="Доступные шаблоны:",
            bg=BG_DARK, fg=FG_DIM, font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(0, 4))

        lb_outer = tk.Frame(list_frame, bg=BG_INPUT)
        lb_outer.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(lb_outer)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            lb_outer,
            bg=BG_INPUT, fg=FG_TEXT,
            font=("Segoe UI", 11),
            selectbackground=ACCENT,
            activestyle="none",
            borderwidth=0, highlightthickness=0,
            yscrollcommand=scrollbar.set
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        self.listbox.bind("<Double-Button-1>",  self._on_double_click)

        detail_frame = tk.Frame(body, bg=BG_DARK, width=200)
        detail_frame.pack(side="right", fill="y", padx=(12, 0))
        detail_frame.pack_propagate(False)

        tk.Label(
            detail_frame, text="Информация:",
            bg=BG_DARK, fg=FG_DIM, font=("Segoe UI", 9)
        ).pack(anchor="w", pady=(0, 4))

        info_bg = tk.Frame(detail_frame, bg=BG_PANEL, padx=10, pady=10)
        info_bg.pack(fill="x")

        self.lbl_name = tk.Label(
            info_bg, text="—", bg=BG_PANEL, fg=FG_TEXT,
            font=("Segoe UI", 10, "bold"), wraplength=170, justify="left"
        )
        self.lbl_name.pack(anchor="w", pady=(0, 6))

        self.lbl_size = tk.Label(
            info_bg, text="", bg=BG_PANEL, fg=FG_DIM,
            font=("Segoe UI", 9), justify="left"
        )
        self.lbl_size.pack(anchor="w")

        self.lbl_boxes = tk.Label(
            info_bg, text="", bg=BG_PANEL, fg=FG_DIM,
            font=("Segoe UI", 9), justify="left"
        )
        self.lbl_boxes.pack(anchor="w")

        self.lbl_date = tk.Label(
            info_bg, text="", bg=BG_PANEL, fg=FG_DIM,
            font=("Segoe UI", 9), justify="left", wraplength=170
        )
        self.lbl_date.pack(anchor="w", pady=(6, 0))

        footer = tk.Frame(self, bg=BG_PANEL, height=54)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        create_btn(
            footer, "✕ Отмена", self.destroy, bg=BG_INPUT
        ).pack(side="right", padx=10, pady=10)

        self.import_btn = create_btn(
            footer, "✅ Загрузить шаблон", self._do_import, bg=ACCENT
        )
        self.import_btn.pack(side="right", padx=4, pady=10)
        self.import_btn.config(state="disabled")

    def _load_list(self):
        for path in self.json_files:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._templates_cache[path] = data
                display_name = data.get("name", os.path.basename(path))
            except Exception:
                display_name = f"⚠ {os.path.basename(path)}"
                self._templates_cache[path] = None

            self.listbox.insert("end", f"  {display_name}")

    def _on_select(self, event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        path = self.json_files[idx]
        data = self._templates_cache.get(path)

        if data:
            self.lbl_name.config(text=data.get("name", "—"))

            w = data.get("imageWidth",  "?")
            h = data.get("imageHeight", "?")
            self.lbl_size.config(text=f"Размер: {w} × {h} px")

            boxes   = data.get("boxes", [])
            n_text  = sum(1 for b in boxes if b.get("type") == "text")
            n_sig   = sum(1 for b in boxes if b.get("type") == "signature")
            self.lbl_boxes.config(
                text=f"Боксов: {len(boxes)}\n"
                     f"  Текст: {n_text}  |  Подпись: {n_sig}"
            )

            raw_date = data.get("createdAt", "")
            if raw_date:
                try:
                    dt = datetime.datetime.fromisoformat(raw_date)
                    self.lbl_date.config(text=f"Создан: {dt.strftime('%d.%m.%Y %H:%M')}")
                except Exception:
                    self.lbl_date.config(text=f"Создан: {raw_date[:16]}")
            else:
                self.lbl_date.config(text="")

            self.import_btn.config(state="normal")
        else:
            self.lbl_name.config(text="Ошибка чтения файла")
            self.lbl_size.config(text="")
            self.lbl_boxes.config(text="")
            self.lbl_date.config(text="")
            self.import_btn.config(state="disabled")

    def _on_double_click(self, event=None):
        sel = self.listbox.curselection()
        if sel:
            self._do_import()

    def _do_import(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        path = self.json_files[idx]
        if self._templates_cache.get(path) is not None:
            self.selected_path = path
            self.destroy()


class AnnotatorApp(tk.Frame):
    def __init__(self, parent, go_back_cmd):
        super().__init__(parent, bg=BG_DARK)
        self.go_back_cmd    = go_back_cmd
        self.image_path     = None
        self.original_image = None
        self.tk_image       = None
        self.scale          = 1.0
        self.offset_x, self.offset_y = 0, 0
        self.img_w, self.img_h = 0, 0
        self.boxes          = []
        self.selected_box   = None
        self.current_type   = "text"
        self.drawing = self.moving = self.resizing = False
        self.start_x = self.start_y = 0
        self.temp_rect = None
        self._pending_import_data = None

        self._build_ui()
        self._bind_events()

    def _build_ui(self):
        top = tk.Frame(self, bg=BG_PANEL, height=50)
        top.pack(fill="x", side="top")
        top.pack_propagate(False)

        create_btn(top, "⬅ Назад", self.go_back_cmd, bg=BG_INPUT).pack(side="left", padx=8)
        tk.Label(top, text="Разметка шаблона", font=("Segoe UI", 12, "bold"),
                 bg=BG_PANEL, fg=FG_TEXT).pack(side="left", padx=8)

        btn_frame = tk.Frame(top, bg=BG_PANEL)
        btn_frame.pack(side="right", padx=10)
        create_btn(btn_frame, "  Открыть фото",   self._open_image).pack(side="left", padx=4)
        create_btn(btn_frame, "  Загрузить JSON", self._load_json).pack(side="left", padx=4)
        create_btn(btn_frame, "  Сохранить JSON", self._save_json).pack(side="left", padx=4)
        create_btn(btn_frame, "  Очистить",       self._clear_all, bg=DANGER).pack(side="left", padx=4)
        create_btn(btn_frame, "Импорт готовых шаблонов",
                   self._import_default_template, bg="#0F766E").pack(side="left", padx=4)

        body = tk.PanedWindow(self, orient="horizontal", bg=BG_DARK, sashwidth=4, sashrelief="flat")
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=CANVAS_BG)
        body.add(left, stretch="always", width=800)
        self.canvas = tk.Canvas(left, bg=CANVAS_BG, highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)

        right = tk.Frame(body, bg=BG_PANEL, width=320)
        body.add(right, stretch="never", width=320)
        right.pack_propagate(False)

        self.type_var = tk.StringVar(value="text")
        for val, info in COLORS.items():
            tk.Radiobutton(
                right, text=info["tag"], variable=self.type_var, value=val,
                bg=BG_PANEL, fg=info["outline"], selectcolor=BG_INPUT,
                font=("Segoe UI", 11), command=self._on_type_change
            ).pack(fill="x", padx=8, pady=2)

        self.box_listbox = tk.Listbox(right, bg=BG_INPUT, fg=FG_TEXT,
                                      font=("Segoe UI", 10), selectbackground=ACCENT)
        self.box_listbox.pack(fill="both", expand=True, padx=8, pady=8)
        self.box_listbox.bind("<<ListboxSelect>>", self._on_listbox_select)

        btn_row = tk.Frame(right, bg=BG_PANEL)
        btn_row.pack(fill="x", padx=8, pady=4)
        create_btn(btn_row, "Переименовать", self._rename_box,     size=9).pack(side="left", fill="x", expand=True, padx=2)
        create_btn(btn_row, "Тип",           self._toggle_box_type, size=9).pack(side="left", fill="x", expand=True, padx=2)
        create_btn(btn_row, "Удалить",       self._delete_selected, bg=DANGER, size=9).pack(side="left", fill="x", expand=True, padx=2)

        self.template_name_var = tk.StringVar(value="Мой шаблон")
        tk.Entry(right, textvariable=self.template_name_var,
                 bg=BG_INPUT, fg=FG_TEXT, font=("Segoe UI", 11)).pack(fill="x", padx=8, pady=8)

        self.status_var = tk.StringVar(value="Откройте изображение или загрузите JSON")
        tk.Label(self, textvariable=self.status_var, bg=BG_PANEL,
                 fg=FG_DIM, anchor="w", padx=10).pack(fill="x", side="bottom")

    def _bind_events(self):
        self.canvas.bind("<ButtonPress-1>",   self._on_press)
        self.canvas.bind("<B1-Motion>",       self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Configure>",       lambda e: self._fit_image())
        self.bind("<Delete>", lambda e: self._delete_selected())
        self.bind("<Escape>", lambda e: self._deselect())

    def _open_image(self):
        path = filedialog.askopenfilename(
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.webp"), ("All", "*.*")])
        if path:
            self.image_path     = path
            self.original_image = Image.open(path)
            self.img_w, self.img_h = self.original_image.size
            self._fit_image()
            self.status_var.set(f"Загружено фото: {os.path.basename(path)}")
            if self._pending_import_data:
                self._apply_pending_import()

    def _fit_image(self):
        if not self.original_image: return
        cw = max(self.canvas.winfo_width(),  10)
        ch = max(self.canvas.winfo_height(), 10)
        self.scale = min(cw / self.img_w, ch / self.img_h, 1.0)
        new_w = int(self.img_w * self.scale)
        new_h = int(self.img_h * self.scale)
        self.offset_x = (cw - new_w) // 2
        self.offset_y = (ch - new_h) // 2

        resized = self.original_image.resize((new_w, new_h), Image.LANCZOS)
        self.tk_image = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.create_image(self.offset_x, self.offset_y, anchor="nw", image=self.tk_image)
        for b in self.boxes:
            self._draw_box(b)

    def _boxes_from_template(self, data: dict) -> list:
        tpl_w = data.get("imageWidth",  self.img_w) or self.img_w
        tpl_h = data.get("imageHeight", self.img_h) or self.img_h

        sx = self.img_w / tpl_w
        sy = self.img_h / tpl_h

        boxes = []
        for b in data.get("boxes", []):
            tl = b["corners"]["topLeft"]
            br = b["corners"]["bottomRight"]

            rx1 = tl["x"] * sx
            ry1 = tl["y"] * sy
            rx2 = br["x"] * sx
            ry2 = br["y"] * sy

            cx1 = rx1 * self.scale + self.offset_x
            cy1 = ry1 * self.scale + self.offset_y
            cx2 = rx2 * self.scale + self.offset_x
            cy2 = ry2 * self.scale + self.offset_y

            box = Box(cx1, cy1, cx2, cy2,
                      b.get("type", "text"),
                      b.get("label", ""))
            boxes.append(box)
        return boxes

    def _load_json(self):
        if not self.original_image:
            messagebox.showwarning("Внимание", "Сначала загрузите изображение!")
            return
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._clear_all()
            self.template_name_var.set(data.get("name", "Загруженный шаблон"))

            for box in self._boxes_from_template(data):
                self.boxes.append(box)
                self._draw_box(box)

            self._refresh_listbox()
            self.status_var.set(
                f"Загружен шаблон: {data.get('name')} "
                f"(источник: {data.get('imageWidth','?')}×{data.get('imageHeight','?')} px, "
                f"текущее фото: {self.img_w}×{self.img_h} px)"
            )
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить JSON: {e}")

    def _import_default_template(self):
        if not os.path.isdir(DEFAULT_TEMPLATES_DIR):
            messagebox.showwarning(
                "Папка не найдена",
                f"Папка с готовыми шаблонами не найдена:\n{DEFAULT_TEMPLATES_DIR}\n\n"
                "Создайте папку и поместите в неё .json файлы шаблонов."
            )
            return

        json_files = glob.glob(os.path.join(DEFAULT_TEMPLATES_DIR, "*.json"))

        if not json_files:
            messagebox.showinfo(
                "Нет шаблонов",
                f"В папке не найдено ни одного JSON-файла:\n{DEFAULT_TEMPLATES_DIR}"
            )
            return

        dlg = DefaultTemplateDialog(self.winfo_toplevel(), json_files)
        self.winfo_toplevel().wait_window(dlg)

        if dlg.selected_path:
            try:
                with open(dlg.selected_path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                self._clear_all()
                self.template_name_var.set(data.get("name", "Импортированный шаблон"))

                if self.original_image:
                    for box in self._boxes_from_template(data):
                        self.boxes.append(box)
                        self._draw_box(box)
                    self._refresh_listbox()
                    self.status_var.set(
                        f"Импортирован шаблон: «{data.get('name')}» "
                        f"({len(self.boxes)} боксов, "
                        f"источник: {data.get('imageWidth','?')}×{data.get('imageHeight','?')} px, "
                        f"текущее фото: {self.img_w}×{self.img_h} px)"
                    )
                else:
                    self._pending_import_data = data
                    self.status_var.set(
                        f"Шаблон «{data.get('name')}» готов. "
                        "Загрузите изображение — боксы применятся автоматически."
                    )

            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить шаблон:\n{e}")

    def _apply_pending_import(self):
        data = self._pending_import_data
        if not data:
            return
        self._pending_import_data = None

        self._clear_all()
        self.template_name_var.set(data.get("name", "Импортированный шаблон"))

        for box in self._boxes_from_template(data):
            self.boxes.append(box)
            self._draw_box(box)

        self._refresh_listbox()
        self.status_var.set(
            f"Шаблон «{data.get('name')}» применён "
            f"({len(self.boxes)} боксов, "
            f"источник: {data.get('imageWidth','?')}×{data.get('imageHeight','?')} px, "
            f"текущее фото: {self.img_w}×{self.img_h} px)"
        )

    def _on_press(self, event):
        if not self.original_image: return
        x, y = event.x, event.y
        for b in reversed(self.boxes):
            if b.x1 <= x <= b.x2 and b.y1 <= y <= b.y2:
                self.moving = True
                self.selected_box = b
                self.move_dx, self.move_dy = x - b.x1, y - b.y1
                self._select_box(b)
                return
        self.drawing = True
        self.start_x, self.start_y = x, y
        self.temp_rect = self.canvas.create_rectangle(
            x, y, x, y,
            outline=COLORS[self.current_type]["outline"], width=2, dash=(4, 4)
        )

    def _on_drag(self, event):
        x, y = event.x, event.y
        if self.drawing and self.temp_rect:
            self.canvas.coords(self.temp_rect, self.start_x, self.start_y, x, y)
        elif self.moving and self.selected_box:
            b = self.selected_box
            w, h = b.width(), b.height()
            b.x1, b.y1 = x - self.move_dx, y - self.move_dy
            b.x2, b.y2 = b.x1 + w, b.y1 + h
            self._update_box_visuals(b)

    def _on_release(self, event):
        if self.drawing:
            self.drawing = False
            if self.temp_rect: self.canvas.delete(self.temp_rect)
            x, y = event.x, event.y
            if abs(x - self.start_x) > 5 and abs(y - self.start_y) > 5:
                box = Box(self.start_x, self.start_y, x, y, self.current_type)
                self.boxes.append(box)
                self._draw_box(box)
                self._select_box(box)
                self._refresh_listbox()
        elif self.moving:
            self.moving = False
            if self.selected_box:
                b = self.selected_box
                b.x1, b.x2 = min(b.x1, b.x2), max(b.x1, b.x2)
                b.y1, b.y2 = min(b.y1, b.y2), max(b.y1, b.y2)
                self._update_box_visuals(b)

    def _draw_box(self, box):
        c = COLORS[box.box_type]["outline"]
        box.canvas_rect  = self.canvas.create_rectangle(box.x1, box.y1, box.x2, box.y2, outline=c, width=2)
        box.canvas_label = self.canvas.create_text(
            box.x1 + 4, box.y1 - 4, text=box.label, anchor="sw",
            fill=c, font=("Segoe UI", 9, "bold")
        )

    def _update_box_visuals(self, box):
        c = COLORS[box.box_type]["outline"]
        if box.canvas_rect:
            self.canvas.coords(box.canvas_rect, box.x1, box.y1, box.x2, box.y2)
            self.canvas.itemconfig(box.canvas_rect, outline=c)
        if box.canvas_label:
            self.canvas.coords(box.canvas_label, box.x1 + 4, box.y1 - 4)
            self.canvas.itemconfig(box.canvas_label, text=box.label, fill=c)

    def _select_box(self, box):
        if self.selected_box and self.selected_box.canvas_rect:
            c = COLORS[self.selected_box.box_type]["outline"]
            self.canvas.itemconfig(self.selected_box.canvas_rect, width=2, outline=c)
        self.selected_box = box
        if box and box.canvas_rect:
            self.canvas.itemconfig(box.canvas_rect, width=3, outline="white")
        if box in self.boxes:
            idx = self.boxes.index(box)
            self.box_listbox.selection_clear(0, "end")
            self.box_listbox.selection_set(idx)

    def _deselect(self):
        if self.selected_box and self.selected_box.canvas_rect:
            c = COLORS[self.selected_box.box_type]["outline"]
            self.canvas.itemconfig(self.selected_box.canvas_rect, width=2, outline=c)
        self.selected_box = None
        self.box_listbox.selection_clear(0, "end")

    def _refresh_listbox(self):
        self.box_listbox.delete(0, "end")
        for b in self.boxes:
            icon = "T" if b.box_type == "text" else "S"
            self.box_listbox.insert("end", f" [{icon}] {b.label}")

    def _on_listbox_select(self, event):
        sel = self.box_listbox.curselection()
        if sel: self._select_box(self.boxes[sel[0]])

    def _on_type_change(self):
        self.current_type = self.type_var.get()

    def _rename_box(self):
        if self.selected_box:
            new_name = simpledialog.askstring("Переименовать", "Имя:", initialvalue=self.selected_box.label)
            if new_name:
                self.selected_box.label = new_name
                self._update_box_visuals(self.selected_box)
                self._refresh_listbox()

    def _toggle_box_type(self):
        if self.selected_box:
            b = self.selected_box
            b.box_type = "signature" if b.box_type == "text" else "text"
            self._update_box_visuals(b)
            self._refresh_listbox()
            self._select_box(b)

    def _delete_selected(self):
        if self.selected_box:
            b = self.selected_box
            if b.canvas_rect:  self.canvas.delete(b.canvas_rect)
            if b.canvas_label: self.canvas.delete(b.canvas_label)
            self.boxes.remove(b)
            self.selected_box = None
            self._refresh_listbox()

    def _clear_all(self):
        for b in self.boxes:
            if b.canvas_rect:  self.canvas.delete(b.canvas_rect)
            if b.canvas_label: self.canvas.delete(b.canvas_label)
        self.boxes.clear()
        self.selected_box = None
        self._refresh_listbox()

    def _save_json(self):
        if not self.boxes:
            return messagebox.showwarning("Внимание", "Нет областей!")
        name = self.template_name_var.get().strip() or "template"
        path = filedialog.asksaveasfilename(
            initialfile=f"{name}.json", defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        if not path: return
        data = {
            "name":        name,
            "sourceImage": os.path.basename(self.image_path) if self.image_path else None,
            "imageWidth":  self.img_w,
            "imageHeight": self.img_h,
            "boxes":       [],
            "createdAt":   datetime.datetime.now().isoformat()
        }
        for b in self.boxes:
            data["boxes"].append({
                "id":      f"box_{b.id}",
                "type":    b.box_type,
                "label":   b.label,
                "corners": b.real_corners(self.scale, self.offset_x, self.offset_y)
            })
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("Готово", f"Сохранено: {path}")


class ImageCard:
    def __init__(self, path, tpl_var, frame):
        self.path    = path
        self.tpl_var = tpl_var
        self.frame   = frame


class ProgressDialog(tk.Toplevel):
    def __init__(self, parent, title="Выполнение..."):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG_DARK)
        self.resizable(True, True)
        self.geometry("680x420")
        self.grab_set()

        tk.Label(self, text=title, font=("Segoe UI", 13, "bold"),
                 bg=BG_DARK, fg=FG_TEXT).pack(pady=(18, 6))

        self.step_var = tk.StringVar(value="Инициализация…")
        self.step_lbl = tk.Label(self, textvariable=self.step_var,
                                 font=("Segoe UI", 10), bg=BG_DARK, fg=ACCENT)
        self.step_lbl.pack(pady=(0, 4))

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Custom.Horizontal.TProgressbar",
                         troughcolor=BG_INPUT, background=ACCENT, thickness=18)
        self.progress = ttk.Progressbar(self, style="Custom.Horizontal.TProgressbar",
                                         orient="horizontal", mode="determinate",
                                         length=580, maximum=100)
        self.progress.pack(pady=6)
        self.progress["value"] = 0

        log_frame = tk.Frame(self, bg=BG_DARK)
        log_frame.pack(fill="both", expand=True, padx=18, pady=(4, 14))

        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")

        self.log_text = tk.Text(log_frame, bg=BG_INPUT, fg=FG_TEXT,
                                font=("Consolas", 9), state="disabled",
                                wrap="word", yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)

    def set_step(self, step_text: str, progress_pct: float = None):
        self.step_var.set(step_text)
        if progress_pct is not None:
            self.progress["value"] = max(0, min(100, progress_pct))
        self.update_idletasks()

    def append_log(self, text: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.update_idletasks()

    def finish(self):
        self.progress["value"] = 100
        self.step_var.set("✅ Завершено!")
        self.grab_release()


class RecognizerApp(tk.Frame):
    def __init__(self, parent, go_back_cmd):
        super().__init__(parent, bg=BG_DARK)
        self.go_back_cmd = go_back_cmd

        self.templates: dict  = {}
        self.image_cards: list = []
        self._selected_idx: int = -1

        self._build_ui()

    def _build_ui(self):
        top = tk.Frame(self, bg=BG_PANEL, height=54)
        top.pack(fill="x", side="top")
        top.pack_propagate(False)

        create_btn(top, "⬅ Назад", self.go_back_cmd, bg=BG_INPUT).pack(side="left", padx=10, pady=8)
        tk.Label(top, text="Нарезка и Word-экспорт",
                 font=("Segoe UI", 12, "bold"), bg=BG_PANEL, fg=FG_TEXT).pack(side="left", padx=6)

        btn_frame = tk.Frame(top, bg=BG_PANEL)
        btn_frame.pack(side="right", padx=10)
        create_btn(btn_frame, "＋ Шаблоны",       self._add_templates).pack(side="left", padx=3)
        create_btn(btn_frame, "＋ Фото",           self._add_images).pack(side="left", padx=3)
        create_btn(btn_frame, "✕ Очистить фото",   self._clear_images, bg=DANGER).pack(side="left", padx=3)

        self.word_btn = create_btn(btn_frame, "W  ВСТАВИТЬ В WORD",
                                   self._full_pipeline, bg="#1D6B9E")
        self.word_btn.pack(side="left", padx=3)

        body = tk.PanedWindow(self, orient="horizontal", bg=BG_DARK, sashwidth=5, sashrelief="flat")
        body.pack(fill="both", expand=True)

        left_col = tk.Frame(body, bg=BG_PANEL)
        body.add(left_col, stretch="never", width=300)

        tpl_hdr = tk.Frame(left_col, bg=BG_PANEL)
        tpl_hdr.pack(fill="x", padx=8, pady=(8, 0))
        tk.Label(tpl_hdr, text="Шаблоны", font=("Segoe UI", 10, "bold"),
                 bg=BG_PANEL, fg=FG_TEXT).pack(side="left")
        self.tpl_count_lbl = tk.Label(tpl_hdr, text="(0)", font=("Segoe UI", 9),
                                      bg=BG_PANEL, fg=FG_DIM)
        self.tpl_count_lbl.pack(side="left", padx=4)

        tpl_list_frame = tk.Frame(left_col, bg=BG_PANEL)
        tpl_list_frame.pack(fill="x", padx=8, pady=4)
        tpl_scroll = tk.Scrollbar(tpl_list_frame, orient="vertical")
        self.tpl_listbox = tk.Listbox(
            tpl_list_frame, bg=BG_INPUT, fg=FG_TEXT,
            font=("Segoe UI", 10), selectbackground=ACCENT,
            height=6, yscrollcommand=tpl_scroll.set, activestyle="none"
        )
        tpl_scroll.config(command=self.tpl_listbox.yview)
        self.tpl_listbox.pack(side="left", fill="x", expand=True)
        tpl_scroll.pack(side="right", fill="y")

        tpl_btn_row = tk.Frame(left_col, bg=BG_PANEL)
        tpl_btn_row.pack(fill="x", padx=8, pady=(0, 6))
        create_btn(tpl_btn_row, "Удалить шаблон", self._remove_template, bg=DANGER, size=9).pack(
            side="left", padx=2)

        tk.Frame(left_col, bg=BG_INPUT, height=1).pack(fill="x", padx=8, pady=4)

        img_hdr = tk.Frame(left_col, bg=BG_PANEL)
        img_hdr.pack(fill="x", padx=8, pady=(4, 0))
        tk.Label(img_hdr, text="Фотографии", font=("Segoe UI", 10, "bold"),
                 bg=BG_PANEL, fg=FG_TEXT).pack(side="left")
        self.img_count_lbl = tk.Label(img_hdr, text="(0)", font=("Segoe UI", 9),
                                      bg=BG_PANEL, fg=FG_DIM)
        self.img_count_lbl.pack(side="left", padx=4)

        img_list_outer = tk.Frame(left_col, bg=BG_PANEL)
        img_list_outer.pack(fill="both", expand=True, padx=8, pady=4)

        img_vbar = tk.Scrollbar(img_list_outer, orient="vertical")
        img_vbar.pack(side="right", fill="y")

        self.img_list_canvas = tk.Canvas(img_list_outer, bg=BG_INPUT,
                                         highlightthickness=0, yscrollcommand=img_vbar.set)
        self.img_list_canvas.pack(side="left", fill="both", expand=True)
        img_vbar.config(command=self.img_list_canvas.yview)

        self.img_list_frame = tk.Frame(self.img_list_canvas, bg=BG_INPUT)
        self._img_list_window = self.img_list_canvas.create_window(
            (0, 0), window=self.img_list_frame, anchor="nw")

        self.img_list_frame.bind(
            "<Configure>",
            lambda e: self.img_list_canvas.configure(
                scrollregion=self.img_list_canvas.bbox("all"))
        )
        self.img_list_canvas.bind(
            "<Configure>",
            lambda e: self.img_list_canvas.itemconfig(
                self._img_list_window, width=e.width)
        )
        self.img_list_canvas.bind("<MouseWheel>", self._on_list_scroll)
        self.img_list_frame.bind("<MouseWheel>",  self._on_list_scroll)

        right_col = tk.Frame(body, bg=BG_DARK)
        body.add(right_col, stretch="always")

        nav_bar = tk.Frame(right_col, bg=BG_PANEL, height=44)
        nav_bar.pack(fill="x")
        nav_bar.pack_propagate(False)

        create_btn(nav_bar, "◀", self._prev_image, bg=BG_INPUT, size=12).pack(side="left", padx=8, pady=6)
        self.nav_label = tk.Label(nav_bar, text="— нет фото —",
                                  font=("Segoe UI", 10), bg=BG_PANEL, fg=FG_DIM)
        self.nav_label.pack(side="left", padx=4)
        create_btn(nav_bar, "▶", self._next_image, bg=BG_INPUT, size=12).pack(side="left", padx=8, pady=6)

        tpl_sel_frame = tk.Frame(nav_bar, bg=BG_PANEL)
        tpl_sel_frame.pack(side="right", padx=12, pady=6)
        tk.Label(tpl_sel_frame, text="Шаблон для этого фото:",
                 font=("Segoe UI", 10), bg=BG_PANEL, fg=FG_TEXT).pack(side="left", padx=(0, 6))

        self.cur_tpl_var = tk.StringVar(value="")
        self.cur_tpl_combo = ttk.Combobox(
            tpl_sel_frame, textvariable=self.cur_tpl_var,
            state="disabled", width=26, font=("Segoe UI", 10)
        )
        self.cur_tpl_combo.pack(side="left")
        self.cur_tpl_combo.bind("<<ComboboxSelected>>", self._on_cur_tpl_changed)

        preview_outer = tk.Frame(right_col, bg=CANVAS_BG)
        preview_outer.pack(fill="both", expand=True)

        self.preview_canvas = tk.Canvas(preview_outer, bg=CANVAS_BG, highlightthickness=0)
        self.preview_canvas.pack(fill="both", expand=True)
        self.preview_canvas.bind("<Configure>", lambda e: self._refresh_preview())

        word_cfg = tk.Frame(right_col, bg=BG_PANEL)
        word_cfg.pack(fill="x", padx=0)

        tk.Label(word_cfg, text="Настройки Word-экспорта",
                 font=("Segoe UI", 9, "bold"), bg=BG_PANEL, fg=FG_DIM).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2))

        tk.Label(word_cfg, text="Папка TempRecognized:",
                 bg=BG_PANEL, fg=FG_TEXT, font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", padx=10)
        self.recognized_dir_var = tk.StringVar(value=DEFAULT_RECOGNIZED_DIR)
        tk.Entry(word_cfg, textvariable=self.recognized_dir_var,
                 bg=BG_INPUT, fg=FG_TEXT, font=("Segoe UI", 9), width=38).grid(row=1, column=1, padx=4, pady=2)
        create_btn(word_cfg, "...",
                   lambda: self.recognized_dir_var.set(
                       filedialog.askdirectory() or self.recognized_dir_var.get()),
                   size=8).grid(row=1, column=2, padx=4)

        tk.Label(word_cfg, text="Папка TempOutputSignature:",
                 bg=BG_PANEL, fg=FG_TEXT, font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", padx=10)
        self.signature_dir_var = tk.StringVar(value=DEFAULT_SIGNATURE_DIR)
        tk.Entry(word_cfg, textvariable=self.signature_dir_var,
                 bg=BG_INPUT, fg=FG_TEXT, font=("Segoe UI", 9), width=38).grid(row=2, column=1, padx=4, pady=2)
        create_btn(word_cfg, "...",
                   lambda: self.signature_dir_var.set(
                       filedialog.askdirectory() or self.signature_dir_var.get()),
                   size=8).grid(row=2, column=2, padx=4)

        tk.Label(word_cfg, text="Папка Word-результата:",
                 bg=BG_PANEL, fg=FG_TEXT, font=("Segoe UI", 9)).grid(row=3, column=0, sticky="w", padx=10)
        self.word_out_dir_var = tk.StringVar(value=DEFAULT_WORD_OUTPUT_DIR)
        tk.Entry(word_cfg, textvariable=self.word_out_dir_var,
                 bg=BG_INPUT, fg=FG_TEXT, font=("Segoe UI", 9), width=38).grid(row=3, column=1, padx=4, pady=2)
        create_btn(word_cfg, "...",
                   lambda: self.word_out_dir_var.set(
                       filedialog.askdirectory() or self.word_out_dir_var.get()),
                   size=8).grid(row=3, column=2, padx=4)

        tk.Label(word_cfg, text="Поля страницы (см):",
                 bg=BG_PANEL, fg=FG_TEXT, font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w", padx=10)
        self.margin_var = tk.StringVar(value=str(MARGIN_CM))
        tk.Entry(word_cfg, textvariable=self.margin_var,
                 bg=BG_INPUT, fg=FG_TEXT, font=("Segoe UI", 9), width=8).grid(row=4, column=1, sticky="w", padx=4, pady=(2, 6))

        tk.Label(word_cfg, text="Checkpoint (.pth):",
                 bg=BG_PANEL, fg=FG_TEXT, font=("Segoe UI", 9)).grid(row=5, column=0, sticky="w", padx=10)
        self.checkpoint_var = tk.StringVar(value=str(BASE_DIR / "checkpoints" / "best_model.pth"))
        tk.Entry(word_cfg, textvariable=self.checkpoint_var,
                 bg=BG_INPUT, fg=FG_TEXT, font=("Segoe UI", 9), width=38).grid(row=5, column=1, padx=4, pady=2)
        create_btn(word_cfg, "...",
                   lambda: self.checkpoint_var.set(
                       filedialog.askopenfilename(
                           filetypes=[("Model checkpoint", "*.pth *.pt"), ("All", "*.*")]
                       ) or self.checkpoint_var.get()),
                   size=8).grid(row=5, column=2, padx=4, pady=(2, 6))

        self.status_var = tk.StringVar(value="Загрузите шаблоны и фотографии")
        tk.Label(self, textvariable=self.status_var,
                 bg=BG_PANEL, fg=FG_DIM, anchor="w", padx=10).pack(fill="x", side="bottom")

    def _on_list_scroll(self, event):
        self.img_list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _add_templates(self):
        paths = filedialog.askopenfilenames(filetypes=[("JSON", "*.json")])
        added = 0
        for p in paths:
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                name = data.get("name", os.path.basename(p))
                if name not in self.templates:
                    self.templates[name] = data
                    self.tpl_listbox.insert("end", name)
                    added += 1
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить {p}: {e}")

        if added:
            self._update_template_lists()
            self.tpl_count_lbl.config(text=f"({len(self.templates)})")
            self.status_var.set(f"Добавлено шаблонов: {added}. Всего: {len(self.templates)}")

    def _remove_template(self):
        sel = self.tpl_listbox.curselection()
        if not sel: return
        name = self.tpl_listbox.get(sel[0])
        self.tpl_listbox.delete(sel[0])
        del self.templates[name]
        self._update_template_lists()
        self.tpl_count_lbl.config(text=f"({len(self.templates)})")

    def _update_template_lists(self):
        tpl_names = list(self.templates.keys())
        self.cur_tpl_combo.config(values=tpl_names)
        if self.cur_tpl_var.get() not in tpl_names:
            self.cur_tpl_var.set(tpl_names[0] if tpl_names else "")
        for card in self.image_cards:
            if card.tpl_var.get() not in tpl_names and tpl_names:
                card.tpl_var.set(tpl_names[0])
        self._refresh_preview()

    def _add_images(self):
        if not self.templates:
            messagebox.showwarning("Внимание", "Сначала загрузите хотя бы один шаблон JSON!")
            return
        paths = filedialog.askopenfilenames(
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.webp"), ("All", "*.*")])
        for p in paths:
            self._create_image_card(p)
        if paths:
            self.img_count_lbl.config(text=f"({len(self.image_cards)})")
            self._select_card(len(self.image_cards) - 1)

    def _create_image_card(self, path):
        idx = len(self.image_cards)
        tpl_var = tk.StringVar(value=list(self.templates.keys())[0] if self.templates else "")

        thumb_frame = tk.Frame(self.img_list_frame, bg=BG_INPUT, cursor="hand2")
        thumb_frame.pack(fill="x", pady=1)

        try:
            img = Image.open(path)
            img.thumbnail((64, 64))
            tk_img = ImageTk.PhotoImage(img)
        except Exception:
            tk_img = None

        lbl_img = tk.Label(thumb_frame, bg=BG_INPUT,
                           image=tk_img if tk_img else None,
                           width=64, height=64)
        lbl_img.image = tk_img
        lbl_img.pack(side="left", padx=4, pady=4)

        name_lbl = tk.Label(thumb_frame, text=os.path.basename(path),
                            bg=BG_INPUT, fg=FG_TEXT, font=("Segoe UI", 9),
                            anchor="w", justify="left", wraplength=160)
        name_lbl.pack(side="left", fill="x", expand=True)

        del_btn = create_btn(thumb_frame, "✕", lambda i=idx: self._remove_by_index(i),
                             bg=DANGER, size=8)
        del_btn.pack(side="right", padx=4)

        card = ImageCard(path, tpl_var, thumb_frame)
        self.image_cards.append(card)

        for w in (thumb_frame, lbl_img, name_lbl):
            w.bind("<Button-1>", lambda e, i=idx: self._select_card(i))

    def _remove_by_index(self, idx):
        if idx < 0 or idx >= len(self.image_cards): return
        card = self.image_cards[idx]
        card.frame.destroy()
        self.image_cards.pop(idx)
        self._rebuild_list_indices()
        if self._selected_idx >= len(self.image_cards):
            self._selected_idx = len(self.image_cards) - 1
        self._select_card(self._selected_idx)
        self.img_count_lbl.config(text=f"({len(self.image_cards)})")

    def _rebuild_list_indices(self):
        for i, card in enumerate(self.image_cards):
            frame = card.frame
            for w in frame.winfo_children():
                if isinstance(w, tk.Button):
                    w.config(command=lambda idx=i: self._remove_by_index(idx))
                elif isinstance(w, (tk.Label, tk.Frame)):
                    w.bind("<Button-1>", lambda e, idx=i: self._select_card(idx))
            frame.bind("<Button-1>", lambda e, idx=i: self._select_card(idx))

    def _clear_images(self):
        for card in self.image_cards:
            card.frame.destroy()
        self.image_cards.clear()
        self._selected_idx = -1
        self._clear_preview()
        self.img_count_lbl.config(text="(0)")
        self.nav_label.config(text="— нет фото —")
        self.cur_tpl_combo.config(state="disabled")
        self.cur_tpl_var.set("")
        self.status_var.set("Список фото очищен")

    def _select_card(self, idx):
        if not self.image_cards:
            self._selected_idx = -1
            self._clear_preview()
            self.nav_label.config(text="— нет фото —")
            self.cur_tpl_combo.config(state="disabled")
            return

        idx = max(0, min(idx, len(self.image_cards) - 1))

        if 0 <= self._selected_idx < len(self.image_cards):
            self.image_cards[self._selected_idx].frame.config(bg=BG_INPUT)
            for w in self.image_cards[self._selected_idx].frame.winfo_children():
                if isinstance(w, tk.Label):
                    w.config(bg=BG_INPUT)

        self._selected_idx = idx
        card = self.image_cards[idx]

        card.frame.config(bg=ACCENT)
        for w in card.frame.winfo_children():
            if isinstance(w, tk.Label):
                w.config(bg=ACCENT)

        self.nav_label.config(
            text=f"{idx + 1} / {len(self.image_cards)}  —  {os.path.basename(card.path)}",
            fg=FG_TEXT
        )

        tpl_names = list(self.templates.keys())
        self.cur_tpl_combo.config(values=tpl_names, state="readonly")
        cur = card.tpl_var.get()
        if cur not in tpl_names and tpl_names:
            cur = tpl_names[0]
            card.tpl_var.set(cur)
        self.cur_tpl_var.set(cur)

        try:
            total = len(self.image_cards)
            if total > 0:
                self.img_list_canvas.yview_moveto(idx / total)
        except Exception:
            pass

        self._refresh_preview()

    def _prev_image(self):
        if self.image_cards:
            self._select_card(self._selected_idx - 1)

    def _next_image(self):
        if self.image_cards:
            self._select_card(self._selected_idx + 1)

    def _on_cur_tpl_changed(self, event=None):
        if 0 <= self._selected_idx < len(self.image_cards):
            new_val = self.cur_tpl_var.get()
            self.image_cards[self._selected_idx].tpl_var.set(new_val)
            self._refresh_preview()

    def _clear_preview(self):
        self.preview_canvas.delete("all")

    def _refresh_preview(self):
        self.preview_canvas.delete("all")

        if self._selected_idx < 0 or self._selected_idx >= len(self.image_cards):
            self.preview_canvas.create_text(
                self.preview_canvas.winfo_width() // 2 or 200,
                self.preview_canvas.winfo_height() // 2 or 200,
                text="Выберите фото из списка слева",
                fill=FG_DIM, font=("Segoe UI", 13)
            )
            return

        card     = self.image_cards[self._selected_idx]
        tpl_name = card.tpl_var.get()

        try:
            img = Image.open(card.path)
        except Exception:
            return

        cw = max(self.preview_canvas.winfo_width(),  10)
        ch = max(self.preview_canvas.winfo_height(), 10)
        scale = min(cw / img.width, ch / img.height, 1.0)
        new_w, new_h = int(img.width * scale), int(img.height * scale)
        off_x = (cw - new_w) // 2
        off_y = (ch - new_h) // 2

        resized = img.resize((new_w, new_h), Image.LANCZOS)
        tk_img  = ImageTk.PhotoImage(resized)
        self.preview_canvas._tk_img = tk_img
        self.preview_canvas.create_image(off_x, off_y, anchor="nw", image=tk_img)

        if tpl_name and tpl_name in self.templates:
            tpl   = self.templates[tpl_name]
            tpl_w = tpl.get("imageWidth",  img.width)
            tpl_h = tpl.get("imageHeight", img.height)
            sx    = img.width  / tpl_w
            sy    = img.height / tpl_h

            for b_idx, b in enumerate(tpl.get("boxes", []), start=1):
                tl = b["corners"]["topLeft"]
                br = b["corners"]["bottomRight"]
                x1 = tl["x"] * sx * scale + off_x
                y1 = tl["y"] * sy * scale + off_y
                x2 = br["x"] * sx * scale + off_x
                y2 = br["y"] * sy * scale + off_y
                color = COLORS.get(b.get("type", "text"), COLORS["text"])["outline"]
                self.preview_canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2)
                self.preview_canvas.create_text(
                    x1 + 4, y1 + 2,
                    text=f"[{b_idx}] {b.get('label', '')}",
                    anchor="nw", fill=color, font=("Segoe UI", 8, "bold")
                )

    def _do_crop(self, out_dir: str, progress_cb=None) -> tuple:
        processed_count = 0
        errors = []

        temp_json = {
            "createdAt": datetime.datetime.now().isoformat(),
            "outputDir": out_dir,
            "pages": []
        }

        try:
            margin_cm = float(self.margin_var.get())
        except ValueError:
            margin_cm = MARGIN_CM

        for page_idx, card in enumerate(self.image_cards, start=1):
            tpl_name = card.tpl_var.get()
            if not tpl_name or tpl_name not in self.templates:
                errors.append(f"Фото №{page_idx}: не задан шаблон")
                continue

            if progress_cb:
                progress_cb(f"Нарезка фото {page_idx}/{len(self.image_cards)}: "
                            f"{os.path.basename(card.path)}")

            try:
                original_img = Image.open(card.path)
                tpl          = self.templates[tpl_name]
                img_w        = original_img.width
                img_h        = original_img.height
                scale_x      = img_w / tpl.get("imageWidth",  1)
                scale_y      = img_h / tpl.get("imageHeight", 1)

                boxes_raw = []
                for b in tpl.get("boxes", []):
                    tl, br = b["corners"]["topLeft"], b["corners"]["bottomRight"]
                    x1, y1 = tl["x"] * scale_x, tl["y"] * scale_y
                    x2, y2 = br["x"] * scale_x, br["y"] * scale_y
                    boxes_raw.append({
                        "x1":    min(x1, x2), "y1": min(y1, y2),
                        "x2":    max(x1, x2), "y2": max(y1, y2),
                        "type":  b.get("type", "text"),
                        "label": b.get("label", "")
                    })

                tolerance = img_h * 0.02
                boxes_raw.sort(
                    key=lambda bx: (round(bx["y1"] / tolerance) * tolerance, bx["x1"]))

                page_entry = {
                    "pageIndex":    page_idx,
                    "imagePath":    card.path,
                    "imageName":    os.path.basename(card.path),
                    "templateName": tpl_name,
                    "imageWidth":   img_w,
                    "imageHeight":  img_h,
                    "boxes":        []
                }

                for i, box in enumerate(boxes_raw, start=1):
                    left   = max(0, int(box["x1"]))
                    top    = max(0, int(box["y1"]))
                    right  = min(img_w, int(box["x2"]))
                    bottom = min(img_h, int(box["y2"]))

                    if right <= left or bottom <= top:
                        continue

                    cropped = original_img.crop((left, top, right, bottom))

                    if cropped.mode in ("RGBA", "LA") or (
                            cropped.mode == "P" and "transparency" in cropped.info):
                        bg_img = Image.new("RGB", cropped.size, (255, 255, 255))
                        if cropped.mode == "P":
                            cropped = cropped.convert("RGBA")
                        bg_img.paste(cropped, mask=cropped.split()[-1])
                        cropped = bg_img
                    elif cropped.mode != "RGB":
                        cropped = cropped.convert("RGB")

                    box_type_str = box["type"]
                    if box_type_str == "signature":
                        out_name = f"list_{page_idx:03d}_signature_{i:03d}.jpg"
                    else:
                        out_name = f"list_{page_idx:03d}_part_{i:03d}.jpg"

                    out_path = os.path.join(out_dir, out_name)
                    cropped.save(out_path, "JPEG")
                    processed_count += 1

                    emu = box_to_word_emu(box, img_w, img_h, margin_cm)

                    page_entry["boxes"].append({
                        "boxIndex":  i,
                        "type":      box["type"],
                        "label":     box["label"],
                        "croppedFile": out_name,
                        "pixelCoords": {
                            "x1": left, "y1": top,
                            "x2": right, "y2": bottom
                        },
                        "wordEmu": emu
                    })

                temp_json["pages"].append(page_entry)

            except Exception as e:
                errors.append(f"Фото №{page_idx} ({os.path.basename(card.path)}): {e}")

        temp_json_path = os.path.join(out_dir, "Temp.json")
        try:
            with open(temp_json_path, "w", encoding="utf-8") as f:
                json.dump(temp_json, f, ensure_ascii=False, indent=2)
        except Exception as e:
            errors.append(f"Не удалось сохранить Temp.json: {e}")
            return None, errors

        print(f"[Crop] Нарезано фрагментов: {processed_count}. Temp.json: {temp_json_path}")
        return temp_json_path, errors

    def _full_pipeline(self):
        if not self.image_cards:
            messagebox.showwarning("Внимание", "Нет добавленных фото!")
            return

        if not self.templates:
            messagebox.showwarning("Внимание", "Нет загруженных шаблонов!")
            return

        crop_dir    = CROP_OUTPUT_DIR
        rec_dir     = self.recognized_dir_var.get().strip()
        sig_dir     = self.signature_dir_var.get().strip()
        word_out    = self.word_out_dir_var.get().strip()
        checkpoint  = self.checkpoint_var.get().strip()

        for d in (crop_dir, rec_dir, sig_dir, word_out):
            try:
                os.makedirs(d, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось создать папку {d}:\n{e}")
                return

        dlg = ProgressDialog(self.winfo_toplevel(), "Полный пайплайн обработки")
        self.word_btn.config(state="disabled")

        def run():
            import time as _time

            all_errors = []
            step_times = {}
            pipeline_start = _time.perf_counter()

            dlg.set_step("Шаг 1/5 — Нарезка фото по шаблону…", 0)
            dlg.append_log("=" * 50)
            dlg.append_log("ШАГ 1: Нарезка фото")
            dlg.append_log(f"  → Папка: {crop_dir}")

            t1_start = _time.perf_counter()
            temp_json_path, crop_errors = self._do_crop(
                crop_dir,
                progress_cb=lambda msg: dlg.append_log(f"  {msg}")
            )
            t1_end = _time.perf_counter()
            step_times["Шаг 1 — Нарезка фото"] = t1_end - t1_start

            if crop_errors:
                for e in crop_errors:
                    dlg.append_log(f"  [WARN] {e}")
                all_errors.extend(crop_errors)

            if temp_json_path is None:
                dlg.append_log("  [ERROR] Нарезка провалилась — прерываем пайплайн.")
                dlg.finish()
                self.after(0, lambda: self.word_btn.config(state="normal"))
                return

            dlg.append_log(f"  Temp.json сохранён: {temp_json_path}")

            dlg.set_step("Шаг 2/5 — Сегментация строк и слов…", 20)
            dlg.append_log("")
            dlg.append_log("=" * 50)
            dlg.append_log("ШАГ 2: segmentation_pipeline.run_pipeline()")
            dlg.append_log(f"  Вход : {crop_dir}")
            dlg.append_log(f"  Выход: {DEFAULT_TEMP_OUTPUT_DIR}")

            t2_start = _time.perf_counter()
            try:
                from segmentation_pipeline import run_pipeline
                run_pipeline(input_dir=crop_dir, output_dir=DEFAULT_TEMP_OUTPUT_DIR)
                dlg.append_log("  Сегментация завершена ✓")
            except Exception as e:
                dlg.append_log(f"  [ERROR] {e}")
                all_errors.append(f"segmentation_pipeline: {e}")
            t2_end = _time.perf_counter()
            step_times["Шаг 2 — Сегментация"] = t2_end - t2_start

            dlg.set_step("Шаг 3/5 — Обработка подписей…", 40)
            dlg.append_log("")
            dlg.append_log("=" * 50)
            dlg.append_log("ШАГ 3: signature_pipeline.process_all_signatures()")
            dlg.append_log(f"  Вход : {crop_dir}")
            dlg.append_log(f"  Выход: {sig_dir}")

            t3_start = _time.perf_counter()
            try:
                from signature_pipeline import process_all_signatures
                results = process_all_signatures(
                    input_folder=crop_dir,
                    output_folder=sig_dir
                )
                dlg.append_log(f"  Подписей обработано: {len(results)} ✓")
            except Exception as e:
                dlg.append_log(f"  [ERROR] {e}")
                all_errors.append(f"signature_pipeline: {e}")
            t3_end = _time.perf_counter()
            step_times["Шаг 3 — Подписи"] = t3_end - t3_start

            dlg.set_step("Шаг 4/5 — AI-распознавание рукописного текста…", 60)
            dlg.append_log("")
            dlg.append_log("=" * 50)
            dlg.append_log("ШАГ 4: AIrecognizer.recognize_all()")
            dlg.append_log(f"  TempOutput    : {DEFAULT_TEMP_OUTPUT_DIR}")
            dlg.append_log(f"  TempRecognized: {rec_dir}")
            dlg.append_log(f"  Checkpoint    : {checkpoint}")

            t4_start = _time.perf_counter()
            try:
                from AIrecognizer import recognize_all
                from pathlib import Path
                created_txts = recognize_all(
                    checkpoint_path=checkpoint,
                    temp_output_dir=Path(DEFAULT_TEMP_OUTPUT_DIR),
                    temp_recognized_dir=Path(rec_dir),
                )
                dlg.append_log(f"  Распознано файлов: {len(created_txts)} ✓")
                for p in created_txts:
                    dlg.append_log(f"    {p}")
            except Exception as e:
                dlg.append_log(f"  [ERROR] {e}")
                all_errors.append(f"AIrecognizer: {e}")
            t4_end = _time.perf_counter()
            step_times["Шаг 4 — AI-распознавание"] = t4_end - t4_start

            dlg.set_step("Шаг 5/5 — Запись в Word…", 80)
            dlg.append_log("")
            dlg.append_log("=" * 50)
            dlg.append_log("ШАГ 5: insert_into_word()")
            dlg.append_log(f"  Temp.json     : {temp_json_path}")
            dlg.append_log(f"  TempRecognized: {rec_dir}")
            dlg.append_log(f"  Signatures    : {sig_dir}")
            dlg.append_log(f"  Word output   : {word_out}")

            t5_start = _time.perf_counter()
            try:
                created_docs, word_errors = insert_into_word(
                    temp_json_path=temp_json_path,
                    recognized_dir=rec_dir,
                    signature_dir=sig_dir,
                    output_dir=word_out,
                    progress_cb=lambda msg: dlg.append_log(f"  {msg}")
                )
                dlg.append_log(f"  Создано документов: {len(created_docs)} ✓")
                for p in created_docs:
                    dlg.append_log(f"    {p}")
                if word_errors:
                    for we in word_errors:
                        dlg.append_log(f"  [WARN] {we}")
                    all_errors.extend(word_errors)
            except Exception as e:
                dlg.append_log(f"  [ERROR] {e}")
                all_errors.append(f"insert_into_word: {e}")
            t5_end = _time.perf_counter()
            step_times["Шаг 5 — Запись в Word"] = t5_end - t5_start

            pipeline_total = _time.perf_counter() - pipeline_start

            dlg.append_log("")
            dlg.append_log("=" * 50)
            if all_errors:
                dlg.append_log(f"Завершено с {len(all_errors)} предупреждениями/ошибками:")
                for err in all_errors:
                    dlg.append_log(f"  error: {err}")
            else:
                dlg.append_log("Весь пайплайн завершён без ошибок!")

            dlg.append_log("")
            dlg.append_log("=" * 50)
            dlg.append_log("⏱  ВРЕМЯ ВЫПОЛНЕНИЯ ЭТАПОВ:")
            dlg.append_log("-" * 50)
            for step_name, elapsed in step_times.items():
                minutes = int(elapsed // 60)
                seconds = elapsed % 60
                if minutes > 0:
                    time_str = f"{minutes} мин {seconds:.2f} сек"
                else:
                    time_str = f"{seconds:.2f} сек"
                dlg.append_log(f"  {step_name:<32} {time_str}")
            dlg.append_log("-" * 50)
            total_min = int(pipeline_total // 60)
            total_sec = pipeline_total % 60
            if total_min > 0:
                total_str = f"{total_min} мин {total_sec:.2f} сек"
            else:
                total_str = f"{total_sec:.2f} сек"
            dlg.append_log(f"  {'ОБЩЕЕ ВРЕМЯ':<32} {total_str}")
            dlg.append_log("=" * 50)

            dlg.finish()
            self.after(0, lambda: (
                self.word_btn.config(state="normal"),
                self.status_var.set(
                    f"Готово. Ошибок: {len(all_errors)}. "
                    f"Общее время: {total_str}. "
                    f"Документ в: {word_out}"
                )
            ))

        t = threading.Thread(target=run, daemon=True)
        t.start()


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Распознавание и Шаблоны")
        self.geometry("1200x750")
        self.minsize(900, 600)
        self.configure(bg=BG_DARK)

        print("[Startup] Очистка временных папок...")
        clear_temp_folders()
        print("[Startup] Очистка завершена.")

        self.current_frame = None
        self.show_start_screen()

    def show_start_screen(self):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = tk.Frame(self, bg=BG_DARK)
        self.current_frame.pack(fill="both", expand=True)

        tk.Label(self.current_frame, text="Добро пожаловать",
                 font=("Segoe UI", 24, "bold"), bg=BG_DARK, fg=FG_TEXT).pack(pady=(120, 40))

        create_btn(self.current_frame, "1.  Создать / Отредактировать Шаблон",
                   self.show_annotator, size=14).pack(pady=10, ipadx=20, ipady=10)
        create_btn(self.current_frame, "2.  Распознать Изображения",
                   self.show_recognizer, bg="#10B981", size=14).pack(pady=10, ipadx=20, ipady=10)

    def show_annotator(self):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = AnnotatorApp(self, self.show_start_screen)
        self.current_frame.pack(fill="both", expand=True)

    def show_recognizer(self):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = RecognizerApp(self, self.show_start_screen)
        self.current_frame.pack(fill="both", expand=True)


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
