import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
import re
import os
import sys
import subprocess
import threading

import generator

# ============================================================
#  БРЕНДОВАЯ ПАЛИТРА / СТИЛЬ
# ============================================================
NAVY = "#0a0e1a"
NAVY_LIGHT = "#141a2b"
CARD = "#1b2338"
ACCENT = "#1428A0"       # фирменный синий Samsung
ACCENT_LIGHT = "#3b57d6"
GOLD = "#d9b382"
TEXT = "#f2f3f7"
TEXT_DIM = "#9aa2b8"
OK_GREEN = "#3fae6a"
WARN_RED = "#c0505a"

selected_excel = "data/samsung-creative-brief-template.xlsx"


def get_base_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CODES_FILE = os.path.join(get_base_path(), "codes.txt")


def load_codes():
    if os.path.exists(CODES_FILE):
        with open(CODES_FILE, "r") as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]
            if lines:
                return lines
    return ["WW80AK6L28BBLT", "RB34C6B2E5A", "QE55Q60C"]


def save_codes(codes):
    try:
        with open(CODES_FILE, "w") as f:
            for code in codes:
                f.write(code + "\n")
    except Exception:
        pass


def select_file():
    global selected_excel
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        lbl_file.config(text=os.path.basename(filename), fg=TEXT)
        selected_excel = filename
        _log(f"Файл выбран: {os.path.basename(filename)}")


def add_new_model():
    new_code = simpledialog.askstring("Новый товар", "Введите код модели:")
    if new_code:
        new_code = new_code.strip().upper()
        if not re.match(r"^[A-Z0-9\-]{3,}$", new_code):
            messagebox.showerror("Ошибка", "Код должен быть на латинице/цифрах, минимум 3 символа.")
            return
        current_values = list(model_cb["values"])
        if new_code not in current_values:
            current_values.append(new_code)
            model_cb["values"] = current_values
            model_cb.set(new_code)
            save_codes(current_values)
            _log(f"Добавлен код модели: {new_code}")
        else:
            messagebox.showwarning("Внимание", "Код уже есть в списке.")


def delete_current_model():
    current_values = list(model_cb["values"])
    selected_value = model_cb.get()
    if len(current_values) <= 1:
        messagebox.showwarning("Ошибка", "Нельзя удалить последний код!")
        return
    if messagebox.askyesno("Подтверждение", f"Удалить {selected_value}?"):
        current_values.remove(selected_value)
        model_cb["values"] = current_values
        model_cb.current(0)
        save_codes(current_values)


def _log(msg, tag="info"):
    status_list.configure(state="normal")
    status_list.insert(tk.END, msg + "\n", tag)
    status_list.see(tk.END)
    status_list.configure(state="disabled")


def _set_busy(is_busy):
    for w in (btn_gen, btn_gen_all, btn_file, model_cb, lang_cb, retailer_cb):
        w.configure(state=("disabled" if is_busy else "normal"))
    if is_busy:
        progress.start(12)
    else:
        progress.stop()


def _refresh_result_table(output_dir):
    for row in result_tree.get_children():
        result_tree.delete(row)
    if not os.path.exists(output_dir):
        return
    files = sorted(os.listdir(output_dir), reverse=True)
    for f in files:
        if f.startswith("."):
            continue
        full = os.path.join(output_dir, f)
        try:
            size_kb = os.path.getsize(full) / 1024
        except OSError:
            size_kb = 0
        ext = os.path.splitext(f)[1].upper().replace(".", "")
        result_tree.insert("", tk.END, values=(f, ext, f"{size_kb:,.0f} КБ".replace(",", " ")))


def _open_output_folder(output_dir):
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", output_dir])
        elif sys.platform.startswith("win"):
            os.startfile(output_dir)
        else:
            subprocess.run(["xdg-open", output_dir])
    except Exception:
        pass


def _output_dir():
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_dir = os.path.join(desktop_path, "Samsung_Output")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def run():
    def worker():
        try:
            output_dir = _output_dir()
            model = model_cb.get()
            lang = lang_cb.get()
            retailer = retailer_cb.get()
            root.after(0, _log, f"Генерирую {model} · {lang} · {retailer}…")
            created, violations = generator.generate_banner(model, lang, retailer, selected_excel, output_dir)
            root.after(0, _refresh_result_table, output_dir)
            root.after(0, _log, f"Готово: {len(created)} файлов создано.", "ok")
            if violations:
                for v in violations:
                    root.after(0, _log, f"  ⚠ {v}", "warn")
            else:
                root.after(0, _log, "  ✔ safe zone / clear space без нарушений", "ok")
        except Exception as e:
            root.after(0, _log, f"Ошибка: {e}", "warn")
            root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            root.after(0, _set_busy, False)

    _set_busy(True)
    threading.Thread(target=worker, daemon=True).start()


def run_all():
    def worker():
        try:
            output_dir = _output_dir()
            lang = lang_cb.get()
            retailer = retailer_cb.get()
            root.after(0, _log, f"Пакетная генерация всех моделей · {lang} · {retailer}…")
            created, errors = generator.generate_all_banners(lang, retailer, selected_excel, output_dir)
            root.after(0, _refresh_result_table, output_dir)
            root.after(0, _log, f"Готово: {len(created)} файлов по всем моделям.", "ok")
            if errors:
                for e in errors:
                    root.after(0, _log, f"  ⚠ {e}", "warn")
            else:
                root.after(0, _log, "  ✔ все модели без ошибок и нарушений safe zone", "ok")
        except Exception as e:
            root.after(0, _log, f"Ошибка: {e}", "warn")
            root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            root.after(0, _set_busy, False)

    _set_busy(True)
    threading.Thread(target=worker, daemon=True).start()


# ============================================================
#  ОКНО
# ============================================================
root = tk.Tk()
root.title("Samsung Creative Generator")
root.geometry("620x760")
root.minsize(560, 680)
root.configure(bg=NAVY)

style = ttk.Style(root)
try:
    style.theme_use("clam")
except tk.TclError:
    pass

style.configure("TFrame", background=NAVY)
style.configure("Card.TFrame", background=CARD)
style.configure("TLabel", background=NAVY, foreground=TEXT, font=("Helvetica", 11))
style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=("Helvetica", 11))
style.configure("Dim.TLabel", background=CARD, foreground=TEXT_DIM, font=("Helvetica", 10))
style.configure("Section.TLabel", background=NAVY, foreground=TEXT_DIM, font=("Helvetica", 10, "bold"))
style.configure("H1.TLabel", background=NAVY, foreground=TEXT, font=("Helvetica", 20, "bold"))
style.configure("Sub.TLabel", background=NAVY, foreground=TEXT_DIM, font=("Helvetica", 11))

style.configure("TCombobox", fieldbackground=CARD, background=CARD, foreground=TEXT, arrowcolor=TEXT)
style.map("TCombobox", fieldbackground=[("readonly", CARD)], foreground=[("readonly", TEXT)])

style.configure("Accent.TButton", background=ACCENT, foreground="white", font=("Helvetica", 12, "bold"),
                borderwidth=0, focuscolor=ACCENT, padding=(14, 12))
style.map("Accent.TButton", background=[("active", ACCENT_LIGHT)])

style.configure("Ghost.TButton", background=CARD, foreground=TEXT, font=("Helvetica", 10),
                borderwidth=0, padding=(10, 8))
style.map("Ghost.TButton", background=[("active", NAVY_LIGHT)])

style.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT, borderwidth=0, rowheight=26)
style.configure("Treeview.Heading", background=NAVY_LIGHT, foreground=TEXT_DIM, borderwidth=0, font=("Helvetica", 9, "bold"))
style.map("Treeview", background=[("selected", ACCENT)])

style.configure("TProgressbar", troughcolor=NAVY_LIGHT, background=ACCENT, borderwidth=0, thickness=4)

# --- Заголовок ---
header = ttk.Frame(root, style="TFrame")
header.pack(fill="x", padx=24, pady=(22, 4))
ttk.Label(header, text="Samsung Creative Generator", style="H1.TLabel").pack(anchor="w")
ttk.Label(header, text="Excel → бриф → PSD/JPG под 4 формата, автоматически", style="Sub.TLabel").pack(anchor="w", pady=(2, 0))

# --- Шаг 1 ---
ttk.Label(root, text="ШАГ 1 · ИСХОДНЫЕ ДАННЫЕ", style="Section.TLabel").pack(anchor="w", padx=24, pady=(18, 6))
card1 = ttk.Frame(root, style="Card.TFrame")
card1.pack(fill="x", padx=24)
inner1 = ttk.Frame(card1, style="Card.TFrame")
inner1.pack(fill="x", padx=16, pady=14)
btn_file = ttk.Button(inner1, text="Выбрать Excel", style="Ghost.TButton", command=select_file)
btn_file.pack(side="left")
lbl_file = tk.Label(inner1, text="Файл не выбран", bg=CARD, fg=TEXT_DIM, font=("Helvetica", 10))
lbl_file.pack(side="left", padx=12)

# --- Шаг 2 ---
ttk.Label(root, text="ШАГ 2 · ПАРАМЕТРЫ", style="Section.TLabel").pack(anchor="w", padx=24, pady=(18, 6))
card2 = ttk.Frame(root, style="Card.TFrame")
card2.pack(fill="x", padx=24)
inner2 = ttk.Frame(card2, style="Card.TFrame")
inner2.pack(fill="x", padx=16, pady=14)

ttk.Label(inner2, text="Модель", style="Dim.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 3))
model_cb = ttk.Combobox(inner2, values=load_codes(), state="readonly")
model_cb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
model_cb.current(0)

btn_row = ttk.Frame(inner2, style="Card.TFrame")
btn_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 12))
ttk.Button(btn_row, text="+ Добавить код", style="Ghost.TButton", command=add_new_model).pack(side="left", fill="x", expand=True, padx=(0, 4))
ttk.Button(btn_row, text="− Удалить код", style="Ghost.TButton", command=delete_current_model).pack(side="left", fill="x", expand=True, padx=(4, 0))

ttk.Label(inner2, text="Язык", style="Dim.TLabel").grid(row=3, column=0, sticky="w", pady=(0, 3))
ttk.Label(inner2, text="Ритейлер", style="Dim.TLabel").grid(row=3, column=1, sticky="w", pady=(0, 3), padx=(8, 0))
lang_cb = ttk.Combobox(inner2, values=["RU", "KZ"], state="readonly")
lang_cb.grid(row=4, column=0, sticky="ew")
lang_cb.current(0)
retailer_cb = ttk.Combobox(inner2, values=["ALL", "Kaspi", "Sulpak", "Mechta", "Technodom"], state="readonly")
retailer_cb.grid(row=4, column=1, sticky="ew", padx=(8, 0))
retailer_cb.current(0)
inner2.columnconfigure(0, weight=1)
inner2.columnconfigure(1, weight=1)

# --- Шаг 3 ---
ttk.Label(root, text="ШАГ 3 · ГЕНЕРАЦИЯ", style="Section.TLabel").pack(anchor="w", padx=24, pady=(18, 6))
btns = ttk.Frame(root, style="TFrame")
btns.pack(fill="x", padx=24)
btn_gen = ttk.Button(btns, text="Сгенерировать модель (PSD + JPG)", style="Accent.TButton", command=run)
btn_gen.pack(fill="x", pady=(0, 8))
btn_gen_all = ttk.Button(btns, text="Сгенерировать ВСЕ модели", style="Accent.TButton", command=run_all)
btn_gen_all.pack(fill="x")

progress = ttk.Progressbar(root, mode="indeterminate", style="TProgressbar")
progress.pack(fill="x", padx=24, pady=(10, 0))

# --- Результаты ---
ttk.Label(root, text="РЕЗУЛЬТАТ", style="Section.TLabel").pack(anchor="w", padx=24, pady=(18, 6))
result_frame = ttk.Frame(root, style="Card.TFrame")
result_frame.pack(fill="both", expand=True, padx=24)
result_tree = ttk.Treeview(result_frame, columns=("name", "type", "size"), show="headings", height=6)
result_tree.heading("name", text="Файл")
result_tree.heading("type", text="Тип")
result_tree.heading("size", text="Размер")
result_tree.column("name", width=300)
result_tree.column("type", width=60, anchor="center")
result_tree.column("size", width=90, anchor="e")
result_tree.pack(fill="both", expand=True, padx=10, pady=(10, 4))
ttk.Button(result_frame, text="Открыть папку Samsung_Output", style="Ghost.TButton",
           command=lambda: _open_output_folder(_output_dir())).pack(fill="x", padx=10, pady=(0, 10))

# --- Лог / статус ---
ttk.Label(root, text="ЛОГ", style="Section.TLabel").pack(anchor="w", padx=24, pady=(14, 6))
status_list = tk.Text(root, height=5, bg=CARD, fg=TEXT_DIM, insertbackground=TEXT,
                       font=("Menlo", 9), relief="flat", padx=10, pady=8, state="disabled")
status_list.tag_configure("ok", foreground=OK_GREEN)
status_list.tag_configure("warn", foreground=WARN_RED)
status_list.tag_configure("info", foreground=TEXT_DIM)
status_list.pack(fill="x", padx=24, pady=(0, 20))

_log("Готова к работе. Выбери Excel и модель, затем жми «Сгенерировать».")

root.mainloop()
