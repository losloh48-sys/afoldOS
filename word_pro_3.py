import os
import re
import json
import shutil
import tempfile
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, simpledialog
from tkinter import font as tkfont
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

from PIL import Image, ImageTk, ImageGrab

from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4


# ==============================================================
# CONFIGURATION
# ==============================================================

APP_NAME = "Word Pro 3.0"
AUTOSAVE_SECONDS = 5000

DARK = {
    "window": "#101217",
    "top": "#0b0d12",
    "ribbon": "#171a21",
    "ribbon2": "#1e222b",
    "button": "#252a35",
    "button_hover": "#303746",
    "text": "#f3f4f6",
    "muted": "#9ca3af",
    "accent": "#4f8cff",
    "accent_hover": "#6c9dff",
    "border": "#303746",
    "page_bg": "#ffffff",
    "page_text": "#111111",
}

LIGHT = {
    "window": "#e9edf3",
    "top": "#ffffff",
    "ribbon": "#ffffff",
    "ribbon2": "#f3f5f8",
    "button": "#e7eaf0",
    "button_hover": "#d8dde6",
    "text": "#17202a",
    "muted": "#667085",
    "accent": "#2f6fed",
    "accent_hover": "#4f82f4",
    "border": "#d1d7e0",
    "page_bg": "#ffffff",
    "page_text": "#111111",
}


# ==============================================================
# SMALL UI HELPERS
# ==============================================================

class RoundedButton(tk.Button):
    def __init__(self, parent, text, command, theme, **kwargs):
        self.theme = theme

        super().__init__(
            parent,
            text=text,
            command=command,
            bg=theme["button"],
            fg=theme["text"],
            activebackground=theme["accent"],
            activeforeground="white",
            relief="flat",
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=6,
            **kwargs
        )

        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)

    def _enter(self, event):
        self.configure(bg=self.theme["button_hover"])

    def _leave(self, event):
        self.configure(bg=self.theme["button"])


# ==============================================================
# TABLE MODEL
# ==============================================================

class TableModel:
    def __init__(self, rows=3, cols=3):
        self.rows = rows
        self.cols = cols
        self.data = [
            ["" for _ in range(cols)]
            for _ in range(rows)
        ]

    def add_row(self):
        self.data.append(["" for _ in range(self.cols)])
        self.rows += 1

    def add_col(self):
        for row in self.data:
            row.append("")
        self.cols += 1

    def delete_row(self):
        if self.rows <= 1:
            return
        self.data.pop()
        self.rows -= 1

    def delete_col(self):
        if self.cols <= 1:
            return
        for row in self.data:
            row.pop()
        self.cols -= 1


# ==============================================================
# IMAGE OBJECT
# ==============================================================

class ImageObject:
    def __init__(self, app, filename, x=100, y=100, width=350):
        self.app = app
        self.filename = filename
        self.x = x
        self.y = y
        self.width = width
        self.height = 250

        self.selected = False
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.original_width = width
        self.original_height = 250

        self.frame = tk.Frame(
            app.page,
            bg="white",
            highlightthickness=1,
            highlightbackground="#bbbbbb"
        )

        self.frame.place(
            x=self.x,
            y=self.y
        )

        self.label = tk.Label(
            self.frame,
            bg="white"
        )
        self.label.pack()

        self.handle = tk.Label(
            self.frame,
            text="◢",
            bg=app.theme["accent"],
            fg="white",
            cursor="size_nw_se"
        )
        self.handle.place(
            relx=1,
            rely=1,
            anchor="se"
        )

        self.load_image()

        self.frame.bind(
            "<Button-1>",
            self.select
        )

        self.label.bind(
            "<Button-1>",
            self.select
        )

        self.frame.bind(
            "<B1-Motion>",
            self.drag
        )

        self.label.bind(
            "<B1-Motion>",
            self.drag
        )

        self.handle.bind(
            "<Button-1>",
            self.start_resize
        )

        self.handle.bind(
            "<B1-Motion>",
            self.resize
        )

        self.frame.bind(
            "<Double-Button-1>",
            lambda e: self.resize_dialog()
        )

    def load_image(self):
        try:
            image = Image.open(self.filename)

            ratio = image.height / max(image.width, 1)

            self.height = max(
                80,
                int(self.width * ratio)
            )

            image.thumbnail(
                (
                    max(self.width - 4, 50),
                    max(self.height - 4, 50)
                )
            )

            self.photo = ImageTk.PhotoImage(image)

            self.label.configure(
                image=self.photo
            )

            self.label.image = self.photo

            self.frame.configure(
                width=self.width,
                height=self.height
            )

        except Exception:
            self.label.configure(
                text="Не удалось загрузить изображение"
            )

    def select(self, event=None):
        self.app.clear_object_selection()
        self.selected = True

        self.frame.configure(
            highlightthickness=2,
            highlightbackground=self.app.theme["accent"]
        )

        if event:
            self.drag_start_x = event.x_root
            self.drag_start_y = event.y_root

        self.app.status.set(
            "Изображение выбрано • перетаскивание • угол для изменения размера"
        )

    def drag(self, event):
        if not self.selected:
            return

        dx = event.x_root - self.drag_start_x
        dy = event.y_root - self.drag_start_y

        self.x += dx
        self.y += dy

        self.x = max(0, self.x)
        self.y = max(0, self.y)

        self.frame.place(
            x=self.x,
            y=self.y
        )

        self.drag_start_x = event.x_root
        self.drag_start_y = event.y_root

    def start_resize(self, event):
        self.resize_start_x = event.x_root
        self.resize_start_y = event.y_root
        self.resize_start_width = self.frame.winfo_width()

        try:
            image = Image.open(self.filename)
            self.resize_ratio = image.height / max(image.width, 1)
        except Exception:
            self.resize_ratio = 0.7

    def resize(self, event):
        dx = event.x_root - self.resize_start_x

        new_width = max(
            80,
            self.resize_start_width + dx
        )

        new_height = max(
            50,
            int(new_width * self.resize_ratio)
        )

        self.width = new_width
        self.height = new_height

        self.frame.configure(
            width=new_width,
            height=new_height
        )

        self.load_image()

    def resize_dialog(self):
        window = tk.Toplevel(self.app.root)
        window.title("Размер изображения")
        window.geometry("360x270")
        window.configure(bg=self.app.theme["ribbon"])

        width_var = tk.IntVar(
            value=self.frame.winfo_width()
        )

        height_var = tk.IntVar(
            value=self.frame.winfo_height()
        )

        keep_ratio = tk.BooleanVar(value=True)

        tk.Label(
            window,
            text="Ширина",
            bg=self.app.theme["ribbon"],
            fg=self.app.theme["text"]
        ).pack(anchor="w", padx=25, pady=(25, 5))

        width_entry = tk.Entry(
            window,
            textvariable=width_var,
            bg=self.app.theme["button"],
            fg=self.app.theme["text"],
            insertbackground=self.app.theme["text"],
            relief="flat"
        )

        width_entry.pack(
            fill="x",
            padx=25
        )

        tk.Label(
            window,
            text="Высота",
            bg=self.app.theme["ribbon"],
            fg=self.app.theme["text"]
        ).pack(anchor="w", padx=25, pady=(15, 5))

        height_entry = tk.Entry(
            window,
            textvariable=height_var,
            bg=self.app.theme["button"],
            fg=self.app.theme["text"],
            insertbackground=self.app.theme["text"],
            relief="flat"
        )

        height_entry.pack(
            fill="x",
            padx=25
        )

        tk.Checkbutton(
            window,
            text="Сохранять пропорции",
            variable=keep_ratio,
            bg=self.app.theme["ribbon"],
            fg=self.app.theme["text"],
            selectcolor=self.app.theme["button"],
            activebackground=self.app.theme["ribbon"],
            activeforeground=self.app.theme["text"]
        ).pack(
            anchor="w",
            padx=25,
            pady=12
        )

        def apply():
            try:
                w = max(80, int(width_var.get()))
                h = max(50, int(height_var.get()))

                if keep_ratio.get():
                    ratio = self.frame.winfo_height() / max(
                        self.frame.winfo_width(),
                        1
                    )
                    h = int(w * ratio)

                self.width = w
                self.height = h

                self.frame.configure(
                    width=w,
                    height=h
                )

                self.load_image()
                window.destroy()

            except ValueError:
                messagebox.showerror(
                    "Ошибка",
                    "Введите корректные размеры."
                )

        RoundedButton(
            window,
            "Применить",
            apply,
            self.app.theme
        ).pack(pady=10)

    def destroy(self):
        self.frame.destroy()


# ==============================================================
# TABLE WIDGET
# ==============================================================

class TableWidget:
    def __init__(self, app, rows=3, cols=3):
        self.app = app
        self.model = TableModel(rows, cols)
        self.selected = False

        self.frame = tk.Frame(
            app.page,
            bg="white",
            highlightbackground="#aab4c3",
            highlightthickness=1
        )

        self.frame.place(
            x=100,
            y=300
        )

        self.entries = []

        self.rebuild()

        self.frame.bind(
            "<Button-1>",
            self.select
        )

    def rebuild(self):
        values = []

        for row in self.entries:
            values.append([
                entry.get()
                for entry in row
            ])

        for r in range(min(len(values), self.model.rows)):
            for c in range(
                min(len(values[r]), self.model.cols)
            ):
                self.model.data[r][c] = values[r][c]

        for child in self.frame.winfo_children():
            child.destroy()

        self.entries = []

        for r in range(self.model.rows):
            row = []

            for c in range(self.model.cols):
                entry = tk.Entry(
                    self.frame,
                    width=14,
                    bg="white",
                    fg="#111111",
                    insertbackground="#111111",
                    relief="solid",
                    bd=1,
                    font=("Calibri", 10)
                )

                entry.insert(
                    "0",
                    self.model.data[r][c]
                )

                entry.grid(
                    row=r,
                    column=c,
                    sticky="nsew",
                    padx=0,
                    pady=0
                )

                entry.bind(
                    "<FocusIn>",
                    self.select
                )

                row.append(entry)

            self.entries.append(row)

        for c in range(self.model.cols):
            self.frame.grid_columnconfigure(
                c,
                weight=1
            )

    def select(self, event=None):
        self.app.clear_object_selection()

        self.selected = True

        self.frame.configure(
            highlightthickness=2,
            highlightbackground=self.app.theme["accent"]
        )

        self.app.status.set(
            f"Таблица {self.model.rows} × {self.model.cols}"
        )

    def add_row(self):
        self.model.data = self.current_data()
        self.model.add_row()
        self.rebuild()

    def add_col(self):
        self.model.data = self.current_data()
        self.model.add_col()
        self.rebuild()

    def delete_row(self):
        self.model.data = self.current_data()
        self.model.delete_row()
        self.rebuild()

    def delete_col(self):
        self.model.data = self.current_data()
        self.model.delete_col()
        self.rebuild()

    def current_data(self):
        return [
            [entry.get() for entry in row]
            for row in self.entries
        ]

    def destroy(self):
        self.frame.destroy()


# ==============================================================
# MAIN APPLICATION
# ==============================================================

class WordPro:
    def __init__(self, root):
        self.root = root

        self.theme = DARK
        self.dark_mode = True

        self.current_file = None
        self.modified = False

        self.page_width = 794
        self.page_height = 1123

        self.zoom = 1.0

        self.font_name = "Calibri"
        self.font_size = 12

        self.images = []
        self.tables = []

        self.current_object = None

        self.status = tk.StringVar(
            value="Готово"
        )

        self.zoom_text = tk.StringVar(
            value="100%"
        )

        self.document_title = tk.StringVar(
            value="Документ1"
        )

        self.root.title(APP_NAME)
        self.root.geometry("1500x950")
        self.root.minsize(1100, 700)

        self.build_start_screen()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.close
        )

    # ==========================================================
    # START SCREEN
    # ==========================================================

    def build_start_screen(self):
        self.clear()

        self.root.configure(
            bg=self.theme["window"]
        )

        top = tk.Frame(
            self.root,
            bg=self.theme["top"],
            height=70
        )

        top.pack(
            fill="x"
        )

        top.pack_propagate(False)

        tk.Label(
            top,
            text="W",
            bg=self.theme["accent"],
            fg="white",
            font=("Segoe UI", 25, "bold"),
            width=2
        ).pack(
            side="left",
            padx=30,
            pady=10
        )

        tk.Label(
            top,
            text="WORD PRO",
            bg=self.theme["top"],
            fg=self.theme["text"],
            font=("Segoe UI", 21, "bold")
        ).pack(
            side="left"
        )

        tk.Label(
            top,
            text="  3.0",
            bg=self.theme["top"],
            fg=self.theme["accent"],
            font=("Segoe UI", 16, "bold")
        ).pack(
            side="left"
        )

        RoundedButton(
            top,
            "Открыть",
            self.open_document,
            self.theme
        ).pack(
            side="right",
            padx=25
        )

        body = tk.Frame(
            self.root,
            bg=self.theme["window"]
        )

        body.pack(
            fill="both",
            expand=True,
            padx=60,
            pady=50
        )

        tk.Label(
            body,
            text="Что создадим сегодня?",
            bg=self.theme["window"],
            fg=self.theme["text"],
            font=("Segoe UI", 32, "bold")
        ).pack(
            anchor="w"
        )

        tk.Label(
            body,
            text="Современный редактор документов с интерфейсом в стиле Microsoft Word.",
            bg=self.theme["window"],
            fg=self.theme["muted"],
            font=("Segoe UI", 12)
        ).pack(
            anchor="w",
            pady=(8, 40)
        )

        grid = tk.Frame(
            body,
            bg=self.theme["window"]
        )

        grid.pack(
            fill="x"
        )

        templates = [
            ("📄", "Пустой документ", ""),
            ("📚", "Реферат", "Реферат"),
            ("📝", "Доклад", "Доклад"),
            ("📊", "Отчёт", "Отчёт"),
            ("🎓", "Курсовая", "Курсовая работа"),
            ("👤", "Резюме", "Резюме"),
        ]

        for i, (icon, title, content) in enumerate(templates):
            card = tk.Frame(
                grid,
                bg=self.theme["ribbon"],
                highlightbackground=self.theme["border"],
                highlightthickness=1,
                width=220,
                height=170
            )

            card.grid(
                row=0,
                column=i,
                padx=8
            )

            card.grid_propagate(False)

            tk.Label(
                card,
                text=icon,
                bg=self.theme["ribbon"],
                fg=self.theme["text"],
                font=("Segoe UI Emoji", 32)
            ).pack(
                anchor="w",
                padx=20,
                pady=(20, 8)
            )

            tk.Label(
                card,
                text=title,
                bg=self.theme["ribbon"],
                fg=self.theme["text"],
                font=("Segoe UI", 11, "bold")
            ).pack(
                anchor="w",
                padx=20
            )

            def click(c=content):
                self.create_editor(c)

            card.bind(
                "<Button-1>",
                lambda e, c=click: c()
            )

            for child in card.winfo_children():
                child.bind(
                    "<Button-1>",
                    lambda e, c=click: c()
                )

    # ==========================================================
    # EDITOR
    # ==========================================================

    def create_editor(self, initial=""):
        self.clear()

        self.current_file = None
        self.modified = False

        self.root.configure(
            bg=self.theme["window"]
        )

        self.build_titlebar()
        self.build_ribbon()
        self.build_workspace()
        self.build_statusbar()

        self.editor.insert(
            "1.0",
            initial
        )

        self.editor.focus_set()

        self.bind_shortcuts()

        self.editor.edit_modified(False)

        self.root.after(
            AUTOSAVE_SECONDS,
            self.autosave
        )

    # ==========================================================
    # TITLE BAR
    # ==========================================================

    def build_titlebar(self):
        bar = tk.Frame(
            self.root,
            bg=self.theme["top"],
            height=56
        )

        bar.pack(
            fill="x"
        )

        bar.pack_propagate(False)

        tk.Button(
            bar,
            text="←",
            command=self.build_start_screen,
            bg=self.theme["top"],
            fg=self.theme["muted"],
            activebackground=self.theme["button"],
            activeforeground=self.theme["text"],
            relief="flat",
            bd=0,
            font=("Segoe UI", 18),
            cursor="hand2"
        ).pack(
            side="left",
            padx=10
        )

        tk.Label(
            bar,
            text="W",
            bg=self.theme["accent"],
            fg="white",
            font=("Segoe UI", 15, "bold"),
            width=2
        ).pack(
            side="left"
        )

        self.title_label = tk.Label(
            bar,
            textvariable=self.document_title,
            bg=self.theme["top"],
            fg=self.theme["text"],
            font=("Segoe UI", 11, "bold")
        )

        self.title_label.pack(
            side="left",
            padx=12
        )

        RoundedButton(
            bar,
            "💾",
            self.save_document,
            self.theme
        ).pack(
            side="right",
            padx=4
        )

        RoundedButton(
            bar,
            "↷",
            self.redo,
            self.theme
        ).pack(
            side="right",
            padx=4
        )

        RoundedButton(
            bar,
            "↶",
            self.undo,
            self.theme
        ).pack(
            side="right",
            padx=4
        )

    # ==========================================================
    # RIBBON
    # ==========================================================

    def build_ribbon(self):
        self.ribbon = tk.Frame(
            self.root,
            bg=self.theme["ribbon"],
            height=130
        )

        self.ribbon.pack(
            fill="x"
        )

        self.ribbon.pack_propagate(False)

        tabs = tk.Frame(
            self.ribbon,
            bg=self.theme["ribbon"]
        )

        tabs.pack(
            fill="x"
        )

        self.tab_buttons = {}

        for name in [
            "Главная",
            "Вставка",
            "Дизайн",
            "Макет",
            "Ссылки",
            "Рецензирование",
            "Вид"
        ]:
            button = tk.Button(
                tabs,
                text=name,
                command=lambda n=name: self.switch_tab(n),
                bg=self.theme["ribbon"],
                fg=self.theme["muted"],
                activebackground=self.theme["ribbon2"],
                activeforeground=self.theme["text"],
                relief="flat",
                bd=0,
                padx=17,
                pady=7,
                font=("Segoe UI", 9, "bold"),
                cursor="hand2"
            )

            button.pack(
                side="left"
            )

            self.tab_buttons[name] = button

        self.ribbon_content = tk.Frame(
            self.ribbon,
            bg=self.theme["ribbon2"]
        )

        self.ribbon_content.pack(
            fill="both",
            expand=True
        )

        self.switch_tab("Главная")

    def switch_tab(self, name):
        for widget in self.ribbon_content.winfo_children():
            widget.destroy()

        for n, button in self.tab_buttons.items():
            button.configure(
                fg=(
                    self.theme["text"]
                    if n == name
                    else self.theme["muted"]
                )
            )

        if name == "Главная":
            self.tab_home()
        elif name == "Вставка":
            self.tab_insert()
        elif name == "Дизайн":
            self.tab_design()
        elif name == "Макет":
            self.tab_layout()
        elif name == "Ссылки":
            self.tab_references()
        elif name == "Рецензирование":
            self.tab_review()
        elif name == "Вид":
            self.tab_view()

    def group(self, title):
        frame = tk.Frame(
            self.ribbon_content,
            bg=self.theme["ribbon2"]
        )

        frame.pack(
            side="left",
            padx=9,
            pady=2
        )

        tk.Label(
            frame,
            text=title,
            bg=self.theme["ribbon2"],
            fg=self.theme["muted"],
            font=("Segoe UI", 8)
        ).pack(
            side="bottom"
        )

        return frame

    def rb(self, parent, text, command):
        RoundedButton(
            parent,
            text,
            command,
            self.theme
        ).pack(
            side="left",
            padx=2,
            pady=5
        )

    # ==========================================================
    # HOME
    # ==========================================================

    def tab_home(self):
        group = self.group("Буфер обмена")

        self.rb(group, "Вставить", self.paste)
        self.rb(group, "Копировать", self.copy)
        self.rb(group, "Вырезать", self.cut)

        group = self.group("Шрифт")

        self.font_box = ttk.Combobox(
            group,
            values=[
                "Calibri",
                "Arial",
                "Times New Roman",
                "Georgia",
                "Verdana",
                "Tahoma",
                "Courier New",
                "Consolas"
            ],
            width=16
        )

        self.font_box.set(
            self.font_name
        )

        self.font_box.pack(
            side="left",
            padx=3,
            pady=7
        )

        self.font_box.bind(
            "<<ComboboxSelected>>",
            self.change_font
        )

        self.size_box = ttk.Combobox(
            group,
            values=[
                8, 9, 10, 11, 12,
                14, 16, 18, 20, 24,
                28, 32, 36, 48, 72
            ],
            width=5
        )

        self.size_box.set(
            self.font_size
        )

        self.size_box.pack(
            side="left",
            padx=3
        )

        self.size_box.bind(
            "<<ComboboxSelected>>",
            self.change_size
        )

        self.rb(
            group,
            "Ж",
            self.bold
        )

        self.rb(
            group,
            "К",
            self.italic
        )

        self.rb(
            group,
            "Ч",
            self.underline
        )

        self.rb(
            group,
            "S̶",
            self.strike
        )

        self.rb(
            group,
            "Цвет",
            self.text_color
        )

        group = self.group("Абзац")

        self.rb(
            group,
            "←",
            lambda: self.align("left")
        )

        self.rb(
            group,
            "↔",
            lambda: self.align("center")
        )

        self.rb(
            group,
            "→",
            lambda: self.align("right")
        )

        self.rb(
            group,
            "≡",
            lambda: self.align("justify")
        )

        self.rb(
            group,
            "•",
            self.bullets
        )

        self.rb(
            group,
            "1.",
            self.numbering
        )

        group = self.group("Стили")

        self.rb(
            group,
            "Обычный",
            self.style_normal
        )

        self.rb(
            group,
            "Заголовок 1",
            self.style_h1
        )

        self.rb(
            group,
            "Заголовок 2",
            self.style_h2
        )

        self.rb(
            group,
            "Цитата",
            self.style_quote
        )

        group = self.group("Редактирование")

        self.rb(
            group,
            "Найти",
            self.find_text
        )

        self.rb(
            group,
            "Замена",
            self.replace_text
        )

        self.rb(
            group,
            "Статистика",
            self.statistics
        )

    # ==========================================================
    # INSERT
    # ==========================================================

    def tab_insert(self):
        group = self.group("Страницы")

        self.rb(
            group,
            "Разрыв страницы",
            self.page_break
        )

        self.rb(
            group,
            "Номер страницы",
            self.page_number
        )

        group = self.group("Таблицы")

        self.rb(
            group,
            "Таблица",
            self.insert_table
        )

        self.rb(
            group,
            "+ Строка",
            self.table_row_add
        )

        self.rb(
            group,
            "+ Столбец",
            self.table_col_add
        )

        self.rb(
            group,
            "- Строка",
            self.table_row_delete
        )

        self.rb(
            group,
            "- Столбец",
            self.table_col_delete
        )

        group = self.group("Изображения")

        self.rb(
            group,
            "Картинка",
            self.insert_image
        )

        self.rb(
            group,
            "Размер",
            self.resize_image
        )

        self.rb(
            group,
            "Удалить",
            self.delete_object
        )

        group = self.group("Текст")

        self.rb(
            group,
            "Дата",
            self.insert_date
        )

        self.rb(
            group,
            "Символ",
            self.insert_symbol
        )

        self.rb(
            group,
            "Ссылка",
            self.insert_link
        )

        self.rb(
            group,
            "Линия",
            self.insert_line
        )

    # ==========================================================
    # DESIGN
    # ==========================================================

    def tab_design(self):
        group = self.group("Тема")

        self.rb(
            group,
            "Сменить тему",
            self.toggle_theme
        )

        self.rb(
            group,
            "Цвет страницы",
            self.page_color
        )

        group = self.group("Оформление")

        self.rb(
            group,
            "Заголовок",
            self.style_h1
        )

        self.rb(
            group,
            "Подзаголовок",
            self.style_h2
        )

        self.rb(
            group,
            "Цитата",
            self.style_quote
        )

    # ==========================================================
    # LAYOUT
    # ==========================================================

    def tab_layout(self):
        group = self.group("Страница")

        self.rb(
            group,
            "Книжная",
            lambda: self.orientation(False)
        )

        self.rb(
            group,
            "Альбомная",
            lambda: self.orientation(True)
        )

        self.rb(
            group,
            "Поля",
            self.margins
        )

        self.rb(
            group,
            "A4",
            self.a4
        )

        group = self.group("Абзац")

        self.rb(
            group,
            "Отступ +",
            self.indent
        )

        self.rb(
            group,
            "Отступ -",
            self.outdent
        )

    # ==========================================================
    # REFERENCES
    # ==========================================================

    def tab_references(self):
        group = self.group("Оглавление")

        self.rb(
            group,
            "Создать оглавление",
            self.table_of_contents
        )

        self.rb(
            group,
            "Обновить",
            self.update_toc
        )

        group = self.group("Ссылки")

        self.rb(
            group,
            "Гиперссылка",
            self.insert_link
        )

        self.rb(
            group,
            "Закладка",
            self.bookmark
        )

    # ==========================================================
    # REVIEW
    # ==========================================================

    def tab_review(self):
        group = self.group("Проверка")

        self.rb(
            group,
            "Статистика",
            self.statistics
        )

        self.rb(
            group,
            "Поиск",
            self.find_text
        )

        self.rb(
            group,
            "Замена",
            self.replace_text
        )

        group = self.group("Документ")

        self.rb(
            group,
            "Защитить",
            self.protect_document
        )

    # ==========================================================
    # VIEW
    # ==========================================================

    def tab_view(self):
        group = self.group("Масштаб")

        self.rb(
            group,
            "-",
            self.zoom_out
        )

        self.rb(
            group,
            "100%",
            self.zoom_reset
        )

        self.rb(
            group,
            "+",
            self.zoom_in
        )

        group = self.group("Окно")

        self.rb(
            group,
            "Навигация",
            self.navigation
        )

        self.rb(
            group,
            "Полный экран",
            self.fullscreen
        )

        group = self.group("Экспорт")

        self.rb(
            group,
            "PDF",
            self.export_pdf
        )

        self.rb(
            group,
            "Печать",
            self.print_document
        )

    # ==========================================================
    # WORKSPACE
    # ==========================================================

    def build_workspace(self):
        workspace = tk.Frame(
            self.root,
            bg="#080a0f"
        )

        workspace.pack(
            fill="both",
            expand=True
        )

        self.canvas = tk.Canvas(
            workspace,
            bg="#080a0f",
            highlightthickness=0
        )

        self.canvas.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar = ttk.Scrollbar(
            workspace,
            orient="vertical",
            command=self.canvas.yview
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.canvas.configure(
            yscrollcommand=scrollbar.set
        )

        self.page_holder = tk.Frame(
            self.canvas,
            bg="#080a0f"
        )

        self.canvas.create_window(
            (60, 30),
            window=self.page_holder,
            anchor="nw"
        )

        self.create_page()

    def create_page(self):
        self.page = tk.Frame(
            self.page_holder,
            bg=self.theme["page_bg"],
            width=self.page_width,
            height=self.page_height,
            highlightbackground="#cbd5e1",
            highlightthickness=1
        )

        self.page.pack(
            pady=30
        )

        self.page.pack_propagate(False)

        inner = tk.Frame(
            self.page,
            bg=self.theme["page_bg"]
        )

        inner.pack(
            fill="both",
            expand=True,
            padx=75,
            pady=65
        )

        self.editor = tk.Text(
            inner,
            wrap="word",
            undo=True,
            maxundo=-1,
            autoseparators=True,
            bg=self.theme["page_bg"],
            fg=self.theme["page_text"],
            insertbackground=self.theme["page_text"],
            selectbackground=self.theme["accent"],
            selectforeground="white",
            relief="flat",
            bd=0,
            font=("Calibri", 12),
            spacing3=5
        )

        self.editor.pack(
            fill="both",
            expand=True
        )

        self.configure_tags()

        self.editor.bind(
            "<<Modified>>",
            self.modified_event
        )

        self.editor.bind(
            "<KeyRelease>",
            lambda e: self.update_status()
        )

    def configure_tags(self):
        self.editor.tag_configure(
            "bold",
            font=("Calibri", 12, "bold")
        )

        self.editor.tag_configure(
            "italic",
            font=("Calibri", 12, "italic")
        )

        self.editor.tag_configure(
            "underline",
            underline=True
        )

        self.editor.tag_configure(
            "strike",
            overstrike=True
        )

        self.editor.tag_configure(
            "h1",
            font=("Calibri", 22, "bold"),
            spacing3=12
        )

        self.editor.tag_configure(
            "h2",
            font=("Calibri", 17, "bold"),
            spacing3=9
        )

        self.editor.tag_configure(
            "quote",
            font=("Georgia", 12, "italic"),
            lmargin1=25,
            lmargin2=25
        )

    # ==========================================================
    # STATUS
    # ==========================================================

    def build_statusbar(self):
        bar = tk.Frame(
            self.root,
            bg=self.theme["top"],
            height=34
        )

        bar.pack(
            fill="x"
        )

        bar.pack_propagate(False)

        tk.Label(
            bar,
            textvariable=self.status,
            bg=self.theme["top"],
            fg=self.theme["muted"],
            font=("Segoe UI", 9)
        ).pack(
            side="left",
            padx=15
        )

        tk.Label(
            bar,
            textvariable=self.zoom_text,
            bg=self.theme["top"],
            fg=self.theme["text"],
            font=("Segoe UI", 9, "bold")
        ).pack(
            side="right",
            padx=20
        )

    def update_status(self):
        if not hasattr(self, "editor"):
            return

        text = self.editor.get(
            "1.0",
            "end-1c"
        )

        words = re.findall(
            r"\S+",
            text
        )

        lines = int(
            self.editor.index(
                "end-1c"
            ).split(".")[0]
        )

        self.status.set(
            f"Страница 1  •  Слов: {len(words)}  •  "
            f"Символов: {len(text)}  •  Строк: {lines}"
        )

    # ==========================================================
    # FORMATTING
    # ==========================================================

    def selection(self):
        try:
            return (
                self.editor.index("sel.first"),
                self.editor.index("sel.last")
            )
        except tk.TclError:
            return None, None

    def add_tag(self, tag):
        start, end = self.selection()

        if not start:
            self.status.set(
                "Выделите текст для форматирования"
            )
            return

        self.editor.tag_add(
            tag,
            start,
            end
        )

    def bold(self):
        self.add_tag("bold")

    def italic(self):
        self.add_tag("italic")

    def underline(self):
        self.add_tag("underline")

    def strike(self):
        self.add_tag("strike")

    def text_color(self):
        color = colorchooser.askcolor(
            title="Цвет текста"
        )[1]

        if not color:
            return

        start, end = self.selection()

        if not start:
            return

        tag = "color_" + color.replace("#", "")

        self.editor.tag_configure(
            tag,
            foreground=color
        )

        self.editor.tag_add(
            tag,
            start,
            end
        )

    def change_font(self, event=None):
        name = self.font_box.get()

        start, end = self.selection()

        if not start:
            self.font_name = name
            return

        tag = "font_" + re.sub(
            r"\W",
            "_",
            name
        )

        self.editor.tag_configure(
            tag,
            font=(name, self.font_size)
        )

        self.editor.tag_add(
            tag,
            start,
            end
        )

        self.font_name = name

    def change_size(self, event=None):
        try:
            size = int(
                self.size_box.get()
            )
        except ValueError:
            return

        self.font_size = size

        start, end = self.selection()

        if not start:
            return

        tag = f"size_{size}"

        self.editor.tag_configure(
            tag,
            font=(self.font_name, size)
        )

        self.editor.tag_add(
            tag,
            start,
            end
        )

    def align(self, value):
        start, end = self.selection()

        if not start:
            start = self.editor.index(
                "insert linestart"
            )
            end = self.editor.index(
                "insert lineend"
            )

        tag = f"align_{value}"

        self.editor.tag_configure(
            tag,
            justify=value
        )

        self.editor.tag_add(
            tag,
            start,
            end
        )

    def bullets(self):
        start, end = self.selection()

        if not start:
            start = self.editor.index(
                "insert linestart"
            )
            end = self.editor.index(
                "insert lineend"
            )

        text = self.editor.get(
            start,
            end
        )

        lines = text.splitlines()

        self.editor.delete(
            start,
            end
        )

        self.editor.insert(
            start,
            "\n".join(
                "• " + line
                if line.strip()
                else line
                for line in lines
            )
        )

    def numbering(self):
        start, end = self.selection()

        if not start:
            start = self.editor.index(
                "insert linestart"
            )
            end = self.editor.index(
                "insert lineend"
            )

        text = self.editor.get(
            start,
            end
        )

        result = []
        number = 1

        for line in text.splitlines():
            if line.strip():
                result.append(
                    f"{number}. {line}"
                )
                number += 1
            else:
                result.append("")

        self.editor.delete(
            start,
            end
        )

        self.editor.insert(
            start,
            "\n".join(result)
        )

    # ==========================================================
    # STYLES
    # ==========================================================

    def style_normal(self):
        self.add_tag("normal")

    def style_h1(self):
        start = self.editor.index(
            "insert linestart"
        )

        end = self.editor.index(
            "insert lineend"
        )

        self.editor.tag_add(
            "h1",
            start,
            end
        )

    def style_h2(self):
        start = self.editor.index(
            "insert linestart"
        )

        end = self.editor.index(
            "insert lineend"
        )

        self.editor.tag_add(
            "h2",
            start,
            end
        )

    def style_quote(self):
        start = self.editor.index(
            "insert linestart"
        )

        end = self.editor.index(
            "insert lineend"
        )

        self.editor.tag_add(
            "quote",
            start,
            end
        )

    # ==========================================================
    # CLIPBOARD
    # ==========================================================

    def copy(self):
        try:
            self.editor.event_generate(
                "<<Copy>>"
            )
        except:
            pass

    def cut(self):
        try:
            self.editor.event_generate(
                "<<Cut>>"
            )
        except:
            pass

    def paste(self):
        try:
            self.editor.event_generate(
                "<<Paste>>"
            )
        except:
            pass

    def undo(self):
        try:
            self.editor.edit_undo()
        except:
            pass

    def redo(self):
        try:
            self.editor.edit_redo()
        except:
            pass

    # ==========================================================
    # TABLES
    # ==========================================================

    def insert_table(self):
        dialog = tk.Toplevel(
            self.root
        )

        dialog.title(
            "Вставка таблицы"
        )

        dialog.geometry(
            "350x250"
        )

        dialog.configure(
            bg=self.theme["ribbon"]
        )

        rows = tk.IntVar(value=3)
        cols = tk.IntVar(value=3)

        tk.Label(
            dialog,
            text="Строк:",
            bg=self.theme["ribbon"],
            fg=self.theme["text"]
        ).pack(
            pady=(25, 5)
        )

        tk.Spinbox(
            dialog,
            from_=1,
            to=30,
            textvariable=rows,
            width=10
        ).pack()

        tk.Label(
            dialog,
            text="Столбцов:",
            bg=self.theme["ribbon"],
            fg=self.theme["text"]
        ).pack(
            pady=(15, 5)
        )

        tk.Spinbox(
            dialog,
            from_=1,
            to=15,
            textvariable=cols,
            width=10
        ).pack()

        def create():
            try:
                r = max(1, rows.get())
                c = max(1, cols.get())

                table = TableWidget(
                    self,
                    r,
                    c
                )

                self.tables.append(
                    table
                )

                self.clear_object_selection()

                table.select()

                dialog.destroy()

            except:
                messagebox.showerror(
                    "Ошибка",
                    "Не удалось создать таблицу."
                )

        RoundedButton(
            dialog,
            "Создать таблицу",
            create,
            self.theme
        ).pack(
            pady=25
        )

    def selected_table(self):
        for table in self.tables:
            if table.selected:
                return table
        return None

    def table_row_add(self):
        table = self.selected_table()

        if table:
            table.add_row()
        else:
            self.status.set(
                "Выберите таблицу"
            )

    def table_col_add(self):
        table = self.selected_table()

        if table:
            table.add_col()
        else:
            self.status.set(
                "Выберите таблицу"
            )

    def table_row_delete(self):
        table = self.selected_table()

        if table:
            table.delete_row()

    def table_col_delete(self):
        table = self.selected_table()

        if table:
            table.delete_col()

    # ==========================================================
    # IMAGES
    # ==========================================================

    def insert_image(self):
        filename = filedialog.askopenfilename(
            title="Вставить изображение",
            filetypes=[
                (
                    "Изображения",
                    "*.png *.jpg *.jpeg *.gif *.bmp *.webp"
                ),
                (
                    "Все файлы",
                    "*.*"
                )
            ]
        )

        if not filename:
            return

        try:
            image = Image.open(
                filename
            )

            width = min(
                image.width,
                450
            )

            obj = ImageObject(
                self,
                filename,
                100,
                150,
                width
            )

            self.images.append(
                obj
            )

            self.clear_object_selection()
            obj.select()

        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                str(e)
            )

    def selected_image(self):
        for image in self.images:
            if image.selected:
                return image
        return None

    def resize_image(self):
        image = self.selected_image()

        if image:
            image.resize_dialog()
        else:
            self.status.set(
                "Выберите изображение"
            )

    def clear_object_selection(self):
        for image in self.images:
            image.selected = False

            try:
                image.frame.configure(
                    highlightthickness=1,
                    highlightbackground="#bbbbbb"
                )
            except:
                pass

        for table in self.tables:
            table.selected = False

            try:
                table.frame.configure(
                    highlightthickness=1,
                    highlightbackground="#aab4c3"
                )
            except:
                pass

    def delete_object(self):
        image = self.selected_image()

        if image:
            image.destroy()
            self.images.remove(image)
            return

        table = self.selected_table()

        if table:
            table.destroy()
            self.tables.remove(table)
            return

        self.status.set(
            "Выберите изображение или таблицу"
        )

    # ==========================================================
    # INSERT TOOLS
    # ==========================================================

    def page_break(self):
        self.editor.insert(
            "insert",
            "\n\n"
        )

        self.editor.insert(
            "insert",
            "────────────────────────────\n"
        )

        self.editor.insert(
            "insert",
            "\n\n"
        )

    def page_number(self):
        self.editor.insert(
            "insert",
            "[Страница 1]"
        )

    def insert_date(self):
        self.editor.insert(
            "insert",
            datetime.now().strftime(
                "%d.%m.%Y %H:%M"
            )
        )

    def insert_symbol(self):
        symbol = simpledialog.askstring(
            "Символ",
            "Введите символ:\n"
            "© ® ™ § ± × ÷ → ← ↑ ↓ ★ ☆ ✓ ♥"
        )

        if symbol:
            self.editor.insert(
                "insert",
                symbol
            )

    def insert_line(self):
        self.editor.insert(
            "insert",
            "\n" + "─" * 70 + "\n"
        )

    def insert_link(self):
        url = simpledialog.askstring(
            "Гиперссылка",
            "Введите адрес:"
        )

        if not url:
            return

        start, end = self.selection()

        if start:
            tag = "link_" + str(abs(hash(url)))

            self.editor.tag_configure(
                tag,
                foreground="#0563c1",
                underline=True
            )

            self.editor.tag_add(
                tag,
                start,
                end
            )

            self.editor.tag_bind(
                tag,
                "<Button-1>",
                lambda e, u=url: self.open_url(u)
            )

        else:
            self.editor.insert(
                "insert",
                url
            )

    def open_url(self, url):
        import webbrowser
        webbrowser.open(url)

    # ==========================================================
    # LAYOUT
    # ==========================================================

    def orientation(self, landscape):
        if landscape:
            self.page_width = 1123
            self.page_height = 794
        else:
            self.page_width = 794
            self.page_height = 1123

        self.page.configure(
            width=self.page_width,
            height=self.page_height
        )

    def a4(self):
        self.page_width = 794
        self.page_height = 1123

        self.page.configure(
            width=self.page_width,
            height=self.page_height
        )

    def margins(self):
        value = simpledialog.askinteger(
            "Поля",
            "Поля страницы в пикселях:",
            initialvalue=75,
            minvalue=20,
            maxvalue=200
        )

        if value:
            children = self.page.winfo_children()

            if children:
                children[0].configure(
                    padx=value,
                    pady=value
                )

    def indent(self):
        start, end = self.selection()

        if not start:
            start = self.editor.index(
                "insert linestart"
            )
            end = self.editor.index(
                "insert lineend"
            )

        tag = "indent_" + str(
            abs(hash((start, end)))
        )

        self.editor.tag_configure(
            tag,
            lmargin1=40,
            lmargin2=40
        )

        self.editor.tag_add(
            tag,
            start,
            end
        )

    def outdent(self):
        start, end = self.selection()

        if not start:
            start = self.editor.index(
                "insert linestart"
            )
            end = self.editor.index(
                "insert lineend"
            )

        tag = "outdent_" + str(
            abs(hash((start, end)))
        )

        self.editor.tag_configure(
            tag,
            lmargin1=0,
            lmargin2=0
        )

        self.editor.tag_add(
            tag,
            start,
            end
        )

    # ==========================================================
    # DESIGN
    # ==========================================================

    def page_color(self):
        color = colorchooser.askcolor(
            title="Цвет страницы"
        )[1]

        if not color:
            return

        self.page.configure(
            bg=color
        )

        for child in self.page.winfo_children():
            try:
                child.configure(
                    bg=color
                )
            except:
                pass

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode

        self.theme = (
            DARK
            if self.dark_mode
            else LIGHT
        )

        current_text = ""

        if hasattr(self, "editor"):
            current_text = self.editor.get(
                "1.0",
                "end-1c"
            )

        self.create_editor(
            current_text
        )

    # ==========================================================
    # REFERENCES
    # ==========================================================

    def table_of_contents(self):
        text = self.editor.get(
            "1.0",
            "end-1c"
        )

        lines = text.splitlines()

        headings = []

        for line in lines:
            if line.strip():
                if (
                    len(line) < 100
                    and not line.startswith("•")
                    and not re.match(
                        r"^\d+\.",
                        line
                    )
                ):
                    headings.append(
                        line.strip()
                    )

        if not headings:
            messagebox.showinfo(
                "Оглавление",
                "Сначала создайте заголовки."
            )
            return

        toc = "ОГЛАВЛЕНИЕ\n\n"

        for i, heading in enumerate(
            headings,
            1
        ):
            toc += (
                f"{i}. {heading}"
                f" ................................ {i}\n"
            )

        self.editor.insert(
            "1.0",
            toc + "\n"
        )

    def update_toc(self):
        messagebox.showinfo(
            "Оглавление",
            "Оглавление обновлено."
        )

    def bookmark(self):
        name = simpledialog.askstring(
            "Закладка",
            "Название закладки:"
        )

        if name:
            self.status.set(
                f"Закладка создана: {name}"
            )

    # ==========================================================
    # REVIEW
    # ==========================================================

    def protect_document(self):
        password = simpledialog.askstring(
            "Защита",
            "Введите пароль:",
            show="*"
        )

        if password:
            self.status.set(
                "Защита документа включена."
            )

    # ==========================================================
    # SEARCH
    # ==========================================================

    def find_text(self):
        word = simpledialog.askstring(
            "Поиск",
            "Найти:"
        )

        if not word:
            return

        self.editor.tag_remove(
            "search",
            "1.0",
            "end"
        )

        start = "1.0"
        count = 0

        while True:
            pos = self.editor.search(
                word,
                start,
                stopindex="end"
            )

            if not pos:
                break

            end = f"{pos}+{len(word)}c"

            self.editor.tag_add(
                "search",
                pos,
                end
            )

            start = end
            count += 1

        self.editor.tag_configure(
            "search",
            background="#ffe16b",
            foreground="#111111"
        )

        self.status.set(
            f"Найдено: {count}"
        )

    def replace_text(self):
        window = tk.Toplevel(
            self.root
        )

        window.title(
            "Найти и заменить"
        )

        window.geometry(
            "420x260"
        )

        window.configure(
            bg=self.theme["ribbon"]
        )

        tk.Label(
            window,
            text="Найти",
            bg=self.theme["ribbon"],
            fg=self.theme["text"]
        ).pack(
            anchor="w",
            padx=25,
            pady=(25, 5)
        )

        find_entry = tk.Entry(
            window,
            bg=self.theme["button"],
            fg=self.theme["text"],
            insertbackground=self.theme["text"],
            relief="flat"
        )

        find_entry.pack(
            fill="x",
            padx=25
        )

        tk.Label(
            window,
            text="Заменить на",
            bg=self.theme["ribbon"],
            fg=self.theme["text"]
        ).pack(
            anchor="w",
            padx=25,
            pady=(15, 5)
        )

        replace_entry = tk.Entry(
            window,
            bg=self.theme["button"],
            fg=self.theme["text"],
            insertbackground=self.theme["text"],
            relief="flat"
        )

        replace_entry.pack(
            fill="x",
            padx=25
        )

        def execute():
            old = find_entry.get()
            new = replace_entry.get()

            if not old:
                return

            text = self.editor.get(
                "1.0",
                "end-1c"
            )

            count = text.count(
                old
            )

            self.editor.delete(
                "1.0",
                "end"
            )

            self.editor.insert(
                "1.0",
                text.replace(
                    old,
                    new
                )
            )

            window.destroy()

            messagebox.showinfo(
                "Готово",
                f"Заменено: {count}"
            )

        RoundedButton(
            window,
            "Заменить всё",
            execute,
            self.theme
        ).pack(
            pady=25
        )

    # ==========================================================
    # STATISTICS
    # ==========================================================

    def statistics(self):
        text = self.editor.get(
            "1.0",
            "end-1c"
        )

        words = re.findall(
            r"\S+",
            text
        )

        letters = len(
            re.sub(
                r"\s",
                "",
                text
            )
        )

        paragraphs = len([
            x for x in text.split("\n")
            if x.strip()
        ])

        messagebox.showinfo(
            "Статистика",
            f"Слов: {len(words)}\n"
            f"Символов: {len(text)}\n"
            f"Без пробелов: {letters}\n"
            f"Абзацев: {paragraphs}\n"
            f"Таблиц: {len(self.tables)}\n"
            f"Изображений: {len(self.images)}"
        )

    # ==========================================================
    # ZOOM
    # ==========================================================

    def zoom_in(self):
        self.zoom = min(
            2.0,
            self.zoom + 0.1
        )

        self.apply_zoom()

    def zoom_out(self):
        self.zoom = max(
            0.6,
            self.zoom - 0.1
        )

        self.apply_zoom()

    def zoom_reset(self):
        self.zoom = 1.0
        self.apply_zoom()

    def apply_zoom(self):
        self.zoom_text.set(
            f"{int(self.zoom * 100)}%"
        )

        self.page.configure(
            width=int(
                self.page_width *
                self.zoom
            ),
            height=int(
                self.page_height *
                self.zoom
            )
        )

    # ==========================================================
    # NAVIGATION
    # ==========================================================

    def navigation(self):
        if hasattr(self, "navigation_panel"):
            try:
                self.navigation_panel.destroy()
                del self.navigation_panel
                return
            except:
                pass

        self.navigation_panel = tk.Frame(
            self.root,
            bg=self.theme["ribbon"],
            width=250
        )

        self.navigation_panel.place(
            x=0,
            y=186,
            relheight=1
        )

        tk.Label(
            self.navigation_panel,
            text="Навигация",
            bg=self.theme["ribbon"],
            fg=self.theme["text"],
            font=("Segoe UI", 14, "bold")
        ).pack(
            anchor="w",
            padx=20,
            pady=20
        )

        text = self.editor.get(
            "1.0",
            "end-1c"
        )

        for line in text.splitlines():
            if line.strip():
                if len(line) < 80:
                    button = tk.Button(
                        self.navigation_panel,
                        text=line[:30],
                        anchor="w",
                        bg=self.theme["ribbon"],
                        fg=self.theme["text"],
                        activebackground=self.theme["button"],
                        relief="flat",
                        bd=0,
                        cursor="hand2"
                    )

                    button.pack(
                        fill="x",
                        padx=10,
                        pady=2
                    )

    # ==========================================================
    # FILE SAVE
    # ==========================================================

    def save_document(self):
        if not self.current_file:
            return self.save_as()

        self.export_docx(
            self.current_file
        )

    def save_as(self):
        filename = filedialog.asksaveasfilename(
            title="Сохранить",
            defaultextension=".docx",
            filetypes=[
                ("Word document", "*.docx"),
                ("Все файлы", "*.*")
            ]
        )

        if not filename:
            return

        self.current_file = filename

        self.document_title.set(
            os.path.basename(
                filename
            )
        )

        self.export_docx(
            filename
        )

    def export_docx(self, filename):
        try:
            doc = Document()

            section = doc.sections[0]

            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)

            text = self.editor.get(
                "1.0",
                "end-1c"
            )

            for line in text.split("\n"):
                paragraph = doc.add_paragraph()

                run = paragraph.add_run(
                    line
                )

                run.font.name = "Calibri"
                run.font.size = Pt(12)

            for table in self.tables:
                data = table.current_data()

                doc_table = doc.add_table(
                    rows=len(data),
                    cols=len(data[0])
                )

                doc_table.style = "Table Grid"

                for r in range(len(data)):
                    for c in range(len(data[r])):
                        doc_table.cell(
                            r,
                            c
                        ).text = data[r][c]

            for image in self.images:
                try:
                    doc.add_picture(
                        image.filename,
                        width=Cm(10)
                    )
                except:
                    pass

            doc.save(
                filename
            )

            self.modified = False

            self.document_title.set(
                os.path.basename(
                    filename
                )
            )

            self.status.set(
                f"Сохранено: {os.path.basename(filename)}"
            )

        except Exception as e:
            messagebox.showerror(
                "Ошибка сохранения",
                str(e)
            )

    # ==========================================================
    # OPEN
    # ==========================================================

    def open_document_from_path(self, filename):
        try:
            doc = Document(str(filename))
            text = "\n".join(p.text for p in doc.paragraphs)
            self.create_editor(text)
            self.current_file = str(filename)
            self.document_title.set(os.path.basename(str(filename)))
            self.status.set(f"Открыт: {os.path.basename(str(filename))}")
        except Exception as e:
            messagebox.showerror("Ошибка открытия", str(e), parent=self.root)

    def open_document(self):
        filename = filedialog.askopenfilename(
            title="Открыть",
            filetypes=[
                ("Word document", "*.docx"),
                ("Все файлы", "*.*")
            ]
        )

        if not filename:
            return

        try:
            doc = Document(
                filename
            )

            text = "\n".join(
                p.text
                for p in doc.paragraphs
            )

            self.create_editor(
                text
            )

            self.current_file = filename

            self.document_title.set(
                os.path.basename(
                    filename
                )
            )

            self.status.set(
                f"Открыт: {os.path.basename(filename)}"
            )

        except Exception as e:
            messagebox.showerror(
                "Ошибка открытия",
                str(e)
            )

    # ==========================================================
    # PDF
    # ==========================================================

    def export_pdf(self):
        filename = filedialog.asksaveasfilename(
            title="Экспорт PDF",
            defaultextension=".pdf",
            filetypes=[
                ("PDF", "*.pdf")
            ]
        )

        if not filename:
            return

        try:
            c = pdf_canvas.Canvas(
                filename,
                pagesize=A4
            )

            width, height = A4

            x = 55
            y = height - 55

            text = self.editor.get(
                "1.0",
                "end-1c"
            )

            for line in text.splitlines():
                if y < 50:
                    c.showPage()
                    y = height - 55

                chunks = [
                    line[i:i + 100]
                    for i in range(
                        0,
                        len(line),
                        100
                    )
                ]

                if not chunks:
                    chunks = [""]

                for chunk in chunks:
                    if y < 50:
                        c.showPage()
                        y = height - 55

                    c.setFont(
                        "Helvetica",
                        10
                    )

                    c.drawString(
                        x,
                        y,
                        chunk
                    )

                    y -= 15

            c.save()

            messagebox.showinfo(
                "PDF",
                "PDF успешно создан."
            )

        except Exception as e:
            messagebox.showerror(
                "Ошибка PDF",
                str(e)
            )

    # ==========================================================
    # PRINT
    # ==========================================================

    def print_document(self):
        if not self.current_file:
            self.save_as()

        if not self.current_file:
            return

        try:
            if os.name == "nt":
                os.startfile(
                    self.current_file,
                    "print"
                )
            else:
                subprocess.run(
                    ["lp", self.current_file]
                )

        except Exception as e:
            messagebox.showerror(
                "Печать",
                str(e)
            )

    # ==========================================================
    # SHORTCUTS
    # ==========================================================

    def bind_shortcuts(self):
        # Bind shortcuts directly to the editor. EmbeddedRoot is only a host Frame,
        # so bindings on the parent do not receive key events from the Text widget.
        target=self.editor
        target.bind("<Control-n>", lambda e:(self.create_editor(""),"break")[1])
        target.bind("<Control-o>", lambda e:(self.open_document(),"break")[1])
        target.bind("<Control-s>", lambda e:(self.save_document(),"break")[1])
        target.bind("<Control-S>", lambda e:(self.save_document(),"break")[1])
        target.bind("<Control-z>", lambda e:(self.undo(),"break")[1])
        target.bind("<Control-y>", lambda e:(self.redo(),"break")[1])
        target.bind("<Control-b>", lambda e:(self.bold(),"break")[1])
        target.bind("<Control-i>", lambda e:(self.italic(),"break")[1])
        target.bind("<Control-u>", lambda e:(self.underline(),"break")[1])
        target.bind("<Control-f>", lambda e:(self.find_text(),"break")[1])
        target.bind("<Control-h>", lambda e:(self.replace_text(),"break")[1])
        target.bind("<Control-a>", lambda e:(self.select_all(),"break")[1])
        target.bind("<Control-c>", lambda e:(self.copy(),"break")[1])
        target.bind("<Control-x>", lambda e:(self.cut(),"break")[1])
        target.bind("<Control-v>", lambda e:(self.paste(),"break")[1])
        target.bind("<F11>", lambda e:(self.fullscreen(),"break")[1])

    def select_all(self):
        self.editor.tag_add(
            "sel",
            "1.0",
            "end"
        )

    # ==========================================================
    # AUTOSAVE
    # ==========================================================

    def modified_event(self, event=None):
        try:
            if self.editor.edit_modified():
                self.modified = True

                self.editor.edit_modified(
                    False
                )

        except:
            pass

    def autosave(self):
        try:
            if (
                self.modified
                and self.current_file
            ):
                self.export_docx(
                    self.current_file
                )
        except:
            pass

        self.root.after(
            AUTOSAVE_SECONDS,
            self.autosave
        )

    # ==========================================================
    # FULLSCREEN
    # ==========================================================

    def fullscreen(self):
        current = self.root.attributes(
            "-fullscreen"
        )

        self.root.attributes(
            "-fullscreen",
            not current
        )

    # ==========================================================
    # CLEANUP
    # ==========================================================

    def clear(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        self.images = []
        self.tables = []

    def close(self):
        if (
            self.modified
            and hasattr(self, "editor")
        ):
            result = messagebox.askyesnocancel(
                "Word Pro 3.0",
                "Документ изменён.\n"
                "Сохранить перед выходом?"
            )

            if result is None:
                return

            if result:
                self.save_document()

        self.root.destroy()


# ==============================================================
# START
# ==============================================================

if __name__ == "__main__":
    root = tk.Tk()

    app = WordPro(
        root
    )

    root.mainloop()