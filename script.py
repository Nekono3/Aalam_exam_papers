import io
import math
import os
import random
from collections import defaultdict, Counter

import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog

from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.colors import grey
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


DEFAULT_LOGO = "school_logo.png"
DEFAULT_EXCEL = "students.xlsx"
DEFAULT_TEMPLATE = "ZipGrade_AnswerSheet_AALAM-1.pdf"
DEFAULT_OUTPUT = "Exam_list_zipgrade.pdf"

START_X = 154.5
START_Y = 645.5
COL_SPACING = 15.45
ROW_SPACING = 17.4
BUBBLE_RADIUS = 4.5

NAME_X = 149
NAME_Y = 696
CLASS_X = 330
CLASS_Y = 696

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FONT_NAME = "DejaVu"
FONT_PATH = os.path.join(BASE_DIR, "DejaVuSans.ttf")

if not os.path.exists(FONT_PATH):
    raise FileNotFoundError(
        f"Font file not found: {FONT_PATH}\n"
        "Put DejaVuSans.ttf in the same folder as this script."
    )

pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    original_cols = list(df.columns)
    cols = {str(c).strip().lower(): c for c in df.columns}

    mapping_candidates = {
        "id": ["id", "student_id", "ogrenciid", "oquuchu_id", "ученик_id"],
        "name": ["name", "ad", "first_name", "isim", "аты"],
        "surname": ["surname", "soyad", "last_name", "familiya", "фамилия"],
        "class": ["class", "sinif", "grade", "klass", "класс"],
        "section": ["section", "sube", "group", "буква", "секция"],
        "class_type": ["class_type", "type", "til", "lang", "язык", "stream"],
    }

    resolved = {}
    for target, candidates in mapping_candidates.items():
        for candidate in candidates:
            if candidate in cols:
                resolved[target] = cols[candidate]
                break

    required = ["id", "name", "surname", "class", "section"]
    missing = [c for c in required if c not in resolved]
    if missing:
        raise ValueError(
            "Excel format is incorrect.\n"
            f"Found columns: {original_cols}\n"
            f"Missing required fields: {missing}"
        )

    out = pd.DataFrame()
    for key, source in resolved.items():
        out[key] = df[source]

    if "class_type" not in out.columns:
        out["class_type"] = ""

    return out.fillna("")


def clean_student_id(raw_id):
    if pd.isna(raw_id) or str(raw_id).strip() == "":
        return "000000000"
    text = str(raw_id).strip()
    if text.endswith(".0"):
        text = text[:-2]
    digits_only = "".join(ch for ch in text if ch.isdigit())
    if not digits_only:
        return "000000000"
    return digits_only.zfill(9)


def format_full_name(row):
    return f"{str(row.get('surname', '')).strip()} {str(row.get('name', '')).strip()}".strip()


def format_class_text(row):
    c = str(row.get("class", "")).strip()
    s = str(row.get("section", "")).strip()
    ct = str(row.get("class_type", "")).strip()
    class_parts = f"{c} {s}".strip()
    return f"{class_parts} | {ct}" if ct and class_parts else (ct or class_parts)


def section_key_from_row(row):
    c = str(row.get("class", "")).strip()
    s = str(row.get("section", "")).strip()
    ct = str(row.get("class_type", "")).strip()
    base = f"{c}-{s}" if c or s else "UNKNOWN"
    return f"{base} | {ct}" if ct else base


def draw_logo(can, w, h, width=105, height=105, x=40, y_offset=88):
    logo_path = os.path.join(BASE_DIR, DEFAULT_LOGO)
    if os.path.exists(logo_path):
        can.drawImage(
            logo_path,
            x,
            h - y_offset,
            width=width,
            height=height,
            preserveAspectRatio=True,
            mask='auto'
        )


def create_overlay(fullname, class_text, student_id):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    can.setFont(FONT_NAME, 12)
    can.drawString(NAME_X, NAME_Y, str(fullname))
    can.drawString(CLASS_X, CLASS_Y, str(class_text))

    id_str = clean_student_id(student_id)

    can.setFont(FONT_NAME, 10)
    digit_y = START_Y + 14
    for col, digit in enumerate(id_str):
        x = START_X + (col * COL_SPACING)
        text_x = x - 3
        can.drawString(text_x, digit_y, str(digit))

    for col, digit in enumerate(id_str):
        if not str(digit).isdigit():
            continue
        row = int(digit)
        x = START_X + (col * COL_SPACING)
        y = START_Y - (row * ROW_SPACING)
        can.circle(x, y, BUBBLE_RADIUS, fill=1)

    can.save()
    packet.seek(0)
    return PdfReader(packet)


def distribute_students(df, room_names, capacity=None, shuffle=True):
    df = df.copy()

    if shuffle:
        df = df.sample(frac=1, random_state=random.randint(1, 10_000)).reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)

    df["_group_size"] = df.groupby(["class", "section", "class_type"])["class"].transform("size")
    df = df.sort_values("_group_size", ascending=False).reset_index(drop=True)

    total_students = len(df)
    room_count = len(room_names)

    if room_count <= 0:
        raise ValueError("Room count must be greater than 0.")

    if capacity is None or capacity <= 0:
        capacity = math.ceil(total_students / room_count)

    if room_count * capacity < total_students:
        raise ValueError(
            f"Not enough total capacity.\n"
            f"Students: {total_students}\nRooms: {room_count}\nCapacity per room: {capacity}\n"
            f"Total seats: {room_count * capacity}"
        )

    rooms = {room: [] for room in room_names}
    room_class_count = {room: defaultdict(int) for room in room_names}
    room_grade_count = {room: defaultdict(int) for room in room_names}

    for _, row in df.iterrows():
        student = row.to_dict()
        grade = str(student.get("class", "")).strip()
        section = str(student.get("section", "")).strip()
        class_type = str(student.get("class_type", "")).strip()

        class_key = f"{grade}-{section}|{class_type}"
        grade_key = grade

        possible_rooms = [room for room in room_names if len(rooms[room]) < capacity]
        if not possible_rooms:
            raise RuntimeError("Could not place a student even though capacity should be enough.")

        def room_score(room):
            return (
                room_class_count[room][class_key],
                room_grade_count[room][grade_key],
                len(rooms[room]),
                random.random(),
            )

        best_room = min(possible_rooms, key=room_score)
        rooms[best_room].append(student)
        room_class_count[best_room][class_key] += 1
        room_grade_count[best_room][grade_key] += 1

    return rooms, capacity


def _draw_table_header(c, y):
    c.setFont(FONT_NAME, 11)
    c.drawString(40, y, "No")
    c.drawString(95, y, "Student ID")
    c.drawString(205, y, "Full Name")
    c.drawString(505, y, "Class")
    y -= 10
    c.line(40, y, 555, y)
    y -= 18
    return y


def build_room_packet_list_pages(room_name, students, title="LIST OF STUDENTS", rows_per_page=25):
    pages = []
    for page_start in range(0, len(students), rows_per_page):
        chunk = students[page_start:page_start + rows_per_page]
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        w, h = A4

        draw_logo(c, w, h, width=105, height=105, x=40, y_offset=88)

        y = h - 50
        c.setFont(FONT_NAME, 18)
        c.drawCentredString(w / 2, y, title)
        y -= 28

        c.setFont(FONT_NAME, 14)
        c.drawCentredString(w / 2, y, f"Room: {room_name}")
        y -= 26

        y = _draw_table_header(c, y)

        for idx, student in enumerate(chunk, start=page_start + 1):
            c.drawString(40, y, str(idx))
            c.drawString(95, y, clean_student_id(student.get("id", "")))
            c.drawString(205, y, format_full_name(student))
            c.drawString(505, y, format_class_text(student))
            y -= 18

        c.save()
        buffer.seek(0)
        reader = PdfReader(buffer)
        pages.extend(reader.pages)
    return pages


def build_admin_room_pages(rooms, title="EXAM ROOM LIST", rows_per_page=25):
    pages = []
    for room_name, students in rooms.items():
        summary = Counter(section_key_from_row(s) for s in students)

        for page_start in range(0, len(students), rows_per_page):
            chunk = students[page_start:page_start + rows_per_page]
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            w, h = A4

            draw_logo(c, w, h, width=95, height=95, x=35, y_offset=95)

            y = h - 50
            c.setFont(FONT_NAME, 18)
            c.drawCentredString(w / 2, y, title)
            y -= 28

            c.setFont(FONT_NAME, 14)
            room_title = f"Room: {room_name}" if page_start == 0 else f"Room: {room_name} (continued)"
            c.drawCentredString(w / 2, y, room_title)
            y -= 26

            y = _draw_table_header(c, y)

            for idx, student in enumerate(chunk, start=page_start + 1):
                c.drawString(40, y, str(idx))
                c.drawString(95, y, clean_student_id(student.get("id", "")))
                c.drawString(205, y, format_full_name(student))
                c.drawString(505, y, format_class_text(student))
                y -= 18

            if page_start + rows_per_page >= len(students):
                y -= 6
                c.line(40, y, 555, y)
                y -= 20

                c.setFont(FONT_NAME, 13)
                c.drawString(40, y, "Summary:")
                y -= 18

                c.setFont(FONT_NAME, 10)
                c.drawString(50, y, f"Total students: {len(students)}")
                y -= 16
                for key, value in sorted(summary.items()):
                    if y < 60:
                        break
                    c.drawString(50, y, f"{key} -> {value} students")
                    y -= 14

            c.save()
            buffer.seek(0)
            reader = PdfReader(buffer)
            pages.extend(reader.pages)
    return pages


def generate_all(excel_path, template_pdf_path, output_pdf_path, room_names, capacity=None, shuffle=True):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    if not os.path.exists(template_pdf_path):
        raise FileNotFoundError(f"Template PDF not found: {template_pdf_path}")

    df_raw = pd.read_excel(excel_path)
    df = normalize_columns(df_raw)
    df = df.dropna(subset=["class", "section"]).fillna("")

    if len(df) == 0:
        raise ValueError("No students found in the Excel file.")

    rooms, actual_capacity = distribute_students(
        df=df,
        room_names=room_names,
        capacity=capacity,
        shuffle=shuffle
    )

    writer = PdfWriter()
    total_students = len(df)
    current_index = 1

    # Main packet uses distribution settings
    for room_name, students in rooms.items():
        room_list_pages = build_room_packet_list_pages(room_name, students, title="LIST OF STUDENTS")
        for p in room_list_pages:
            writer.add_page(p)

        for student in students:
            reader = PdfReader(template_pdf_path)
            page = reader.pages[0]

            overlay = create_overlay(
                format_full_name(student),
                format_class_text(student),
                student.get("id", "")
            )
            page.merge_page(overlay.pages[0])
            writer.add_page(page)

            print(f"[{current_index}/{total_students}] Answer sheet prepared for {format_full_name(student)}")
            current_index += 1

    # Admin papers at the end, also based on distribution settings
    admin_pages = build_admin_room_pages(rooms, title="EXAM ROOM LIST")
    for p in admin_pages:
        writer.add_page(p)

    with open(output_pdf_path, "wb") as f:
        writer.write(f)

    return {
        "students": len(df),
        "rooms": len(room_names),
        "capacity": actual_capacity,
        "output": output_pdf_path,
        "distribution": {room: len(students) for room, students in rooms.items()},
    }


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("ZipGrade + Room Distribution")
        self.root.geometry("980x760")
        self.root.minsize(980, 760)

        self.excel_var = tk.StringVar(value=DEFAULT_EXCEL)
        self.template_var = tk.StringVar(value=DEFAULT_TEMPLATE)
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)

        self.room_count_var = tk.StringVar(value="5")
        self.capacity_var = tk.StringVar(value="")
        self.room_names_var = tk.StringVar(value="10A,10B,9A,9B,8A")
        self.shuffle_var = tk.BooleanVar(value=True)

        self.setup_theme()
        self.center_window()
        self.build_ui()

    def setup_theme(self):
        self.colors = {
            "bg": "#0f172a",
            "panel": "#111827",
            "border": "#334155",
            "text": "#e5e7eb",
            "muted": "#94a3b8",
            "accent": "#2563eb",
            "accent_hover": "#1d4ed8",
            "button2": "#374151",
            "button2_hover": "#4b5563",
            "entry_bg": "#475569",
            "log_bg": "#020617",
            "log_fg": "#d1fae5",
        }
        self.root.configure(bg=self.colors["bg"])

    def center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def create_card(self, parent, title):
        outer = tk.Frame(parent, bg=self.colors["border"], bd=0, highlightthickness=0)
        outer.pack(fill="x", pady=(0, 14))
        inner = tk.Frame(outer, bg=self.colors["panel"], padx=14, pady=14)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text=title, font=("Arial", 11, "bold"), bg=self.colors["panel"], fg=self.colors["text"]).pack(anchor="w", pady=(0, 10))
        body = tk.Frame(inner, bg=self.colors["panel"])
        body.pack(fill="both", expand=True)
        return body

    def styled_label(self, parent, text, width=None):
        return tk.Label(parent, text=text, font=("Arial", 10), bg=self.colors["panel"], fg=self.colors["text"], width=width, anchor="w")

    def styled_button(self, parent, text, command, primary=False, width=None):
        bg = self.colors["accent"] if primary else self.colors["button2"]
        hover = self.colors["accent_hover"] if primary else self.colors["button2_hover"]
        btn = tk.Button(parent, text=text, command=command, font=("Arial", 10, "bold"), bg=bg, fg="#ffffff",
                        activebackground=hover, activeforeground="#ffffff", relief="flat", bd=0,
                        cursor="hand2", padx=14, pady=8, width=width)
        btn.bind("<Enter>", lambda e: btn.config(bg=hover))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    def styled_checkbutton(self, parent, text, variable):
        return tk.Checkbutton(parent, text=text, variable=variable, font=("Arial", 10), bg=self.colors["panel"],
                              fg=self.colors["text"], activebackground=self.colors["panel"],
                              activeforeground=self.colors["text"], selectcolor=self.colors["entry_bg"],
                              highlightthickness=0, bd=0)

    def add_entry_with_border(self, parent, textvariable=None, width=None, side="left", fill=None, expand=False, padx=(0, 0), ipady=6):
        wrapper = tk.Frame(parent, bg=self.colors["border"], bd=0, highlightthickness=0)
        wrapper.pack(side=side, fill=fill, expand=expand, padx=padx)
        entry = tk.Entry(wrapper, textvariable=textvariable, font=("Arial", 10), bg=self.colors["entry_bg"],
                         fg="#ffffff", insertbackground="#ffffff", disabledbackground=self.colors["entry_bg"],
                         disabledforeground="#ffffff", relief="flat", bd=0, width=width)
        entry.pack(fill="both", expand=True, padx=1, pady=1, ipady=ipady)
        return entry

    def _modern_file_row(self, parent, label, var, command, button_text="Browse"):
        row = tk.Frame(parent, bg=self.colors["panel"])
        row.pack(fill="x", pady=6)
        self.styled_label(row, label, width=14).pack(side="left", padx=(0, 10))
        self.add_entry_with_border(row, textvariable=var, side="left", fill="x", expand=True, padx=(0, 10), ipady=7)
        self.styled_button(row, text=button_text, command=command, primary=False, width=10).pack(side="left")

    def build_ui(self):
        container = tk.Frame(self.root, bg=self.colors["bg"], padx=20, pady=18)
        container.pack(fill="both", expand=True)

        header = tk.Frame(container, bg=self.colors["bg"])
        header.pack(fill="x", pady=(0, 16))
        tk.Label(header, text="Combined Generator", font=("Arial", 22, "bold"), bg=self.colors["bg"], fg=self.colors["text"]).pack(anchor="w")
        tk.Label(header, text="Distribution settings now control packet room grouping and admin papers",
                 font=("Arial", 10), bg=self.colors["bg"], fg=self.colors["muted"]).pack(anchor="w", pady=(4, 0))

        files_body = self.create_card(container, "Files")
        self._modern_file_row(files_body, "Students Excel", self.excel_var, self.pick_excel, "Browse")
        self._modern_file_row(files_body, "Template PDF", self.template_var, self.pick_template, "Browse")
        self._modern_file_row(files_body, "Output PDF", self.output_var, self.pick_output, "Save As")

        dist_body = self.create_card(container, "Distribution Settings")
        top_row = tk.Frame(dist_body, bg=self.colors["panel"])
        top_row.pack(fill="x", pady=(0, 12))

        self.styled_label(top_row, "Room count", width=16).pack(side="left", padx=(0, 8))
        self.add_entry_with_border(top_row, textvariable=self.room_count_var, width=10, side="left", padx=(0, 24), ipady=6)

        self.styled_label(top_row, "Capacity per room", width=18).pack(side="left", padx=(0, 8))
        self.add_entry_with_border(top_row, textvariable=self.capacity_var, width=10, side="left", padx=(0, 10), ipady=6)

        tk.Label(top_row, text="Leave empty for auto distribution", font=("Arial", 9), bg=self.colors["panel"], fg=self.colors["muted"]).pack(side="left")

        room_row = tk.Frame(dist_body, bg=self.colors["panel"])
        room_row.pack(fill="x", pady=(0, 12))
        self.styled_label(room_row, "Room names (comma-separated)").pack(anchor="w", pady=(0, 6))
        self.add_entry_with_border(room_row, textvariable=self.room_names_var, side="top", fill="x", expand=True, ipady=7)

        opts_row = tk.Frame(dist_body, bg=self.colors["panel"])
        opts_row.pack(fill="x")
        self.styled_checkbutton(opts_row, "Shuffle students before distribution", self.shuffle_var).pack(anchor="w")

        notes_body = self.create_card(container, "Notes")
        notes = (
            "• Distribution settings are now used in the final packet\n"
            "• For each room: room student list page -> answer sheets for that room\n"
            "• Admin pages at the end use the same room distribution\n"
            "• Admin pages include summary and total student count\n"
            "• School logo is on room packet pages and admin pages"
        )
        tk.Label(notes_body, text=notes, justify="left", font=("Arial", 10), bg=self.colors["panel"], fg=self.colors["muted"]).pack(anchor="w")

        btn_area = tk.Frame(container, bg=self.colors["bg"])
        btn_area.pack(fill="x", pady=(6, 16))
        self.styled_button(btn_area, text="GENERATE COMBINED PDF", command=self.run, primary=True, width=28).pack()

        result_outer = tk.Frame(container, bg=self.colors["border"])
        result_outer.pack(fill="both", expand=True)
        result_inner = tk.Frame(result_outer, bg=self.colors["panel"], padx=14, pady=14)
        result_inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(result_inner, text="Result", font=("Arial", 11, "bold"), bg=self.colors["panel"], fg=self.colors["text"]).pack(anchor="w", pady=(0, 10))

        log_border = tk.Frame(result_inner, bg=self.colors["border"])
        log_border.pack(fill="both", expand=True)
        self.log_box = tk.Text(log_border, height=14, wrap="word", font=("Consolas", 10),
                               bg=self.colors["log_bg"], fg=self.colors["log_fg"],
                               insertbackground=self.colors["log_fg"], relief="flat", bd=0, padx=10, pady=10)
        self.log_box.pack(fill="both", expand=True, padx=1, pady=1)

    def pick_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_var.set(path)

    def pick_template(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.template_var.set(path)

    def pick_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path:
            self.output_var.set(path)

    def log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.root.update_idletasks()

    def run(self):
        try:
            excel_path = self.excel_var.get().strip()
            template_path = self.template_var.get().strip()
            output_path = self.output_var.get().strip()

            room_count = int(self.room_count_var.get().strip())
            capacity_raw = self.capacity_var.get().strip()
            capacity = int(capacity_raw) if capacity_raw else None

            room_names = [x.strip() for x in self.room_names_var.get().split(",") if x.strip()]
            if not room_names:
                raise ValueError("Please provide room names.")
            if len(room_names) < room_count:
                raise ValueError("Room count is bigger than the number of room names provided.")
            room_names = room_names[:room_count]

            self.log_box.delete("1.0", "end")
            self.log("Starting generation...")
            self.log(f"Excel: {excel_path}")
            self.log(f"Template: {template_path}")
            self.log(f"Output: {output_path}")
            self.log(f"Rooms: {room_names}")
            self.log(f"Capacity: {capacity if capacity else 'AUTO'}")

            result = generate_all(
                excel_path=excel_path,
                template_pdf_path=template_path,
                output_pdf_path=output_path,
                room_names=room_names,
                capacity=capacity,
                shuffle=self.shuffle_var.get(),
            )

            self.log("")
            self.log("Done.")
            self.log(f"Students: {result['students']}")
            self.log(f"Rooms: {result['rooms']}")
            self.log(f"Capacity used: {result['capacity']}")
            self.log(f"Output: {result['output']}")
            self.log("Distribution:")
            for room, count in result["distribution"].items():
                self.log(f"  - {room}: {count}")

            messagebox.showinfo("Success", f"Combined PDF created:\n{result['output']}")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log(f"ERROR: {e}")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
