import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# EXCEL REKAP TOOLS V3.0
# ============================================================

APP_TITLE = "EXCEL REKAP TOOLS"
APP_SUBTITLE = "Pengolahan & Rekap Data Excel"
APP_VERSION = "v3.0"


# ============================================================
# UTILITIES
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def extract_percentage(value):
    """
    Mengambil angka persentase dari teks seperti:
    PRS (10%)
    Tugas1 (20%)
    UTS (30%)
    """
    if value is None:
        return 0

    text = str(value)

    match = re.search(r"\((\d+(?:\.\d+)?)\s*%\)", text)

    if match:
        return float(match.group(1))

    return 0


def read_components_from_row5(ws):
    components = {}

    for cell in ws[5]:
        value = clean_text(cell.value)

        if not value:
            continue

        upper_value = value.upper()

        if "PRS" in upper_value:
            components["PRS"] = extract_percentage(value)

        elif "TUGAS1" in upper_value:
            components["Tugas1"] = extract_percentage(value)

        elif "TUGAS2" in upper_value:
            components["Tugas2"] = extract_percentage(value)

        elif "TUGAS3" in upper_value:
            components["Tugas3"] = extract_percentage(value)

        elif "QUIZ" in upper_value or "KUIS" in upper_value:
            components["QUIZ"] = extract_percentage(value)

        elif "PRESNTA" in upper_value:
            components["Presnta"] = extract_percentage(value)

        elif "LAB" in upper_value:
            components["LAB"] = extract_percentage(value)

        elif "UTS" in upper_value:
            components["UTS"] = extract_percentage(value)

        elif "UAS" in upper_value:
            components["UAS"] = extract_percentage(value)

    return components


def format_worksheet(ws):
    """
    Formatting umum untuk output Excel.
    """

    # Header
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # Border
    thin = Side(
        style="thin",
        color="D9E2F3"
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

    # Freeze header
    ws.freeze_panes = "A2"

    # Auto width
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:
            try:
                length = len(str(cell.value)) if cell.value is not None else 0

                if length > max_length:
                    max_length = length

            except Exception:
                pass

        ws.column_dimensions[column_letter].width = min(
            max(max_length + 2, 12),
            40
        )

    ws.row_dimensions[1].height = 35


# ============================================================
# MODULE 1
# REKAP MATA KULIAH & EVALUASI
# ============================================================

def process_excel(input_file, output_file):
    wb = load_workbook(
        input_file,
        data_only=False
    )

    output_wb = Workbook()

    ws_rekap = output_wb.active
    ws_rekap.title = "Rekap Utama"

    headers_rekap = [
        "No",
        "Nama Kelas",
        "Kode Matkul",
        "Nama Matkul",
        "PRS",
        "Tugas1",
        "Tugas2",
        "Tugas3",
        "QUIZ",
        "Presnta",
        "LAB",
        "UTS",
        "UAS",
        "Total",
        "Status"
    ]

    ws_rekap.append(headers_rekap)

    ws_eval = output_wb.create_sheet(
        "Evaluasi Detail"
    )

    headers_eval = [
        "No",
        "Nama Kelas",
        "Kode Matkul",
        "Nama Matkul",
        "Basis Evaluasi",
        "Komponen Evaluasi",
        "Bobot (%)",
        "Nama (Inggris)",
        "Status Konversi"
    ]

    ws_eval.append(headers_eval)

    nomor = 1

    for ws in wb.worksheets:

        nama_kelas = clean_text(
            ws["C6"].value
        )

        nama_matkul = clean_text(
            ws["D6"].value
        )

        kode_matkul = clean_text(
            ws["E6"].value
        )

        components = read_components_from_row5(ws)

        prs = components.get("PRS", 0)
        tugas1 = components.get("Tugas1", 0)
        tugas2 = components.get("Tugas2", 0)
        tugas3 = components.get("Tugas3", 0)
        quiz = components.get("QUIZ", 0)
        presnta = components.get("Presnta", 0)
        lab = components.get("LAB", 0)
        uts = components.get("UTS", 0)
        uas = components.get("UAS", 0)

        total = (
            prs
            + tugas1
            + tugas2
            + tugas3
            + quiz
            + presnta
            + lab
            + uts
            + uas
        )

        if abs(total - 100) < 0.01:
            status = "OK"
        else:
            status = "Perlu Review"

        ws_rekap.append([
            nomor,
            nama_kelas,
            kode_matkul,
            nama_matkul,
            prs,
            tugas1,
            tugas2,
            tugas3,
            quiz,
            presnta,
            lab,
            uts,
            uas,
            total,
            status
        ])

        # ----------------------------------------------------
        # Evaluasi Detail
        # ----------------------------------------------------

        # PRS
        if prs > 0:
            ws_eval.append([
                nomor,
                nama_kelas,
                kode_matkul,
                nama_matkul,
                2,
                "",
                prs,
                "Participatory Activity",
                "OK"
            ])

        # Tugas
        for tugas_name, tugas_value in [
            ("Tugas1", tugas1),
            ("Tugas2", tugas2),
            ("Tugas3", tugas3)
        ]:

            if tugas_value > 0:
                ws_eval.append([
                    nomor,
                    nama_kelas,
                    kode_matkul,
                    nama_matkul,
                    4,
                    "TGS",
                    tugas_value,
                    "Assignment",
                    "OK"
                ])

        # Quiz
        if quiz > 0:
            ws_eval.append([
                nomor,
                nama_kelas,
                kode_matkul,
                nama_matkul,
                4,
                "QUIZ",
                quiz,
                "Quiz",
                "OK"
            ])

        # Presnta
        if presnta > 0:
            ws_eval.append([
                nomor,
                nama_kelas,
                kode_matkul,
                nama_matkul,
                3,
                "",
                presnta,
                "Project Outcomes",
                "Provisional"
            ])

        # LAB
        if lab > 0:
            ws_eval.append([
                nomor,
                nama_kelas,
                kode_matkul,
                nama_matkul,
                3,
                "",
                lab,
                "Project Outcomes",
                "Provisional"
            ])

        # UTS
        if uts > 0:
            ws_eval.append([
                nomor,
                nama_kelas,
                kode_matkul,
                nama_matkul,
                4,
                "UTS",
                uts,
                "Midterm Exam",
                "OK"
            ])

        # UAS
        if uas > 0:
            ws_eval.append([
                nomor,
                nama_kelas,
                kode_matkul,
                nama_matkul,
                4,
                "UAS",
                uas,
                "Finalterm Exam",
                "OK"
            ])

        nomor += 1

    format_worksheet(ws_rekap)
    format_worksheet(ws_eval)

    output_wb.save(output_file)


# ============================================================
# MODULE 2
# REKAP MATA KULIAH & DOSEN
# ============================================================

def process_matkul_dosen(input_file, output_file):

    wb = load_workbook(
        input_file,
        data_only=True
    )

    output_wb = Workbook()

    ws_out = output_wb.active
    ws_out.title = "Rekap Mata Kuliah Dosen"

    headers = [
        "No",
        "Nama Sheet",
        "Kode Mata Kuliah",
        "Nama Mata Kuliah",
        "Nama Kelas",
        "Nama Dosen"
    ]

    ws_out.append(headers)

    nomor = 1

    for ws in wb.worksheets:

        nama_kelas = clean_text(
            ws["C6"].value
        )

        kode_mk = clean_text(
            ws["D6"].value
        )

        nama_mk = clean_text(
            ws["E6"].value
        )

        nama_dosen = clean_text(
            ws["L2"].value
        )

        ws_out.append([
            nomor,
            ws.title,
            kode_mk,
            nama_mk,
            nama_kelas,
            nama_dosen
        ])

        nomor += 1

    format_worksheet(ws_out)

    output_wb.save(output_file)


# ============================================================
# MODULE 3
# MERGE SEMUA SHEET
# ============================================================

def process_merge_sheets(input_file, output_file):

    wb = load_workbook(
        input_file,
        data_only=True
    )

    output_wb = Workbook()

    ws_out = output_wb.active
    ws_out.title = "Merge Semua Sheet"

    all_rows = []
    max_columns = 0

    # Ambil seluruh data mulai baris ke-5
    for ws in wb.worksheets:

        for row in ws.iter_rows(
            min_row=5,
            max_row=ws.max_row,
            values_only=True
        ):

            values = list(row)

            # Skip baris kosong
            if not any(
                value is not None and str(value).strip() != ""
                for value in values
            ):
                continue

            all_rows.append(
                values + [ws.title]
            )

            if len(values) > max_columns:
                max_columns = len(values)

    # Header
    if all_rows:

        first_row = all_rows[0]

        headers = []

        for i in range(
            len(first_row) - 1
        ):
            headers.append(
                f"Column {i + 1}"
            )

        headers.append("Nama Sheet")

        ws_out.append(headers)

        # Data
        for row in all_rows:

            normalized = row + [
                None
            ] * (
                len(headers) - len(row)
            )

            ws_out.append(
                normalized[:len(headers)]
            )

    else:

        ws_out.append([
            "Tidak ada data"
        ])

    format_worksheet(ws_out)

    output_wb.save(output_file)


# ============================================================
# GUI APPLICATION
# ============================================================

class ExcelRekapApp:

    def __init__(self, root):

        self.root = root

        self.root.title(
            f"{APP_TITLE} {APP_VERSION}"
        )

        self.root.geometry(
            "1050x680"
        )

        self.root.minsize(
            900,
            600
        )

        self.root.configure(
            bg="#F4F7FB"
        )

        self.selected_file = ""
        self.selected_module = 0

        self.setup_styles()
        self.build_ui()

        self.show_home()

    # ========================================================
    # STYLE
    # ========================================================

    def setup_styles(self):

        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "TProgressbar",
            thickness=8
        )

        style.configure(
            "Modern.TButton",
            font=(
                "Segoe UI",
                10,
                "bold"
            ),
            padding=10
        )

    # ========================================================
    # MAIN UI
    # ========================================================

    def build_ui(self):

        # ====================================================
        # HEADER
        # ====================================================

        self.header = tk.Frame(
            self.root,
            bg="#17365D",
            height=85
        )

        self.header.pack(
            side="top",
            fill="x"
        )

        self.header.pack_propagate(
            False
        )

        title_frame = tk.Frame(
            self.header,
            bg="#17365D"
        )

        title_frame.pack(
            side="left",
            padx=25
        )

        tk.Label(
            title_frame,
            text="📊",
            font=(
                "Segoe UI Emoji",
                28
            ),
            bg="#17365D",
            fg="white"
        ).pack(
            side="left",
            padx=(0, 12)
        )

        text_frame = tk.Frame(
            title_frame,
            bg="#17365D"
        )

        text_frame.pack(
            side="left"
        )

        tk.Label(
            text_frame,
            text=APP_TITLE,
            font=(
                "Segoe UI",
                19,
                "bold"
            ),
            bg="#17365D",
            fg="white"
        ).pack(
            anchor="w"
        )

        tk.Label(
            text_frame,
            text=APP_SUBTITLE,
            font=(
                "Segoe UI",
                9
            ),
            bg="#17365D",
            fg="#D9EAF7"
        ).pack(
            anchor="w"
        )

        tk.Label(
            self.header,
            text=APP_VERSION,
            font=(
                "Segoe UI",
                10,
                "bold"
            ),
            bg="#17365D",
            fg="#B8D8F0"
        ).pack(
            side="right",
            padx=25
        )

        # ====================================================
        # BODY
        # ====================================================

        body = tk.Frame(
            self.root,
            bg="#F4F7FB"
        )

        body.pack(
            fill="both",
            expand=True
        )

        # ====================================================
        # SIDEBAR
        # ====================================================

        self.sidebar = tk.Frame(
            body,
            bg="#102A43",
            width=220
        )

        self.sidebar.pack(
            side="left",
            fill="y"
        )

        self.sidebar.pack_propagate(
            False
        )

        tk.Label(
            self.sidebar,
            text="MENU UTAMA",
            font=(
                "Segoe UI",
                9,
                "bold"
            ),
            bg="#102A43",
            fg="#9FB3C8"
        ).pack(
            anchor="w",
            padx=20,
            pady=(25, 10)
        )

        self.create_menu_button(
            "🏠  Beranda",
            self.show_home
        )

        self.create_menu_button(
            "📊  Rekap Evaluasi",
            lambda: self.show_module(0)
        )

        self.create_menu_button(
            "👨‍🏫  Rekap Dosen",
            lambda: self.show_module(1)
        )

        self.create_menu_button(
            "📑  Merge Sheet",
            lambda: self.show_module(2)
        )

        tk.Frame(
            self.sidebar,
            bg="#102A43",
            height=20
        ).pack()

        self.create_menu_button(
            "ℹ️  Tentang",
            self.show_about
        )

        # ====================================================
        # CONTENT
        # ====================================================

        self.content = tk.Frame(
            body,
            bg="#F4F7FB"
        )

        self.content.pack(
            side="left",
            fill="both",
            expand=True
        )

        # ====================================================
        # FOOTER
        # ====================================================

        self.footer = tk.Frame(
            self.root,
            bg="#E8EEF5",
            height=35
        )

        self.footer.pack(
            side="bottom",
            fill="x"
        )

        self.footer.pack_propagate(
            False
        )

        self.status_label = tk.Label(
            self.footer,
            text="● Ready",
            font=(
                "Segoe UI",
                9
            ),
            bg="#E8EEF5",
            fg="#486581"
        )

        self.status_label.pack(
            side="left",
            padx=20
        )

        tk.Label(
            self.footer,
            text=f"Excel Rekap Tools {APP_VERSION}",
            font=(
                "Segoe UI",
                9
            ),
            bg="#E8EEF5",
            fg="#829AB1"
        ).pack(
            side="right",
            padx=20
        )

    # ========================================================
    # SIDEBAR BUTTON
    # ========================================================

    def create_menu_button(
        self,
        text,
        command
    ):

        button = tk.Button(
            self.sidebar,
            text=text,
            command=command,
            font=(
                "Segoe UI",
                10
            ),
            bg="#102A43",
            fg="#FFFFFF",
            activebackground="#1F4E78",
            activeforeground="#FFFFFF",
            bd=0,
            relief="flat",
            anchor="w",
            padx=20,
            pady=13,
            cursor="hand2"
        )

        button.pack(
            fill="x"
        )

        def on_enter(event):
            button.configure(
                bg="#1F4E78"
            )

        def on_leave(event):
            button.configure(
                bg="#102A43"
            )

        button.bind(
            "<Enter>",
            on_enter
        )

        button.bind(
            "<Leave>",
            on_leave
        )

    # ========================================================
    # CLEAR CONTENT
    # ========================================================

    def clear_content(self):

        for widget in self.content.winfo_children():
            widget.destroy()

    # ========================================================
    # HOME
    # ========================================================

    def show_home(self):

        self.clear_content()

        self.status_label.config(
            text="● Ready"
        )

        tk.Label(
            self.content,
            text="Selamat Datang",
            font=(
                "Segoe UI",
                24,
                "bold"
            ),
            bg="#F4F7FB",
            fg="#17365D"
        ).pack(
            anchor="w",
            padx=35,
            pady=(35, 5)
        )

        tk.Label(
            self.content,
            text="Pilih salah satu tools untuk mulai mengolah data Excel.",
            font=(
                "Segoe UI",
                11
            ),
            bg="#F4F7FB",
            fg="#627D98"
        ).pack(
            anchor="w",
            padx=35
        )

        cards = tk.Frame(
            self.content,
            bg="#F4F7FB"
        )

        cards.pack(
            fill="x",
            padx=35,
            pady=35
        )

        self.create_card(
            cards,
            0,
            "📊",
            "Rekap Evaluasi",
            "Rekap mata kuliah dan\nkomponen evaluasi.",
            "#1F4E78",
            lambda: self.show_module(0)
        )

        self.create_card(
            cards,
            1,
            "👨‍🏫",
            "Rekap Dosen",
            "Rekap mata kuliah,\nkelas, dan dosen.",
            "#2E7D32",
            lambda: self.show_module(1)
        )

        self.create_card(
            cards,
            2,
            "📑",
            "Merge Sheet",
            "Menggabungkan data\ndari seluruh sheet.",
            "#7B1FA2",
            lambda: self.show_module(2)
        )

        info = tk.Frame(
            self.content,
            bg="white",
            bd=0,
            highlightthickness=1,
            highlightbackground="#D9E2EC"
        )

        info.pack(
            fill="x",
            padx=35,
            pady=(0, 20)
        )

        tk.Label(
            info,
            text="💡  Tips",
            font=(
                "Segoe UI",
                11,
                "bold"
            ),
            bg="white",
            fg="#17365D"
        ).pack(
            anchor="w",
            padx=20,
            pady=(15, 5)
        )

        tk.Label(
            info,
            text=(
                "Pastikan file Excel yang dipilih memiliki struktur "
                "data yang sesuai dengan tools yang digunakan."
            ),
            font=(
                "Segoe UI",
                10
            ),
            bg="white",
            fg="#627D98"
        ).pack(
            anchor="w",
            padx=20,
            pady=(0, 15)
        )

    # ========================================================
    # CARD
    # ========================================================

    def create_card(
        self,
        parent,
        column,
        icon,
        title,
        description,
        accent,
        command
    ):

        card = tk.Frame(
            parent,
            bg="white",
            width=220,
            height=190,
            highlightthickness=1,
            highlightbackground="#D9E2EC"
        )

        card.grid(
            row=0,
            column=column,
            padx=8,
            sticky="nsew"
        )

        parent.grid_columnconfigure(
            column,
            weight=1
        )

        card.grid_propagate(
            False
        )

        tk.Label(
            card,
            text=icon,
            font=(
                "Segoe UI Emoji",
                30
            ),
            bg="white",
            fg=accent
        ).pack(
            pady=(20, 5)
        )

        tk.Label(
            card,
            text=title,
            font=(
                "Segoe UI",
                12,
                "bold"
            ),
            bg="white",
            fg="#17365D"
        ).pack()

        tk.Label(
            card,
            text=description,
            font=(
                "Segoe UI",
                9
            ),
            bg="white",
            fg="#627D98",
            justify="center"
        ).pack(
            pady=7
        )

        button = tk.Button(
            card,
            text="BUKA TOOL",
            command=command,
            font=(
                "Segoe UI",
                8,
                "bold"
            ),
            bg=accent,
            fg="white",
            activebackground=accent,
            activeforeground="white",
            bd=0,
            padx=15,
            pady=7,
            cursor="hand2"
        )

        button.pack(
            pady=5
        )

    # ========================================================
    # MODULE PAGE
    # ========================================================

    def show_module(
        self,
        module_index
    ):

        self.selected_module = module_index

        self.clear_content()

        modules = [
            (
                "📊",
                "Rekap Mata Kuliah & Evaluasi",
                "Mengolah komponen evaluasi dan bobot mata kuliah.",
                "#1F4E78"
            ),
            (
                "👨‍🏫",
                "Rekap Mata Kuliah & Dosen",
                "Mengambil data mata kuliah, kelas, dan dosen dari setiap sheet.",
                "#2E7D32"
            ),
            (
                "📑",
                "Merge Semua Sheet",
                "Menggabungkan seluruh data sheet mulai dari baris ke-5.",
                "#7B1FA2"
            )
        ]

        icon, title, description, accent = modules[
            module_index
        ]

        # Title
        tk.Label(
            self.content,
            text=f"{icon}  {title}",
            font=(
                "Segoe UI",
                20,
                "bold"
            ),
            bg="#F4F7FB",
            fg="#17365D"
        ).pack(
            anchor="w",
            padx=35,
            pady=(30, 5)
        )

        tk.Label(
            self.content,
            text=description,
            font=(
                "Segoe UI",
                10
            ),
            bg="#F4F7FB",
            fg="#627D98"
        ).pack(
            anchor="w",
            padx=35
        )

        # Main panel
        panel = tk.Frame(
            self.content,
            bg="white",
            highlightthickness=1,
            highlightbackground="#D9E2EC"
        )

        panel.pack(
            fill="x",
            padx=35,
            pady=30
        )

        # File input
        tk.Label(
            panel,
            text="1. Pilih File Excel",
            font=(
                "Segoe UI",
                11,
                "bold"
            ),
            bg="white",
            fg="#17365D"
        ).pack(
            anchor="w",
            padx=25,
            pady=(25, 8)
        )

        file_frame = tk.Frame(
            panel,
            bg="white"
        )

        file_frame.pack(
            fill="x",
            padx=25
        )

        self.file_entry = tk.Entry(
            file_frame,
            font=(
                "Segoe UI",
                10
            ),
            bd=1,
            relief="solid",
            bg="#F8FAFC"
        )

        self.file_entry.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=8
        )

        tk.Button(
            file_frame,
            text="📂  Pilih File",
            command=self.choose_file,
            font=(
                "Segoe UI",
                9,
                "bold"
            ),
            bg=accent,
            fg="white",
            activebackground=accent,
            activeforeground="white",
            bd=0,
            padx=18,
            pady=9,
            cursor="hand2"
        ).pack(
            side="left",
            padx=(10, 0)
        )

        # Output
        tk.Label(
            panel,
            text="2. Lokasi File Output",
            font=(
                "Segoe UI",
                11,
                "bold"
            ),
            bg="white",
            fg="#17365D"
        ).pack(
            anchor="w",
            padx=25,
            pady=(25, 8)
        )

        output_names = [
            "Rekap_Mata_Kuliah_dan_Evaluasi.xlsx",
            "Rekap_Mata_Kuliah_dan_Dosen.xlsx",
            "Merge_Semua_Sheet.xlsx"
        ]

        self.output_name = output_names[
            module_index
        ]

        tk.Label(
            panel,
            text=f"File akan disimpan sebagai: {self.output_name}",
            font=(
                "Segoe UI",
                9
            ),
            bg="white",
            fg="#829AB1"
        ).pack(
            anchor="w",
            padx=25
        )

        # Process button
        process_button = tk.Button(
            panel,
            text="▶  PROSES DATA",
            command=self.process,
            font=(
                "Segoe UI",
                11,
                "bold"
            ),
            bg="#2E7D32",
            fg="white",
            activebackground="#256628",
            activeforeground="white",
            bd=0,
            padx=30,
            pady=12,
            cursor="hand2"
        )

        process_button.pack(
            pady=(25, 10)
        )

        # Progress
        self.progress = ttk.Progressbar(
            panel,
            mode="indeterminate"
        )

        self.progress.pack(
            fill="x",
            padx=25,
            pady=(5, 25)
        )

        self.module_status = tk.Label(
            panel,
            text="Siap. Silakan pilih file Excel.",
            font=(
                "Segoe UI",
                9
            ),
            bg="white",
            fg="#627D98"
        )

        self.module_status.pack(
            pady=(0, 20)
        )

    # ========================================================
    # CHOOSE FILE
    # ========================================================

    def choose_file(self):

        file_path = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx *.xlsm"
                ),
                (
                    "All Files",
                    "*.*"
                )
            ]
        )

        if not file_path:
            return

        self.selected_file = file_path

        self.file_entry.delete(
            0,
            tk.END
        )

        self.file_entry.insert(
            0,
            file_path
        )

        self.module_status.config(
            text="File berhasil dipilih."
        )

        self.status_label.config(
            text="● File Excel dipilih"
        )

    # ========================================================
    # PROCESS
    # ========================================================

    def process(self):

        if not self.selected_file:

            messagebox.showwarning(
                "File Belum Dipilih",
                "Silakan pilih file Excel terlebih dahulu."
            )

            return

        output_file = filedialog.asksaveasfilename(
            title="Simpan File Output",
            defaultextension=".xlsx",
            initialfile=self.output_name,
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx"
                )
            ]
        )

        if not output_file:
            return

        try:

            self.progress.start(10)

            self.module_status.config(
                text="Sedang memproses data..."
            )

            self.status_label.config(
                text="● Processing..."
            )

            self.root.update_idletasks()

            if self.selected_module == 0:

                process_excel(
                    self.selected_file,
                    output_file
                )

            elif self.selected_module == 1:

                process_matkul_dosen(
                    self.selected_file,
                    output_file
                )

            elif self.selected_module == 2:

                process_merge_sheets(
                    self.selected_file,
                    output_file
                )

            self.progress.stop()

            self.module_status.config(
                text="✓ Proses selesai dengan sukses."
            )

            self.status_label.config(
                text="● Selesai"
            )

            messagebox.showinfo(
                "Berhasil",
                "Data berhasil diproses!\n\n"
                f"File output:\n{output_file}"
            )

        except Exception as e:

            self.progress.stop()

            self.module_status.config(
                text="✕ Terjadi kesalahan."
            )

            self.status_label.config(
                text="● Error"
            )

            messagebox.showerror(
                "Error",
                "Terjadi kesalahan saat memproses file:\n\n"
                f"{str(e)}"
            )

    # ========================================================
    # ABOUT
    # ========================================================

    def show_about(self):

        self.clear_content()

        tk.Label(
            self.content,
            text="ℹ️  Tentang Aplikasi",
            font=(
                "Segoe UI",
                22,
                "bold"
            ),
            bg="#F4F7FB",
            fg="#17365D"
        ).pack(
            anchor="w",
            padx=35,
            pady=(35, 20)
        )

        panel = tk.Frame(
            self.content,
            bg="white",
            highlightthickness=1,
            highlightbackground="#D9E2EC"
        )

        panel.pack(
            fill="x",
            padx=35
        )

        tk.Label(
            panel,
            text="📊",
            font=(
                "Segoe UI Emoji",
                45
            ),
            bg="white"
        ).pack(
            pady=(30, 5)
        )

        tk.Label(
            panel,
            text=APP_TITLE,
            font=(
                "Segoe UI",
                20,
                "bold"
            ),
            bg="white",
            fg="#17365D"
        ).pack()

        tk.Label(
            panel,
            text=APP_SUBTITLE,
            font=(
                "Segoe UI",
                10
            ),
            bg="white",
            fg="#627D98"
        ).pack(
            pady=5
        )

        tk.Label(
            panel,
            text=APP_VERSION,
            font=(
                "Segoe UI",
                10,
                "bold"
            ),
            bg="white",
            fg="#2E7D32"
        ).pack(
            pady=5
        )

        tk.Label(
            panel,
            text=(
                "\nAplikasi untuk membantu pengolahan dan "
                "rekapitulasi data akademik berbasis Excel.\n\n"
                "Modul yang tersedia:\n"
                "• Rekap Mata Kuliah & Evaluasi\n"
                "• Rekap Mata Kuliah & Dosen\n"
                "• Merge Semua Sheet"
            ),
            font=(
                "Segoe UI",
                10
            ),
            bg="white",
            fg="#486581",
            justify="center"
        ).pack(
            pady=(10, 30)
        )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = ExcelRekapApp(
        root
    )

    root.mainloop()
