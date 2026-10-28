#!/usr/bin/env python3
"""Generate exam-room incident logs from the access-arrangements workbook."""

from __future__ import annotations

import copy
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from posixpath import normpath
from tkinter import StringVar, Tk, filedialog, messagebox, ttk
from xml.etree import ElementTree as ET

import app_updates


BASE_DIR = Path(__file__).resolve().parent


def bundled_path(filename: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / filename
    return BASE_DIR / filename


DEFAULT_DOCX = bundled_path("Exam Room Incident Log Template.docx")
DEFAULT_OUTPUT = Path.home() / "Downloads" / "Generated Incident Logs"
NO_ADDITIONAL_SUBJECT = "No additional subject"
MAX_EXAMS = 8

NS_X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ET.register_namespace("w", NS_W)


@dataclass(frozen=True)
class Student:
    name: str
    arrangements: str
    subjects: tuple[str, ...]
    room: str = ""


@dataclass(frozen=True)
class ExamSection:
    subject: str
    students: tuple[Student, ...]


def sections_by_room(sections: tuple[ExamSection, ...]) -> dict[str, tuple[ExamSection, ...]]:
    rooms = {}
    labels = {}
    for section in sections:
        groups = {}
        for student in section.students:
            room = clean(student.room) or "Room not specified"
            room_key = room.casefold()
            if room_key == "alternative room":
                continue
            labels.setdefault(room_key, room)
            groups.setdefault(room_key, []).append(student)
        for room_key, students in groups.items():
            rooms.setdefault(room_key, []).append(ExamSection(section.subject, tuple(students)))
    return {labels[room_key]: tuple(groups) for room_key, groups in rooms.items()}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean(value).lower())


def output_lines(section: ExamSection) -> list[str]:
    lines = [section.subject, "Exam Candidate\tAccess arrangements"]
    lines.extend(f"{student.name}\t{student.arrangements}" for student in section.students)
    return lines


def open_folder(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    if sys.platform == "win32":
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def col_index(cell_ref: str) -> int:
    letters = "".join(char for char in cell_ref if char.isalpha())
    index = 0
    for char in letters:
        index = index * 26 + ord(char.upper()) - 64
    return index


class WorkbookReader:
    def __init__(self, path: Path):
        self.path = path
        self.sheets: dict[str, list[dict[int, str]]] = {}
        self.load()

    def load(self) -> None:
        with zipfile.ZipFile(self.path) as archive:
            shared = self._shared_strings(archive)
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

            for sheet in workbook.findall(f".//{{{NS_X}}}sheet"):
                name = sheet.attrib["name"]
                rel_id = sheet.attrib[f"{{{NS_REL}}}id"]
                target = rel_map[rel_id]
                if target.startswith("/"):
                    target = target.lstrip("/")
                elif not target.startswith("xl/"):
                    target = normpath(f"xl/{target}")
                self.sheets[name] = self._sheet_rows(archive, target, shared)

    def _shared_strings(self, archive: zipfile.ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return []
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        strings = []
        for item in root.findall(f"{{{NS_X}}}si"):
            strings.append("".join(node.text or "" for node in item.iter(f"{{{NS_X}}}t")))
        return strings

    def _sheet_rows(
        self, archive: zipfile.ZipFile, target: str, shared: list[str]
    ) -> list[dict[int, str]]:
        root = ET.fromstring(archive.read(target))
        rows: list[dict[int, str]] = []
        for row_node in root.findall(f".//{{{NS_X}}}sheetData/{{{NS_X}}}row"):
            row: dict[int, str] = {}
            for cell in row_node.findall(f"{{{NS_X}}}c"):
                ref = cell.attrib.get("r", "")
                if not ref:
                    continue
                value = self._cell_value(cell, shared)
                if value != "":
                    row[col_index(ref)] = value
            rows.append(row)
        return rows

    def _cell_value(self, cell: ET.Element, shared: list[str]) -> str:
        if cell.attrib.get("t") == "inlineStr":
            return clean("".join(node.text or "" for node in cell.iter(f"{{{NS_X}}}t")))

        value_node = cell.find(f"{{{NS_X}}}v")
        if value_node is None or value_node.text is None:
            return ""

        value = value_node.text
        if cell.attrib.get("t") == "s":
            return clean(shared[int(value)])
        if value.endswith(".0"):
            value = value[:-2]
        return clean(value)


class ArrangementData:
    def __init__(self, workbook_path: Path):
        reader = WorkbookReader(workbook_path)
        self.years: dict[str, list[Student]] = {}
        for sheet_name, rows in reader.sheets.items():
            if re.fullmatch(r"Year \d+", sheet_name):
                students = self._parse_year(rows)
                if students:
                    self.years[sheet_name] = students
        self.subject_index = self._build_subject_index()
        self.subject_names = self._build_subject_names()

    def subjects(self) -> list[str]:
        return ["All subjects"] + sorted(self.subject_names.values(), key=str.lower)

    def subjects_for_year(self, year: str, include_optional_blank: bool = False) -> list[str]:
        subject_names = {
            self.subject_names[subject_key]
            for subject_key in self.subject_index.get(year, {})
            if subject_key in self.subject_names
        }
        subjects = ["All subjects"] + sorted(subject_names, key=str.lower)
        if include_optional_blank:
            return [NO_ADDITIONAL_SUBJECT] + subjects
        return subjects

    def students_for(self, year: str, subject: str) -> list[Student]:
        selected = key(subject)
        if selected == key("All subjects"):
            return self.years.get(year, [])
        return self.subject_index.get(year, {}).get(selected, [])

    def year_has_subject(self, year: str, subject: str) -> bool:
        return bool(self.students_for(year, subject))

    def _build_subject_index(self) -> dict[str, dict[str, list[Student]]]:
        index: dict[str, dict[str, list[Student]]] = {}
        for year, students in self.years.items():
            year_index: dict[str, dict[tuple[str, str, str], Student]] = {}
            for student in students:
                student_key = (key(student.name), key(student.arrangements), clean(student.room).casefold())
                for subject in student.subjects:
                    for subject_key in self._subject_keys(subject):
                        year_index.setdefault(subject_key, {})[student_key] = student
            index[year] = {
                subject_key: sorted(matches.values(), key=lambda item: item.name.lower())
                for subject_key, matches in year_index.items()
            }
        return index

    def _build_subject_names(self) -> dict[str, str]:
        names = {}
        for students in self.years.values():
            for student in students:
                for subject in student.subjects:
                    for subject_key, display in self._subject_display_pairs(subject):
                        names.setdefault(subject_key, display)
        return names

    def _subject_keys(self, subject: str) -> tuple[str, ...]:
        return tuple(subject_key for subject_key, _display in self._subject_display_pairs(subject))

    def _subject_display_pairs(self, subject: str) -> tuple[tuple[str, str], ...]:
        subject = clean(subject)
        if not subject:
            return ()
        pairs = [(key(subject), subject)]
        if ":" in subject:
            parent = clean(subject.split(":", 1)[0])
            if parent:
                pairs.append((key(parent), parent))
        return tuple(dict.fromkeys(pairs))

    def _parse_year(self, rows: list[dict[int, str]]) -> list[Student]:
        header_index = self._header_index(rows)
        if header_index is None:
            return []

        headers = rows[header_index]
        header_map = {key(value): col for col, value in headers.items() if value}
        first_col = header_map.get("firstname")
        surname_col = header_map.get("surname")
        arrangement_col = header_map.get("accessarrangement") or header_map.get(
            "accessarrangements"
        )
        subject_col = header_map.get("subjectsapplicable")
        room_col = header_map.get("room")
        if not first_col or not surname_col or not arrangement_col:
            return []

        subject_cols = self._subject_columns(rows, header_index, headers, subject_col)
        students: dict[tuple[str, str, str], Student] = {}
        for row in rows[header_index + 1 :]:
            first = clean(row.get(first_col, ""))
            surname = clean(row.get(surname_col, ""))
            name = clean(f"{first} {surname}")
            arrangements = clean(row.get(arrangement_col, ""))
            room = clean(row.get(room_col, ""))
            if not name or not arrangements:
                continue

            subjects = self._row_subjects(row, subject_col, subject_cols)
            unique_key = (key(name), key(arrangements), room.casefold())
            existing = students.get(unique_key)
            if existing:
                subjects = tuple(dict.fromkeys((*existing.subjects, *subjects)))
            students[unique_key] = Student(name, arrangements, subjects, room)
        return list(students.values())

    def _header_index(self, rows: list[dict[int, str]]) -> int | None:
        for index, row in enumerate(rows):
            values = {key(value) for value in row.values()}
            if "firstname" in values and "surname" in values:
                return index
        return None

    def _subject_columns(
        self,
        rows: list[dict[int, str]],
        header_index: int,
        headers: dict[int, str],
        subject_col: int | None,
    ) -> dict[int, str]:
        fixed_cols = {
            col
            for col, value in headers.items()
            if key(value)
            in {
                "firstname",
                "surname",
                "candidatenumber",
                "needs",
                "accessarrangement",
                "accessarrangements",
                "room",
                "evidencereports",
                "form189",
                "paao",
                "aao",
            }
        }
        start_col = max(fixed_cols or {0}) + 1
        subject_row = rows[header_index - 1] if header_index > 0 else {}
        return {
            col: clean(value)
            for col, value in subject_row.items()
            if col >= start_col and clean(value)
        }

    def _row_subjects(
        self, row: dict[int, str], subject_col: int | None, subject_cols: dict[int, str]
    ) -> tuple[str, ...]:
        if subject_cols:
            subjects = []
            for col, subject in subject_cols.items():
                marker = key(row.get(col, ""))
                if marker in {"y", "yes", "1", "true"}:
                    subjects.append(subject)
            return tuple(subjects)
        if subject_col:
            return tuple(
                part.strip()
                for part in re.split(r"[,;/\n]+", row.get(subject_col, ""))
                if part.strip()
            )
        return ()


class IncidentLogWriter:
    def __init__(self, template_path: Path):
        self.template_path = template_path

    def write(self, output_path: Path, students: list[Student]) -> None:
        self.write_sections(output_path, (ExamSection("Students", tuple(students)),))

    def write_sections(self, output_path: Path, sections: tuple[ExamSection, ...]) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(self.template_path, output_path)

        with zipfile.ZipFile(output_path, "r") as source:
            files = {name: source.read(name) for name in source.namelist()}

        root = ET.fromstring(files["word/document.xml"])
        first_table = root.find(f".//{{{NS_W}}}tbl")
        if first_table is None:
            raise ValueError("The Word template does not contain a table to fill.")

        rows = first_table.findall(f"{{{NS_W}}}tr")
        if len(rows) < 2:
            raise ValueError("The first table must include a header and one blank row.")

        blank_row = copy.deepcopy(rows[1])
        needed_rows = max(sum(1 + len(section.students) for section in sections), 1)
        body_rows = rows[1:]
        for row in body_rows[needed_rows:]:
            first_table.remove(row)
        while len(first_table.findall(f"{{{NS_W}}}tr")) - 1 < needed_rows:
            first_table.append(copy.deepcopy(blank_row))

        body_rows = first_table.findall(f"{{{NS_W}}}tr")[1:]
        row_index = 0
        for section in sections:
            cells = body_rows[row_index].findall(f"{{{NS_W}}}tc")
            if len(cells) >= 4:
                self._set_cell_text(cells[0], section.subject, bold=True)
                self._set_cell_text(cells[1], "")
                self._set_cell_text(cells[2], "")
                self._set_cell_text(cells[3], "")
            row_index += 1
            for student in section.students:
                cells = body_rows[row_index].findall(f"{{{NS_W}}}tc")
                if len(cells) >= 2:
                    self._set_cell_text(cells[0], student.name)
                    self._set_cell_text(cells[1], student.arrangements)
                row_index += 1
        for row in body_rows[row_index:]:
            cells = row.findall(f"{{{NS_W}}}tc")
            for cell in cells[:4]:
                self._set_cell_text(cell, "")

        files["word/document.xml"] = ET.tostring(
            root, encoding="utf-8", xml_declaration=True
        )
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as target:
            for name, data in files.items():
                target.writestr(name, data)

    def _set_cell_text(self, cell: ET.Element, text: str, bold: bool = False) -> None:
        for child in list(cell):
            if child.tag != f"{{{NS_W}}}tcPr":
                cell.remove(child)
        paragraph = ET.SubElement(cell, f"{{{NS_W}}}p")
        run = ET.SubElement(paragraph, f"{{{NS_W}}}r")
        if bold:
            run_properties = ET.SubElement(run, f"{{{NS_W}}}rPr")
            ET.SubElement(run_properties, f"{{{NS_W}}}b")
        text_node = ET.SubElement(run, f"{{{NS_W}}}t")
        text_node.text = text


class TkText(ttk.Frame):
    def __init__(self, parent, **text_options):
        super().__init__(parent)
        from tkinter import Text

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.text = Text(self, **text_options)
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

    def set_text(self, value: str) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", value)
        self.text.configure(state="disabled")


class Tooltip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None) -> None:
        if self.tip:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + 18
        from tkinter import Label, Toplevel

        window = Toplevel(self.widget)
        self.tip = window
        window.wm_overrideredirect(True)
        window.wm_geometry(f"+{x}+{y}")
        label = Label(
            window,
            text=self.text,
            justify="left",
            background="#f7f7f7",
            foreground="#555555",
            relief="solid",
            borderwidth=1,
            padx=6,
            pady=4,
            font=("TkDefaultFont", 9),
        )
        label.pack()

    def hide(self, _event=None) -> None:
        if self.tip:
            self.tip.destroy()
            self.tip = None


class InfoIcon(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        from tkinter import Canvas

        canvas = Canvas(self, width=16, height=16, highlightthickness=0)
        canvas.create_oval(3, 3, 13, 13, outline="#9aa0a6", width=1)
        canvas.create_text(8, 8, text="i", fill="#8a8f95", font=("TkDefaultFont", 8, "bold"))
        canvas.pack()
        Tooltip(canvas, "Designed by Emmanuel\nFor Jilyan")


class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(f"Exam Access Incident Log Generator - {app_updates.VERSION}")
        self.root.geometry("760x560")
        self.root.minsize(700, 520)

        self.xlsx_path = StringVar(value="")
        self.output_dir = StringVar(value=str(DEFAULT_OUTPUT))
        self.status = StringVar(value="Choose the exam groups, then generate logs.")
        self.exam_rows = []
        self.data: ArrangementData | None = None
        self.preview = None
        self.update_busy = False

        self._build()
        self.set_exam_controls_enabled(False)
        self.status.set("Choose the access arrangements spreadsheet to begin.")

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)

        self._file_row(outer, 0, "Access arrangements spreadsheet", self.xlsx_path, self._choose_xlsx)
        self._file_row(outer, 1, "Output folder", self.output_dir, self._choose_output)

        ttk.Button(outer, text="Reload spreadsheet", command=self.load_workbook).grid(
            row=2, column=0, sticky="w", pady=(10, 14)
        )

        ttk.Label(outer, text="Exam groups").grid(row=3, column=0, sticky="nw")
        self.exam_frame = ttk.Frame(outer)
        self.exam_frame.grid(row=3, column=1, columnspan=2, sticky="ew")
        self.exam_frame.columnconfigure(2, weight=1)
        self._add_exam_row(0, "Exam 1", optional=False)
        self._add_exam_row(1, "Exam 2", optional=True)
        self._add_exam_row(2, "Exam 3", optional=True)
        self.add_exam_button = ttk.Button(
            self.exam_frame, text="+ Additional exam", command=self.add_exam
        )
        self.add_exam_button.grid(row=3, column=1, columnspan=2, sticky="w", pady=(6, 0))

        button_row = ttk.Frame(outer)
        button_row.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(24, 10))
        ttk.Button(button_row, text="Generate room logs", command=self.generate).pack(
            side="left"
        )
        ttk.Button(button_row, text="Copy list", command=self.copy_preview).pack(
            side="left", padx=(10, 0)
        )
        ttk.Button(button_row, text="Open output folder", command=self.open_output_folder).pack(
            side="left", padx=(10, 0)
        )

        ttk.Label(outer, text="Preview").grid(row=5, column=0, sticky="nw")
        self.preview = TkText(outer, height=12, wrap="word")
        self.preview.grid(row=5, column=1, columnspan=2, sticky="nsew", pady=(0, 10))
        outer.rowconfigure(5, weight=1)

        footer = ttk.Frame(outer)
        footer.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status, wraplength=450).grid(
            row=0, column=0, sticky="ew"
        )
        self.update_button = ttk.Button(footer, text="Check for updates", command=self.check_updates)
        self.update_button.grid(row=0, column=1, sticky="e", padx=(8, 0))
        InfoIcon(footer).grid(row=0, column=2, sticky="e", padx=(8, 0))

    def update_task(self, task, finished) -> None:
        self.update_busy = True
        self.update_button.configure(state="disabled")
        results = queue.Queue()

        def work():
            try:
                results.put((task(), None))
            except Exception as exc:
                results.put((None, str(exc)))

        def poll():
            try:
                result, error = results.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            self.update_busy = False
            self.update_button.configure(state="normal")
            if error:
                self.status.set("Update could not be completed.")
                messagebox.showerror("Update unavailable", error)
            else:
                finished(result)

        threading.Thread(target=work, daemon=True).start()
        self.root.after(100, poll)

    def check_updates(self) -> None:
        if self.update_busy:
            return
        self.status.set("Checking for updates...")
        self.update_task(app_updates.check_for_update, self.update_checked)

    def update_checked(self, update) -> None:
        self.status.set(f"Current version: {app_updates.VERSION}")
        if update is None:
            messagebox.showinfo("Updates", "You have the latest available version.")
            return
        if sys.platform != "win32" or not getattr(sys, "frozen", False):
            if messagebox.askyesno("Update available", "A new Windows version is available. Open its download page?"):
                webbrowser.open(app_updates.RELEASES_URL)
            return
        if messagebox.askyesno(
            "Update available",
            f"Version {update['version']} is available. Download and restart now?\n\n"
            "Generate any needed logs first. After restarting, choose your spreadsheet again.",
        ):
            self.status.set("Downloading and verifying update...")
            self.update_task(lambda: app_updates.download_update(update), self.install_update)

    def install_update(self, download) -> None:
        try:
            app_updates.start_install(download)
        except Exception as exc:
            messagebox.showerror("Could not install update", str(exc))
            return
        self.root.destroy()

    def _add_exam_row(self, index: int, label: str, optional: bool) -> None:
        year_var = StringVar()
        subject_var = StringVar(value=NO_ADDITIONAL_SUBJECT if optional else "All subjects")
        ttk.Label(self.exam_frame, text=label).grid(row=index, column=0, sticky="w", pady=4)
        year_box = ttk.Combobox(self.exam_frame, textvariable=year_var, state="readonly", width=10)
        year_box.grid(row=index, column=1, sticky="w", padx=(0, 8), pady=4)
        subject_box = ttk.Combobox(
            self.exam_frame, textvariable=subject_var, state="readonly"
        )
        subject_box.grid(row=index, column=2, sticky="ew", pady=4)
        year_box.bind("<<ComboboxSelected>>", lambda _event, row=index: self.exam_year_changed(row))
        subject_box.bind("<<ComboboxSelected>>", lambda _event: self.update_preview())
        self.exam_rows.append(
            {
                "optional": optional,
                "year": year_var,
                "subject": subject_var,
                "year_box": year_box,
                "subject_box": subject_box,
            }
        )

    def add_exam(self) -> None:
        if not self.data or len(self.exam_rows) >= MAX_EXAMS:
            return
        index = len(self.exam_rows)
        self._add_exam_row(index, f"Exam {index + 1}", optional=True)
        row = self.exam_rows[index]
        years = sorted(self.data.years, key=lambda year: int(year.split()[1]))
        row["year_box"].configure(values=years)
        row["year"].set(self.exam_rows[index - 1]["year"].get() or years[0])
        self.update_exam_subjects(index)
        self.add_exam_button.grid_configure(row=len(self.exam_rows))
        self.set_exam_controls_enabled(True)
        self.root.minsize(700, 520 + (len(self.exam_rows) - 3) * 32)
        self.update_preview()

    def _file_row(self, parent, row: int, label: str, variable: StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(parent, text="Browse", command=command).grid(row=row, column=2, sticky="e", pady=4)

    def _choose_xlsx(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel workbooks", "*.xlsx")])
        if path:
            self.xlsx_path.set(path)
            self.load_workbook()

    def _choose_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def load_workbook(self) -> None:
        workbook_path = Path(self.xlsx_path.get().strip())
        if not workbook_path:
            self.data = None
            self.set_exam_controls_enabled(False)
            self.update_preview()
            messagebox.showerror(
                "Choose spreadsheet",
                "Choose the access arrangements spreadsheet first.",
            )
            self.status.set("Choose the access arrangements spreadsheet to begin.")
            return
        if not workbook_path.exists():
            self.data = None
            self.set_exam_controls_enabled(False)
            self.update_preview()
            messagebox.showerror("Missing spreadsheet", "The selected spreadsheet was not found.")
            self.status.set("Spreadsheet could not be loaded.")
            return
        try:
            self.data = ArrangementData(workbook_path)
        except Exception as exc:
            self.data = None
            self.set_exam_controls_enabled(False)
            self.update_preview()
            messagebox.showerror("Could not load spreadsheet", str(exc))
            self.status.set("Spreadsheet could not be loaded.")
            return

        years = sorted(self.data.years, key=lambda y: int(y.split()[1]))
        if not years:
            self.data = None
            self.set_exam_controls_enabled(False)
            self.update_preview()
            messagebox.showerror("No year sheets found", "No usable Year 10, 11, 12 or 13 sheets were found.")
            self.status.set("Spreadsheet could not be loaded.")
            return
        defaults = ["Year 11", "Year 12", "Year 13"]
        for index, row in enumerate(self.exam_rows):
            row["year_box"].configure(values=years)
            current_year = row["year"].get()
            default_year = defaults[index] if index < len(defaults) and defaults[index] in years else years[0]
            row["year"].set(current_year if current_year in years else default_year)
            if row["optional"]:
                row["subject"].set(NO_ADDITIONAL_SUBJECT)
            elif row["subject"].get() == NO_ADDITIONAL_SUBJECT:
                row["subject"].set("All subjects")
            self.update_exam_subjects(index)
        self.set_exam_controls_enabled(True)
        self.update_preview()
        self.status.set(f"Loaded {len(self.data.years)} year groups from the spreadsheet.")

    def set_exam_controls_enabled(self, enabled: bool) -> None:
        self.add_exam_button.configure(
            state="normal" if enabled and len(self.exam_rows) < MAX_EXAMS else "disabled"
        )
        state = "readonly" if enabled else "disabled"
        for row in self.exam_rows:
            row["year_box"].configure(state=state)
            row["subject_box"].configure(state=state)
            if not enabled:
                row["year"].set("")
                row["subject"].set(NO_ADDITIONAL_SUBJECT if row["optional"] else "All subjects")

    def exam_year_changed(self, index: int) -> None:
        self.update_exam_subjects(index)
        self.update_preview()

    def update_exam_subjects(self, index: int) -> None:
        if not self.data:
            return
        row = self.exam_rows[index]
        year = row["year"].get()
        subjects = self.data.subjects_for_year(year, include_optional_blank=row["optional"])
        current = row["subject"].get()
        row["subject_box"].configure(values=subjects)
        row["subject"].set(current if current in subjects else subjects[0])

    def selected_exams(self) -> list[tuple[str, str]]:
        exams = []
        seen = set()
        for row in self.exam_rows:
            year = row["year"].get()
            subject = row["subject"].get()
            if not year or not subject or subject == NO_ADDITIONAL_SUBJECT:
                continue
            exam_key = (year, subject)
            if exam_key not in seen:
                exams.append(exam_key)
                seen.add(exam_key)
        return exams

    def selected_sections(self) -> tuple[ExamSection, ...]:
        if not self.data:
            return ()
        sections = {}
        order = []
        for year, subject in self.selected_exams():
            section_key = key(subject)
            if section_key not in sections:
                sections[section_key] = {"subject": subject, "students": {}}
                order.append(section_key)
            for student in self.data.students_for(year, subject):
                student_key = (key(student.name), key(student.arrangements), clean(student.room).casefold())
                sections[section_key]["students"][student_key] = student
        return tuple(
            ExamSection(
                sections[section_key]["subject"],
                tuple(
                    sorted(
                        sections[section_key]["students"].values(),
                        key=lambda item: item.name.lower(),
                    )
                ),
            )
            for section_key in order
        )

    def preview_text(self) -> str:
        if not self.data:
            return ""
        return "\n\n".join(
            room + "\n" + "\n\n".join("\n".join(output_lines(section)) for section in sections)
            for room, sections in sections_by_room(self.selected_sections()).items()
        )

    def update_preview(self) -> None:
        if not self.preview:
            return
        self.preview.set_text(self.preview_text())

    def copy_preview(self) -> None:
        text = self.preview_text()
        if not text:
            messagebox.showerror("Nothing to copy", "Choose at least one exam group first.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status.set("Copied the list to the clipboard.")

    def open_output_folder(self) -> None:
        try:
            open_folder(Path(self.output_dir.get()))
        except Exception as exc:
            messagebox.showerror("Could not open folder", str(exc))

    def generate(self) -> None:
        if not self.data:
            messagebox.showerror("Missing data", "Load the spreadsheet first.")
            return
        sections = self.selected_sections()
        if not sections:
            messagebox.showerror("No exam groups selected", "Choose at least one exam group.")
            return

        rooms = sections_by_room(sections)
        if not rooms:
            messagebox.showinfo("No room logs", "No matching students outside Alternative room. No files were created.")
            return

        if not DEFAULT_DOCX.exists():
            messagebox.showerror(
                "Missing template",
                f"The bundled Word template was not found:\n{DEFAULT_DOCX}",
            )
            return

        writer = IncidentLogWriter(DEFAULT_DOCX)
        date_stamp = datetime.now().strftime("%d-%m-%y")
        used_names = set()
        for index, (room, room_sections) in enumerate(rooms.items(), 1):
            safe_room = re.sub(r"[^A-Za-z0-9]+", "_", room).strip("_")[:80] or "Room"
            if safe_room.casefold() in used_names:
                safe_room += f"_{index}"
            used_names.add(safe_room.casefold())
            output = Path(self.output_dir.get()) / f"Incident Log - {safe_room} - {date_stamp}.docx"
            writer.write_sections(output, room_sections)

        self.update_preview()
        self.status.set(f"Created {len(rooms)} room log(s): " + "; ".join(rooms))
        if messagebox.askyesno(
            "Done",
            f"Created {len(rooms)} room log(s).\n\nOpen the output folder now?",
        ):
            self.open_output_folder()

    def run(self) -> None:
        self.root.mainloop()


def main() -> int:
    if len(sys.argv) == 4 and sys.argv[1] == "--apply-update":
        try:
            app_updates.apply_update(sys.argv[2], sys.argv[3])
        except Exception as exc:
            root = Tk()
            root.withdraw()
            messagebox.showerror("Update could not finish", str(exc))
            root.destroy()
            return 1
        return 0
    try:
        App().run()
    except ImportError as exc:
        messagebox.showerror(
            "Python problem",
            f"This Python installation cannot load a required built-in module:\n{exc}",
        )
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
