from pathlib import Path
import argparse
import json
import sys
import os
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

SETTINGS_FILE = "project_bundle_settings.json"

EXCLUDED_DIRS = {
    ".git",
    ".venv",
    "venv",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".vscode",
    ".idea",
    "build",
    "dist",
    ".eggs",
    "Debug",
    "Release",
}

HEADER_BAR = "=" * 120
SUB_BAR = "-" * 120

CHECKED = "checked"
UNCHECKED = "unchecked"
PARTIAL = "partial"

STATE_SYMBOLS = {
    CHECKED: "☑",
    UNCHECKED: "☐",
    PARTIAL: "◩",
}


def get_file_size(path: Path) -> int:
    try:
        return path.stat().st_size
    except OSError:
        return 0


def format_bytes(size_bytes: int) -> str:
    size = float(max(size_bytes, 0))
    units = ["B", "KB", "MB", "GB", "TB"]

    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1024

    return f"{int(size_bytes)} B"


def sum_file_sizes(files: list[Path]) -> int:
    return sum(get_file_size(path) for path in files)


def get_directory_included_size(directory: Path, included_set: set[Path], scan_root: Path) -> int:
    total = 0
    for path in directory.rglob("*"):
        if is_excluded(path, scan_root):
            continue
        if path.is_file() and path in included_set:
            total += get_file_size(path)
    return total


def get_path_size_for_tree(path: Path, included_set: set[Path], scan_root: Path) -> int:
    if path.is_file():
        return get_file_size(path) if path in included_set else 0
    return get_directory_included_size(path, included_set, scan_root)


def get_display_name_with_size(path: Path, included_set: set[Path], scan_root: Path) -> str:
    name = path.name
    size = get_path_size_for_tree(path, included_set, scan_root)
    return f"{name} ({format_bytes(size)})"


def get_total_size_label(total_size: int) -> str:
    return f"{format_bytes(total_size)} ({total_size:,} bytes)"


def get_included_size_label(included_size: int) -> str:
    return f"{format_bytes(included_size)} ({included_size:,} bytes)"


def get_settings_path() -> Path:
    return Path(__file__).resolve().parent / SETTINGS_FILE


def load_settings() -> dict:
    settings_path = get_settings_path()
    if not settings_path.exists():
        return {}

    try:
        return json.loads(settings_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(
    input_folder: str,
    output_folder: str,
    output_filename: str,
    selected_extensions: list[str],
    ignored_paths: list[str] | None = None,
) -> None:
    settings_path = get_settings_path()
    data = {
        "input_folder": input_folder,
        "output_folder": output_folder,
        "output_filename": output_filename,
        "selected_extensions": sorted(selected_extensions),
        "ignored_paths": sorted(ignored_paths or []),
    }
    settings_path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def is_excluded(path: Path, scan_root: Path) -> bool:
    rel_parts = path.relative_to(scan_root).parts
    return any(part in EXCLUDED_DIRS for part in rel_parts)


def read_docx_text(path: Path) -> str:
    """Extract readable plain text from a .docx file using only the Python standard library.

    A .docx file is a ZIP package containing XML parts. Reading it as raw text produces
    unreadable PK/ZIP binary content, so this extracts the visible Word document text instead.
    """
    namespaces = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    def extract_paragraph_text(paragraph: ET.Element) -> str:
        pieces: list[str] = []
        for node in paragraph.iter():
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "t" and node.text:
                pieces.append(node.text)
            elif tag == "tab":
                pieces.append("\t")
            elif tag in {"br", "cr"}:
                pieces.append("\n")
        return "".join(pieces).strip()

    try:
        with zipfile.ZipFile(path) as zf:
            part_names = [
                "word/document.xml",
                *sorted(name for name in zf.namelist() if name.startswith("word/header") and name.endswith(".xml")),
                *sorted(name for name in zf.namelist() if name.startswith("word/footer") and name.endswith(".xml")),
            ]

            sections: list[str] = []
            for part_name in part_names:
                if part_name not in zf.namelist():
                    continue

                try:
                    xml_data = zf.read(part_name)
                    root = ET.fromstring(xml_data)
                except Exception as ex:
                    sections.append(f"[ERROR READING DOCX PART {part_name}: {ex}]")
                    continue

                paragraphs: list[str] = []
                for paragraph in root.findall(".//w:p", namespaces):
                    text = extract_paragraph_text(paragraph)
                    if text:
                        paragraphs.append(text)

                if paragraphs:
                    if part_name != "word/document.xml":
                        sections.append(f"[{part_name}]\n" + "\n".join(paragraphs))
                    else:
                        sections.append("\n".join(paragraphs))

            return "\n\n".join(sections).strip() or "[DOCX file contained no extractable text]"
    except zipfile.BadZipFile:
        return "[ERROR: File has .docx extension but is not a valid DOCX/ZIP package]"
    except Exception as ex:
        return f"[ERROR READING DOCX FILE: {ex}]"


def safe_read_text(path: Path) -> str:
    if normalize_extension(path) == ".docx":
        return read_docx_text(path)

    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
        except Exception as ex:
            return f"[ERROR READING FILE: {ex}]"
    return "[ERROR: Could not decode file]"


def normalize_extension(path: Path) -> str:
    ext = path.suffix.lower()
    return ext if ext else "[no extension]"


def rel_key(path: Path, scan_root: Path) -> str:
    if path == scan_root:
        return "."
    return path.relative_to(scan_root).as_posix()


def scan_files(scan_root: Path) -> list[Path]:
    files = []
    for path in scan_root.rglob("*"):
        if not path.is_file():
            continue
        if is_excluded(path, scan_root):
            continue
        files.append(path)

    files.sort(key=lambda p: str(p.relative_to(scan_root)).lower())
    return files


def collect_extension_counts(files: list[Path]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for path in files:
        ext = normalize_extension(path)
        counts[ext] = counts.get(ext, 0) + 1

    return dict(sorted(counts.items(), key=lambda kv: (kv[0] == "[no extension]", kv[0])))


def filter_files_by_extensions(files: list[Path], selected_extensions: set[str]) -> list[Path]:
    return [path for path in files if normalize_extension(path) in selected_extensions]


def filter_files_by_ignored_paths(files: list[Path], scan_root: Path, ignored_paths: set[str]) -> list[Path]:
    included = []
    for path in files:
        key = rel_key(path, scan_root)
        parts = key.split("/")
        ignored = False

        for i in range(1, len(parts) + 1):
            ancestor_key = "/".join(parts[:i])
            if ancestor_key in ignored_paths:
                ignored = True
                break

        if not ignored:
            included.append(path)

    return included


def build_tree(scan_root: Path, included_files: list[Path]) -> list[str]:
    included_set = set(included_files)
    lines = [get_display_name_with_size(scan_root, included_set, scan_root)]

    def directory_has_visible_content(directory: Path) -> bool:
        for path in directory.rglob("*"):
            if is_excluded(path, scan_root):
                continue
            if path.is_file() and path in included_set:
                return True
        return False

    def walk(directory: Path, prefix: str = "") -> None:
        visible_children: list[Path] = []

        for child in sorted(directory.iterdir(), key=lambda p: (p.is_file(), p.name.lower())):
            if is_excluded(child, scan_root):
                continue

            if child.is_dir():
                if directory_has_visible_content(child):
                    visible_children.append(child)
            elif child in included_set:
                visible_children.append(child)

        for i, child in enumerate(visible_children):
            is_last = i == len(visible_children) - 1
            branch = "└── " if is_last else "├── "
            lines.append(prefix + branch + get_display_name_with_size(child, included_set, scan_root))

            if child.is_dir():
                extension = "    " if is_last else "│   "
                walk(child, prefix + extension)

    walk(scan_root)
    return lines



def make_recommended_output_filename(input_folder_str: str) -> str:
    """Create a safe default output filename from the selected input folder."""
    folder_name = Path(input_folder_str).resolve().name if input_folder_str.strip() else "project_bundle"
    safe_chars = []

    for char in folder_name.strip():
        if char.isalnum() or char in ("-", "_"):
            safe_chars.append(char)
        elif char.isspace():
            safe_chars.append("_")

    safe_name = "".join(safe_chars).strip("._-")
    if not safe_name:
        safe_name = "project_bundle"

    return f"{safe_name}.txt"

def build_output_path(output_folder_str: str, output_filename_str: str) -> Path:
    output_folder = Path(output_folder_str).resolve()
    filename = output_filename_str.strip()

    if not filename:
        filename = "project_bundle.txt"

    if not filename.lower().endswith(".txt"):
        filename += ".txt"

    return output_folder / filename


def write_bundle(
    scan_root: Path,
    all_files: list[Path],
    included_files: list[Path],
    selected_extensions: list[str],
    extension_counts: dict[str, int],
    output_path: Path,
    ignored_paths: list[str] | None = None,
) -> None:
    with output_path.open("w", encoding="utf-8", newline="\n") as f:
        f.write("PROJECT TEXT BUNDLE\n")
        f.write(f"Scan Root: {scan_root}\n")
        f.write(f"Output File: {output_path}\n")
        total_scanned_size = sum_file_sizes(all_files)
        total_included_size = sum_file_sizes(included_files)
        f.write(f"Total scanned files: {len(all_files)}\n")
        f.write(f"Total scanned file size: {get_total_size_label(total_scanned_size)}\n")
        f.write(f"Total included files: {len(included_files)}\n")
        f.write(f"Total included file size: {get_included_size_label(total_included_size)}\n")
        f.write("\n")

        f.write("EXCLUDED DIRECTORIES\n")
        f.write(SUB_BAR + "\n")
        for item in sorted(EXCLUDED_DIRS):
            f.write(f"- {item}\n")
        f.write("\n")

        f.write("DISCOVERED EXTENSIONS\n")
        f.write(SUB_BAR + "\n")
        for ext, count in extension_counts.items():
            f.write(f"{ext}: {count}\n")
        f.write("\n")

        f.write("SELECTED EXTENSIONS\n")
        f.write(SUB_BAR + "\n")
        for ext in selected_extensions:
            f.write(f"- {ext}\n")
        f.write("\n")

        if any(normalize_extension(path) == ".docx" for path in included_files):
            f.write("DOCUMENT EXTRACTION NOTES\n")
            f.write(SUB_BAR + "\n")
            f.write("- .docx files are ZIP/XML packages, so their visible document text is extracted instead of raw binary content.\n")
            f.write("- Formatting, images, and complex Word layout are not preserved in this text bundle.\n")
            f.write("\n")

        if ignored_paths:
            f.write("IGNORED TREE PATHS\n")
            f.write(SUB_BAR + "\n")
            for item in sorted(ignored_paths):
                f.write(f"- {item}\n")
            f.write("\n")

        f.write("PROJECT TREE\n")
        f.write(SUB_BAR + "\n")
        for line in build_tree(scan_root, included_files):
            f.write(line + "\n")
        f.write("\n")

        f.write("FILE INDEX\n")
        f.write(SUB_BAR + "\n")
        for i, path in enumerate(included_files, start=1):
            rel = path.relative_to(scan_root)
            f.write(f"{i:04d} | {format_bytes(get_file_size(path)):>10} | {rel}\n")
        f.write("\n\n")

        for i, path in enumerate(included_files, start=1):
            rel = path.relative_to(scan_root)
            content = safe_read_text(path)

            f.write(HEADER_BAR + "\n")
            f.write(f"FILE {i:04d}: {rel}\n")
            f.write(f"FULL PATH: {path}\n")
            f.write(f"TYPE: {normalize_extension(path)}\n")
            f.write(f"SIZE: {get_total_size_label(get_file_size(path))}\n")
            f.write(HEADER_BAR + "\n")
            f.write(content)

            if not content.endswith("\n"):
                f.write("\n")

            f.write(HEADER_BAR + "\n")
            f.write(f"END FILE {i:04d}: {rel}\n")
            f.write(HEADER_BAR + "\n\n")


def bundle_project(
    input_folder_str: str,
    output_folder_str: str,
    output_filename_str: str,
    selected_extensions: list[str],
    ignored_paths: list[str] | None = None,
) -> tuple[bool, str]:
    try:
        scan_root = Path(input_folder_str).resolve()
        output_folder = Path(output_folder_str).resolve()
        output_path = build_output_path(output_folder_str, output_filename_str)

        if not scan_root.exists() or not scan_root.is_dir():
            return False, f"Input folder does not exist or is not a folder:\n{scan_root}"

        if not output_folder.exists() or not output_folder.is_dir():
            return False, f"Output folder does not exist or is not a folder:\n{output_folder}"

        if not selected_extensions:
            return False, "No file extensions were selected."

        all_files = scan_files(scan_root)
        extension_counts = collect_extension_counts(all_files)
        included_files = filter_files_by_extensions(all_files, set(selected_extensions))
        included_files = filter_files_by_ignored_paths(included_files, scan_root, set(ignored_paths or []))

        if not included_files:
            return False, "No files matched the selected extensions and tree selection."

        write_bundle(
            scan_root=scan_root,
            all_files=all_files,
            included_files=included_files,
            selected_extensions=sorted(selected_extensions),
            extension_counts=extension_counts,
            output_path=output_path,
            ignored_paths=ignored_paths or [],
        )

        save_settings(str(scan_root), str(output_folder), output_path.name, selected_extensions, ignored_paths or [])

        return True, (
            f"Bundle created successfully.\n\n"
            f"Scan root: {scan_root}\n"
            f"Output folder: {output_folder}\n"
            f"Output file: {output_path.name}\n"
            f"Full output path: {output_path}\n"
            f"Included files: {len(included_files)}\n"
            f"Included file size: {get_included_size_label(sum_file_sizes(included_files))}"
        )
    except Exception as ex:
        return False, f"Unexpected error:\n{ex}"


def browse_input_folder(entry: tk.Entry) -> None:
    folder = filedialog.askdirectory(title="Select Input Folder")
    if folder:
        entry.delete(0, tk.END)
        entry.insert(0, folder)


def browse_output_folder(entry: tk.Entry) -> None:
    folder = filedialog.askdirectory(title="Select Output Folder")
    if folder:
        entry.delete(0, tk.END)
        entry.insert(0, folder)


def open_folder_in_file_explorer(folder: Path) -> tuple[bool, str]:
    """Open a folder in the operating system's file explorer."""
    try:
        folder = folder.resolve()

        if not folder.exists() or not folder.is_dir():
            return False, f"Output folder does not exist or is not a folder:\n{folder}"

        if sys.platform.startswith("win"):
            os.startfile(str(folder))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])

        return True, f"Opened output folder:\n{folder}"
    except Exception as ex:
        return False, f"Could not open output folder:\n{folder}\n\n{ex}"


class App:
    def __init__(self, root_window: tk.Tk):
        self.root_window = root_window
        self.settings = load_settings()

        self.extension_vars: dict[str, tk.BooleanVar] = {}
        self.scanned_files: list[Path] = []
        self.extension_counts: dict[str, int] = {}
        self.scan_root: Path | None = None

        self.item_by_rel: dict[str, str] = {}
        self.rel_by_item: dict[str, str] = {}
        self.path_by_rel: dict[str, Path] = {}
        self.state_by_rel: dict[str, str] = {}
        self.total_size_by_rel: dict[str, int] = {}
        self.included_size_by_rel: dict[str, int] = {}

        self.root_window.title("Project Bundle Utility")
        self.root_window.geometry("1450x850")
        self.root_window.resizable(True, True)

        self.build_ui()
        self.load_initial_values()

    def build_ui(self) -> None:
        pad_x = 12
        pad_y = 10

        tk.Label(self.root_window, text="Input Folder:").grid(
            row=0, column=0, padx=pad_x, pady=(pad_y, 4), sticky="w"
        )

        self.input_entry = tk.Entry(self.root_window, width=100)
        self.input_entry.grid(row=1, column=0, padx=pad_x, pady=4, sticky="we")

        tk.Button(
            self.root_window,
            text="Browse...",
            width=12,
            command=self.browse_input_folder,
        ).grid(row=1, column=1, padx=(0, pad_x), pady=4)

        tk.Label(self.root_window, text="Output Folder:").grid(
            row=2, column=0, padx=pad_x, pady=(pad_y, 4), sticky="w"
        )

        self.output_folder_entry = tk.Entry(self.root_window, width=100)
        self.output_folder_entry.grid(row=3, column=0, padx=pad_x, pady=4, sticky="we")

        tk.Button(
            self.root_window,
            text="Browse...",
            width=12,
            command=lambda: browse_output_folder(self.output_folder_entry),
        ).grid(row=3, column=1, padx=(0, pad_x), pady=4)

        tk.Label(self.root_window, text="Output File Name:").grid(
            row=4, column=0, padx=pad_x, pady=(pad_y, 4), sticky="w"
        )

        self.output_filename_entry = tk.Entry(self.root_window, width=100)
        self.output_filename_entry.grid(row=5, column=0, padx=pad_x, pady=4, sticky="we")

        button_frame = tk.Frame(self.root_window)
        button_frame.grid(row=6, column=0, columnspan=2, padx=pad_x, pady=(14, 6), sticky="w")

        tk.Button(button_frame, text="Scan Folder", width=16, command=self.on_scan).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Select All Ext", width=16, command=self.select_all_extensions).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Clear All Ext", width=16, command=self.clear_all_extensions).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Include All Tree", width=16, command=self.include_all_tree).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Ignore All Tree", width=16, command=self.ignore_all_tree).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Create Bundle", width=16, command=self.on_bundle).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Open Output Folder", width=20, command=self.open_output_folder).pack(side="left")

        self.main_paned = ttk.PanedWindow(self.root_window, orient="horizontal")
        self.main_paned.grid(row=7, column=0, columnspan=2, padx=pad_x, pady=(8, 4), sticky="nsew")

        extensions_outer = ttk.Frame(self.main_paned)
        tree_outer = ttk.Frame(self.main_paned)

        self.main_paned.add(extensions_outer, weight=1)
        self.main_paned.add(tree_outer, weight=5)

        tk.Label(extensions_outer, text="Discovered Extensions:").pack(anchor="w", pady=(0, 4))

        self.extensions_container = tk.Frame(extensions_outer, bd=1, relief="sunken")
        self.extensions_container.pack(fill="both", expand=True)

        self.extensions_canvas = tk.Canvas(self.extensions_container, width=220)
        self.extensions_scrollbar = tk.Scrollbar(self.extensions_container, orient="vertical", command=self.extensions_canvas.yview)
        self.extensions_inner = tk.Frame(self.extensions_canvas)

        self.extensions_inner.bind(
            "<Configure>",
            lambda e: self.extensions_canvas.configure(scrollregion=self.extensions_canvas.bbox("all")),
        )

        self.extensions_canvas.create_window((0, 0), window=self.extensions_inner, anchor="nw")
        self.extensions_canvas.configure(yscrollcommand=self.extensions_scrollbar.set)

        self.extensions_canvas.pack(side="left", fill="both", expand=True)
        self.extensions_scrollbar.pack(side="right", fill="y")

        tk.Label(
            tree_outer,
            text="Project Tree Include / Ignore Selection:"
        ).pack(anchor="w", pady=(0, 4))

        tree_help = (
            "Click the checkbox symbol/text or press Space/Enter on a selected row to toggle it. "
            "Use the + / - expander only to open or close folders. "
            "Changing extension checkboxes rebuilds this tree. "
            "◩ means some children are included and some are ignored."
        )
        tk.Label(tree_outer, text=tree_help, anchor="w", justify="left").pack(anchor="w", pady=(0, 6))

        tree_frame = tk.Frame(tree_outer, bd=1, relief="sunken")
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("type", "extension", "included_size", "total_size"),
            show="tree headings",
            selectmode="browse",
        )
        self.tree.heading("#0", text="Include / Path")
        self.tree.heading("type", text="Type")
        self.tree.heading("extension", text="Extension")
        self.tree.heading("included_size", text="Included Size")
        self.tree.heading("total_size", text="Total Size")
        self.tree.column("#0", width=650, minwidth=320, stretch=True)
        self.tree.column("type", width=75, minwidth=70, stretch=False)
        self.tree.column("extension", width=95, minwidth=80, stretch=False)
        self.tree.column("included_size", width=135, minwidth=120, stretch=False, anchor="e")
        self.tree.column("total_size", width=135, minwidth=120, stretch=False, anchor="e")

        y_scrollbar = tk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        x_scrollbar = tk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<space>", self.on_tree_keyboard_toggle)
        self.tree.bind("<Return>", self.on_tree_keyboard_toggle)

        self.status_label = tk.Label(self.root_window, text="", anchor="w", justify="left")
        self.status_label.grid(row=8, column=0, columnspan=2, padx=pad_x, pady=(8, 10), sticky="w")

        self.root_window.columnconfigure(0, weight=1)
        self.root_window.rowconfigure(7, weight=1)
        self.root_window.after(100, self.set_initial_pane_position)
        self.root_window.protocol("WM_DELETE_WINDOW", self.on_close)

    def browse_input_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, folder)
            self.update_recommended_output_filename(folder)

    def update_recommended_output_filename(self, input_folder: str | None = None) -> None:
        input_folder = input_folder if input_folder is not None else self.input_entry.get().strip()
        recommended_name = make_recommended_output_filename(input_folder)
        self.output_filename_entry.delete(0, tk.END)
        self.output_filename_entry.insert(0, recommended_name)

    def set_initial_pane_position(self) -> None:
        try:
            self.main_paned.sashpos(0, 260)
        except tk.TclError:
            pass

    def load_initial_values(self) -> None:
        default_input = self.settings.get("input_folder", "")
        default_output_folder = self.settings.get("output_folder", "")
        default_output_filename = self.settings.get("output_filename", "")

        if not default_output_filename and default_input:
            default_output_filename = make_recommended_output_filename(default_input)
        elif not default_output_filename:
            default_output_filename = "project_bundle.txt"

        self.input_entry.insert(0, default_input)
        self.output_folder_entry.insert(0, default_output_folder)
        self.output_filename_entry.insert(0, default_output_filename)

        if default_input:
            self.status_label.config(text="Click 'Scan Folder' to discover extensions and build the project tree.")

    def rebuild_extension_checkboxes(self) -> None:
        for widget in self.extensions_inner.winfo_children():
            widget.destroy()

        self.extension_vars.clear()

        saved_selected = set(self.settings.get("selected_extensions", []))

        if not self.extension_counts:
            tk.Label(self.extensions_inner, text="No scan results yet.").grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return

        row = 0
        for ext, count in self.extension_counts.items():
            checked = ext in saved_selected
            var = tk.BooleanVar(value=checked)
            self.extension_vars[ext] = var

            cb = tk.Checkbutton(
                self.extensions_inner,
                text=f"{ext} ({count})",
                variable=var,
                anchor="w",
                justify="left",
                command=self.on_extension_selection_changed,
            )
            cb.grid(row=row, column=0, sticky="w", padx=8, pady=2)
            row += 1

    def on_extension_selection_changed(self) -> None:
        if self.scan_root is None:
            return

        current_ignored_paths = set(self.get_ignored_paths()) if self.item_by_rel else set(self.settings.get("ignored_paths", []))
        self.rebuild_file_tree(ignored_paths=current_ignored_paths)

        selected_count = len(self.get_selected_extensions())
        visible_files = self.get_files_matching_selected_extensions()
        visible_size = sum_file_sizes(visible_files)
        self.status_label.config(
            text=(
                f"Extension selection updated. Showing {len(visible_files)} files "
                f"({format_bytes(visible_size)}) across {selected_count} selected extension groups."
            )
        )

    def get_files_matching_selected_extensions(self) -> list[Path]:
        selected_extensions = set(self.get_selected_extensions())
        if not selected_extensions:
            return []
        return filter_files_by_extensions(self.scanned_files, selected_extensions)

    def rebuild_file_tree(self, ignored_paths: set[str] | None = None) -> None:
        self.tree.delete(*self.tree.get_children())

        self.item_by_rel.clear()
        self.rel_by_item.clear()
        self.path_by_rel.clear()
        self.state_by_rel.clear()
        self.total_size_by_rel.clear()
        self.included_size_by_rel.clear()

        if self.scan_root is None:
            return

        saved_ignored = set(self.settings.get("ignored_paths", [])) if ignored_paths is None else set(ignored_paths)
        visible_files = self.get_files_matching_selected_extensions()

        root_key = "."
        root_item = self.tree.insert(
            "",
            "end",
            text=f"{STATE_SYMBOLS[CHECKED]} {self.scan_root.name}",
            values=("folder", "", "", ""),
            open=True,
        )

        self.item_by_rel[root_key] = root_item
        self.rel_by_item[root_item] = root_key
        self.path_by_rel[root_key] = self.scan_root
        self.state_by_rel[root_key] = CHECKED

        # Insert folders and files needed to represent every file allowed by the selected extensions.
        # Files whose extensions are unchecked are intentionally hidden from the tree and cannot be bundled.
        for file_path in visible_files:
            parts = file_path.relative_to(self.scan_root).parts
            parent_item = root_item
            parent_path = self.scan_root
            accumulated_parts: list[str] = []

            for part in parts[:-1]:
                accumulated_parts.append(part)
                folder_key = "/".join(accumulated_parts)
                folder_path = parent_path / part

                if folder_key not in self.item_by_rel:
                    item = self.tree.insert(
                        parent_item,
                        "end",
                        text=f"{STATE_SYMBOLS[CHECKED]} {part}",
                        values=("folder", "", "", ""),
                        open=False,
                    )
                    self.item_by_rel[folder_key] = item
                    self.rel_by_item[item] = folder_key
                    self.path_by_rel[folder_key] = folder_path
                    self.state_by_rel[folder_key] = CHECKED

                parent_item = self.item_by_rel[folder_key]
                parent_path = folder_path

            file_key = file_path.relative_to(self.scan_root).as_posix()
            if file_key not in self.item_by_rel:
                item = self.tree.insert(
                    parent_item,
                    "end",
                    text=f"{STATE_SYMBOLS[CHECKED]} {file_path.name}",
                    values=("file", normalize_extension(file_path), "", ""),
                    open=False,
                )
                self.item_by_rel[file_key] = item
                self.rel_by_item[item] = file_key
                self.path_by_rel[file_key] = file_path
                self.state_by_rel[file_key] = CHECKED

        # Apply ignored paths from the previous session.
        for ignored_key in sorted(saved_ignored, key=lambda x: x.count("/")):
            if ignored_key in self.item_by_rel:
                self.set_subtree_state(ignored_key, UNCHECKED, refresh=False)

        self.refresh_all_parent_states()
        self.refresh_all_tree_sizes()
        self.refresh_all_tree_labels()

    def display_name_for_key(self, key: str) -> str:
        path = self.path_by_rel[key]
        return path.name if key != "." else path.name

    def set_subtree_state(self, key: str, state: str, refresh: bool = True) -> None:
        item = self.item_by_rel[key]
        self.state_by_rel[key] = state

        for child_item in self.tree.get_children(item):
            child_key = self.rel_by_item[child_item]
            self.set_subtree_state(child_key, state, refresh=False)

        if refresh:
            self.refresh_parent_states_from_key(key)
            self.refresh_all_tree_sizes()
            self.refresh_all_tree_labels()

    def refresh_parent_states_from_key(self, key: str) -> None:
        item = self.item_by_rel[key]
        parent = self.tree.parent(item)

        while parent:
            parent_key = self.rel_by_item[parent]
            child_states = [self.state_by_rel[self.rel_by_item[child]] for child in self.tree.get_children(parent)]

            if child_states and all(state == CHECKED for state in child_states):
                self.state_by_rel[parent_key] = CHECKED
            elif child_states and all(state == UNCHECKED for state in child_states):
                self.state_by_rel[parent_key] = UNCHECKED
            else:
                self.state_by_rel[parent_key] = PARTIAL

            parent = self.tree.parent(parent)

    def tree_key_depth(self, key: str) -> int:
        if key == ".":
            return 0
        return len(key.split("/"))

    def refresh_all_parent_states(self) -> None:
        keys_by_depth = sorted(self.item_by_rel.keys(), key=self.tree_key_depth, reverse=True)

        for key in keys_by_depth:
            item = self.item_by_rel[key]
            children = self.tree.get_children(item)
            if not children:
                continue

            child_states = [self.state_by_rel[self.rel_by_item[child]] for child in children]
            if all(state == CHECKED for state in child_states):
                self.state_by_rel[key] = CHECKED
            elif all(state == UNCHECKED for state in child_states):
                self.state_by_rel[key] = UNCHECKED
            else:
                self.state_by_rel[key] = PARTIAL

    def refresh_all_tree_sizes(self) -> None:
        keys_by_depth = sorted(self.item_by_rel.keys(), key=self.tree_key_depth, reverse=True)

        for key in keys_by_depth:
            item = self.item_by_rel[key]
            path = self.path_by_rel[key]
            children = self.tree.get_children(item)

            if path.is_file():
                total_size = get_file_size(path)
                included_size = total_size if self.state_by_rel.get(key, CHECKED) == CHECKED else 0
            else:
                total_size = sum(self.total_size_by_rel.get(self.rel_by_item[child], 0) for child in children)
                included_size = sum(self.included_size_by_rel.get(self.rel_by_item[child], 0) for child in children)

            self.total_size_by_rel[key] = total_size
            self.included_size_by_rel[key] = included_size

    def refresh_all_tree_labels(self) -> None:
        for key, item in self.item_by_rel.items():
            state = self.state_by_rel.get(key, CHECKED)
            current_values = list(self.tree.item(item, "values"))
            while len(current_values) < 4:
                current_values.append("")

            current_values[2] = format_bytes(self.included_size_by_rel.get(key, 0))
            current_values[3] = format_bytes(self.total_size_by_rel.get(key, 0))

            self.tree.item(
                item,
                text=f"{STATE_SYMBOLS[state]} {self.display_name_for_key(key)}",
                values=tuple(current_values),
            )

    def toggle_tree_key(self, key: str) -> None:
        current = self.state_by_rel.get(key, CHECKED)
        new_state = UNCHECKED if current == CHECKED else CHECKED
        self.set_subtree_state(key, new_state)

    def on_tree_click(self, event: tk.Event) -> str | None:
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        element_id = self.tree.identify_element(event.x, event.y)

        if not row_id:
            return None

        # Treeview uses the same #0 column for the expander and the text.
        # Only toggle when the actual item text/check symbol is clicked.
        # Let clicks on the expander indicator fall through so + / - only expands/collapses.
        if column_id == "#0" and element_id == "text":
            key = self.rel_by_item[row_id]
            self.toggle_tree_key(key)
            return "break"

        return None

    def on_tree_keyboard_toggle(self, event: tk.Event) -> str:
        selected = self.tree.selection()
        if selected:
            key = self.rel_by_item[selected[0]]
            self.toggle_tree_key(key)
        return "break"

    def include_all_tree(self) -> None:
        if "." in self.item_by_rel:
            self.set_subtree_state(".", CHECKED)

    def ignore_all_tree(self) -> None:
        if "." in self.item_by_rel:
            self.set_subtree_state(".", UNCHECKED)

    def on_scan(self) -> None:
        input_folder = self.input_entry.get().strip()

        if not input_folder:
            messagebox.showerror("Missing Input", "Please select an input folder.")
            return

        scan_root = Path(input_folder).resolve()
        if not scan_root.exists() or not scan_root.is_dir():
            messagebox.showerror("Invalid Input", f"Input folder does not exist or is not a folder:\n{scan_root}")
            return

        self.update_recommended_output_filename(str(scan_root))
        self.status_label.config(text="Scanning folder recursively...")
        self.root_window.update_idletasks()

        try:
            self.scan_root = scan_root
            self.scanned_files = scan_files(scan_root)
            self.extension_counts = collect_extension_counts(self.scanned_files)
            self.rebuild_extension_checkboxes()
            self.rebuild_file_tree()

            if not self.scanned_files:
                self.status_label.config(text="Scan complete. No files found after exclusions.")
                messagebox.showinfo("Scan Complete", "No files were found after applying excluded folders.")
                return

            total_size = sum_file_sizes(self.scanned_files)
            self.status_label.config(
                text=(
                    f"Scan complete. Found {len(self.scanned_files)} files "
                    f"({format_bytes(total_size)}) across {len(self.extension_counts)} extension groups. "
                    "The tree shows only files matching the selected extensions."
                )
            )
        except Exception as ex:
            self.status_label.config(text="Scan failed.")
            messagebox.showerror("Error", f"Unexpected error while scanning:\n{ex}")

    def select_all_extensions(self) -> None:
        for var in self.extension_vars.values():
            var.set(True)
        self.on_extension_selection_changed()

    def clear_all_extensions(self) -> None:
        for var in self.extension_vars.values():
            var.set(False)
        self.on_extension_selection_changed()

    def get_selected_extensions(self) -> list[str]:
        return sorted([ext for ext, var in self.extension_vars.items() if var.get()])

    def get_ignored_paths(self) -> list[str]:
        ignored: list[str] = []

        for key, state in self.state_by_rel.items():
            if key == ".":
                continue
            if state != UNCHECKED:
                continue

            parent_key = "/".join(key.split("/")[:-1])
            if not parent_key:
                parent_key = "."

            # If the parent is already fully ignored, storing this child is redundant.
            if self.state_by_rel.get(parent_key) == UNCHECKED:
                continue

            ignored.append(key)

        return sorted(ignored)

    def get_tree_included_files(self) -> list[Path]:
        if self.scan_root is None:
            return []

        ignored_paths = set(self.get_ignored_paths())
        extension_filtered_files = self.get_files_matching_selected_extensions()
        return filter_files_by_ignored_paths(extension_filtered_files, self.scan_root, ignored_paths)

    def on_bundle(self) -> None:
        input_folder = self.input_entry.get().strip()
        output_folder = self.output_folder_entry.get().strip()
        output_filename = self.output_filename_entry.get().strip()

        if not input_folder:
            messagebox.showerror("Missing Input", "Please select an input folder.")
            return

        if not output_folder:
            messagebox.showerror("Missing Output", "Please select an output folder.")
            return

        if not self.scanned_files or self.scan_root is None:
            messagebox.showerror("Missing Scan", "Please scan the input folder before creating a bundle.")
            return

        if not output_filename:
            output_filename = "project_bundle.txt"
            self.output_filename_entry.delete(0, tk.END)
            self.output_filename_entry.insert(0, output_filename)

        selected_extensions = self.get_selected_extensions()
        if not selected_extensions:
            messagebox.showerror("Missing Selection", "Please select at least one file extension.")
            return

        ignored_paths = self.get_ignored_paths()

        self.status_label.config(text="Creating bundle...")
        self.root_window.update_idletasks()

        ok, message = bundle_project(input_folder, output_folder, output_filename, selected_extensions, ignored_paths)
        if ok:
            self.settings["selected_extensions"] = selected_extensions
            self.settings["ignored_paths"] = ignored_paths
            self.status_label.config(text="Bundle complete.")
            messagebox.showinfo("Success", message)
        else:
            self.status_label.config(text="Bundle failed.")
            messagebox.showerror("Error", message)

    def open_output_folder(self) -> None:
        output_folder = self.output_folder_entry.get().strip()

        if not output_folder:
            messagebox.showerror("Missing Output", "Please select an output folder first.")
            return

        ok, message = open_folder_in_file_explorer(Path(output_folder))
        if ok:
            self.status_label.config(text=message.replace("\n", " "))
        else:
            self.status_label.config(text="Could not open output folder.")
            messagebox.showerror("Open Output Folder", message)

    def on_close(self) -> None:
        try:
            save_settings(
                self.input_entry.get().strip(),
                self.output_folder_entry.get().strip(),
                self.output_filename_entry.get().strip(),
                self.get_selected_extensions(),
                self.get_ignored_paths(),
            )
        except Exception:
            pass

        self.root_window.destroy()


def run_gui() -> None:
    root_window = tk.Tk()
    App(root_window)
    root_window.mainloop()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scan a folder recursively, choose file extensions, and bundle matching files into one text file."
    )
    parser.add_argument(
        "input_folder",
        nargs="?",
        help="Path to the folder to scan recursively.",
    )
    parser.add_argument(
        "-d",
        "--output-folder",
        help="Output folder path.",
    )
    parser.add_argument(
        "-n",
        "--output-name",
        help="Output file name.",
    )
    parser.add_argument(
        "-e",
        "--extensions",
        nargs="+",
        help='Extensions to include, for example: .c .h .py .json',
    )
    parser.add_argument(
        "--ignore",
        nargs="*",
        default=[],
        help="Relative files or folders to ignore, for example: tests data/generated.json",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if len(sys.argv) == 1:
        run_gui()
        return

    if not args.input_folder:
        raise SystemExit("Error: input_folder is required in CLI mode.")

    output_folder = args.output_folder or "."
    output_name = args.output_name or "project_bundle.txt"
    selected_extensions = args.extensions or []

    if not selected_extensions:
        raise SystemExit("Error: at least one extension must be supplied in CLI mode using -e")

    normalized_extensions = []
    for ext in selected_extensions:
        ext = ext.strip().lower()
        if ext != "[no extension]" and not ext.startswith("."):
            ext = "." + ext
        normalized_extensions.append(ext)

    ignored_paths = [item.strip().replace("\\", "/").strip("/") for item in args.ignore if item.strip()]

    ok, message = bundle_project(args.input_folder, output_folder, output_name, normalized_extensions, ignored_paths)
    if not ok:
        raise SystemExit(message)

    print(message)


if __name__ == "__main__":
    main()
