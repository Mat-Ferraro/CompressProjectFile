from pathlib import Path
import argparse
import json
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

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


def save_settings(input_folder: str, output_folder: str, output_filename: str, selected_extensions: list[str]) -> None:
    settings_path = get_settings_path()
    data = {
        "input_folder": input_folder,
        "output_folder": output_folder,
        "output_filename": output_filename,
        "selected_extensions": sorted(selected_extensions),
    }
    settings_path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def is_excluded(path: Path, scan_root: Path) -> bool:
    rel_parts = path.relative_to(scan_root).parts
    return any(part in EXCLUDED_DIRS for part in rel_parts)


def safe_read_text(path: Path) -> str:
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


def build_tree(scan_root: Path, selected_extensions: set[str]) -> list[str]:
    lines = [scan_root.name]

    def walk(directory: Path, prefix: str = "") -> bool:
        visible_children: list[Path] = []

        for child in sorted(directory.iterdir(), key=lambda p: (p.is_file(), p.name.lower())):
            if is_excluded(child, scan_root):
                continue

            if child.is_dir():
                if directory_has_visible_content(child, scan_root, selected_extensions):
                    visible_children.append(child)
            elif normalize_extension(child) in selected_extensions:
                visible_children.append(child)

        for i, child in enumerate(visible_children):
            is_last = i == len(visible_children) - 1
            branch = "└── " if is_last else "├── "
            lines.append(prefix + branch + child.name)

            if child.is_dir():
                extension = "    " if is_last else "│   "
                walk(child, prefix + extension)

        return bool(visible_children)

    walk(scan_root)
    return lines


def directory_has_visible_content(directory: Path, scan_root: Path, selected_extensions: set[str]) -> bool:
    for path in directory.rglob("*"):
        if is_excluded(path, scan_root):
            continue
        if path.is_file() and normalize_extension(path) in selected_extensions:
            return True
    return False


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
) -> None:
    with output_path.open("w", encoding="utf-8", newline="\n") as f:
        f.write("PROJECT TEXT BUNDLE\n")
        f.write(f"Scan Root: {scan_root}\n")
        f.write(f"Output File: {output_path}\n")
        f.write(f"Total scanned files: {len(all_files)}\n")
        f.write(f"Total included files: {len(included_files)}\n")
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

        f.write("PROJECT TREE\n")
        f.write(SUB_BAR + "\n")
        for line in build_tree(scan_root, set(selected_extensions)):
            f.write(line + "\n")
        f.write("\n")

        f.write("FILE INDEX\n")
        f.write(SUB_BAR + "\n")
        for i, path in enumerate(included_files, start=1):
            rel = path.relative_to(scan_root)
            f.write(f"{i:04d} | {rel}\n")
        f.write("\n\n")

        for i, path in enumerate(included_files, start=1):
            rel = path.relative_to(scan_root)
            content = safe_read_text(path)

            f.write(HEADER_BAR + "\n")
            f.write(f"FILE {i:04d}: {rel}\n")
            f.write(f"FULL PATH: {path}\n")
            f.write(f"TYPE: {normalize_extension(path)}\n")
            f.write(HEADER_BAR + "\n")
            f.write(content)

            if not content.endswith("\n"):
                f.write("\n")

            f.write(HEADER_BAR + "\n")
            f.write(f"END FILE {i:04d}: {rel}\n")
            f.write(HEADER_BAR + "\n\n")


def bundle_project(input_folder_str: str, output_folder_str: str, output_filename_str: str, selected_extensions: list[str]) -> tuple[bool, str]:
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

        if not included_files:
            return False, "No files matched the selected extensions."

        write_bundle(
            scan_root=scan_root,
            all_files=all_files,
            included_files=included_files,
            selected_extensions=sorted(selected_extensions),
            extension_counts=extension_counts,
            output_path=output_path,
        )

        save_settings(str(scan_root), str(output_folder), output_path.name, selected_extensions)

        return True, (
            f"Bundle created successfully.\n\n"
            f"Scan root: {scan_root}\n"
            f"Output folder: {output_folder}\n"
            f"Output file: {output_path.name}\n"
            f"Full output path: {output_path}\n"
            f"Included files: {len(included_files)}"
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


class App:
    def __init__(self, root_window: tk.Tk):
        self.root_window = root_window
        self.settings = load_settings()
        self.extension_vars: dict[str, tk.BooleanVar] = {}
        self.scanned_files: list[Path] = []
        self.extension_counts: dict[str, int] = {}

        self.root_window.title("Project Bundle Utility")
        self.root_window.geometry("900x720")
        self.root_window.resizable(True, True)

        self.build_ui()
        self.load_initial_values()

    def build_ui(self) -> None:
        pad_x = 12
        pad_y = 10

        tk.Label(self.root_window, text="Input Folder:").grid(
            row=0, column=0, padx=pad_x, pady=(pad_y, 4), sticky="w"
        )

        self.input_entry = tk.Entry(self.root_window, width=92)
        self.input_entry.grid(row=1, column=0, padx=pad_x, pady=4, sticky="we")

        tk.Button(
            self.root_window,
            text="Browse...",
            width=12,
            command=lambda: browse_input_folder(self.input_entry),
        ).grid(row=1, column=1, padx=(0, pad_x), pady=4)

        tk.Label(self.root_window, text="Output Folder:").grid(
            row=2, column=0, padx=pad_x, pady=(pad_y, 4), sticky="w"
        )

        self.output_folder_entry = tk.Entry(self.root_window, width=92)
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

        self.output_filename_entry = tk.Entry(self.root_window, width=92)
        self.output_filename_entry.grid(row=5, column=0, padx=pad_x, pady=4, sticky="we")

        button_frame = tk.Frame(self.root_window)
        button_frame.grid(row=6, column=0, columnspan=2, padx=pad_x, pady=(14, 6), sticky="w")

        tk.Button(button_frame, text="Scan Folder", width=16, command=self.on_scan).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Select All", width=16, command=self.select_all_extensions).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Clear All", width=16, command=self.clear_all_extensions).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Create Bundle", width=16, command=self.on_bundle).pack(side="left", padx=(0, 8))
        tk.Button(button_frame, text="Close", width=12, command=self.on_close).pack(side="left")

        tk.Label(self.root_window, text="Discovered Extensions:").grid(
            row=7, column=0, padx=pad_x, pady=(pad_y, 4), sticky="w"
        )

        self.extensions_container = tk.Frame(self.root_window, bd=1, relief="sunken")
        self.extensions_container.grid(row=8, column=0, columnspan=2, padx=pad_x, pady=4, sticky="nsew")

        self.extensions_canvas = tk.Canvas(self.extensions_container, height=360)
        self.extensions_scrollbar = tk.Scrollbar(self.extensions_container, orient="vertical", command=self.extensions_canvas.yview)
        self.extensions_inner = tk.Frame(self.extensions_canvas)

        self.extensions_inner.bind(
            "<Configure>",
            lambda e: self.extensions_canvas.configure(scrollregion=self.extensions_canvas.bbox("all"))
        )

        self.extensions_canvas.create_window((0, 0), window=self.extensions_inner, anchor="nw")
        self.extensions_canvas.configure(yscrollcommand=self.extensions_scrollbar.set)

        self.extensions_canvas.pack(side="left", fill="both", expand=True)
        self.extensions_scrollbar.pack(side="right", fill="y")

        self.status_label = tk.Label(self.root_window, text="", anchor="w", justify="left")
        self.status_label.grid(row=9, column=0, columnspan=2, padx=pad_x, pady=(8, 10), sticky="w")

        self.root_window.columnconfigure(0, weight=1)
        self.root_window.rowconfigure(8, weight=1)
        self.root_window.protocol("WM_DELETE_WINDOW", self.on_close)

    def load_initial_values(self) -> None:
        default_input = self.settings.get("input_folder", "")
        default_output_folder = self.settings.get("output_folder", "")
        default_output_filename = self.settings.get("output_filename", "project_bundle.txt")

        self.input_entry.insert(0, default_input)
        self.output_folder_entry.insert(0, default_output_folder)
        self.output_filename_entry.insert(0, default_output_filename)

        if default_input:
            self.status_label.config(text="Click 'Scan Folder' to discover extensions.")

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

            cb = tk.Checkbutton(self.extensions_inner, text=f"{ext} ({count})", variable=var, anchor="w", justify="left")
            cb.grid(row=row, column=0, sticky="w", padx=8, pady=2)
            row += 1

    def on_scan(self) -> None:
        input_folder = self.input_entry.get().strip()

        if not input_folder:
            messagebox.showerror("Missing Input", "Please select an input folder.")
            return

        scan_root = Path(input_folder).resolve()
        if not scan_root.exists() or not scan_root.is_dir():
            messagebox.showerror("Invalid Input", f"Input folder does not exist or is not a folder:\n{scan_root}")
            return

        self.status_label.config(text="Scanning folder recursively...")
        self.root_window.update_idletasks()

        try:
            self.scanned_files = scan_files(scan_root)
            self.extension_counts = collect_extension_counts(self.scanned_files)
            self.rebuild_extension_checkboxes()

            if not self.scanned_files:
                self.status_label.config(text="Scan complete. No files found after exclusions.")
                messagebox.showinfo("Scan Complete", "No files were found after applying excluded folders.")
                return

            self.status_label.config(
                text=f"Scan complete. Found {len(self.scanned_files)} files across {len(self.extension_counts)} extension groups."
            )
        except Exception as ex:
            self.status_label.config(text="Scan failed.")
            messagebox.showerror("Error", f"Unexpected error while scanning:\n{ex}")

    def select_all_extensions(self) -> None:
        for var in self.extension_vars.values():
            var.set(True)

    def clear_all_extensions(self) -> None:
        for var in self.extension_vars.values():
            var.set(False)

    def get_selected_extensions(self) -> list[str]:
        return sorted([ext for ext, var in self.extension_vars.items() if var.get()])

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

        if not output_filename:
            output_filename = "project_bundle.txt"
            self.output_filename_entry.delete(0, tk.END)
            self.output_filename_entry.insert(0, output_filename)

        selected_extensions = self.get_selected_extensions()
        if not selected_extensions:
            messagebox.showerror("Missing Selection", "Please scan first and select at least one file extension.")
            return

        self.status_label.config(text="Creating bundle...")
        self.root_window.update_idletasks()

        ok, message = bundle_project(input_folder, output_folder, output_filename, selected_extensions)
        if ok:
            self.settings["selected_extensions"] = selected_extensions
            self.status_label.config(text="Bundle complete.")
            messagebox.showinfo("Success", message)
        else:
            self.status_label.config(text="Bundle failed.")
            messagebox.showerror("Error", message)

    def on_close(self) -> None:
        try:
            save_settings(
                self.input_entry.get().strip(),
                self.output_folder_entry.get().strip(),
                self.output_filename_entry.get().strip(),
                self.get_selected_extensions(),
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
        help="Path to the folder to scan recursively."
    )
    parser.add_argument(
        "-d",
        "--output-folder",
        help="Output folder path."
    )
    parser.add_argument(
        "-n",
        "--output-name",
        help="Output file name."
    )
    parser.add_argument(
        "-e",
        "--extensions",
        nargs="+",
        help='Extensions to include, for example: .c .h .py .json'
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

    ok, message = bundle_project(args.input_folder, output_folder, output_name, normalized_extensions)
    if not ok:
        raise SystemExit(message)

    print(message)


if __name__ == "__main__":
    main()