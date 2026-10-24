import sys
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk


# Send to: make a pythonw.exe shortcut targeting this script and place it in Win+R > shell:sendto.
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
PDF_EXTENSION = ".pdf"
WORD_EXTENSIONS = {".doc", ".docx"}
EXCEL_EXTENSIONS = {".xls", ".xlsx"}


def convert_office_to_pdf(source: Path, output: Path) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Word and Excel conversion requires pywin32 and Microsoft Office."
        ) from exc

    pythoncom.CoInitialize()
    application = None
    document = None
    try:
        if source.suffix.lower() in WORD_EXTENSIONS:
            application = win32com.client.DispatchEx("Word.Application")
            application.Visible = False
            application.DisplayAlerts = 0
            document = application.Documents.Open(
                str(source), ReadOnly=True, AddToRecentFiles=False
            )
            document.ExportAsFixedFormat(str(output), 17)
        else:
            application = win32com.client.DispatchEx("Excel.Application")
            application.Visible = False
            application.DisplayAlerts = False
            document = application.Workbooks.Open(
                str(source), ReadOnly=True, UpdateLinks=0, AddToMru=False
            )
            document.ExportAsFixedFormat(0, str(output))
    except Exception as exc:
        raise RuntimeError(f"Office conversion failed for {source.name}: {exc}") from exc
    finally:
        try:
            if document is not None:
                document.Close(False)
        finally:
            try:
                if application is not None:
                    application.Quit()
            finally:
                pythoncom.CoUninitialize()


def merge_files(sources: list[Path]) -> Path:
    from PIL import Image
    from pypdf import PdfReader, PdfWriter

    if not sources:
        raise ValueError("Add at least one PDF, JPG/PNG image, or Word/Excel document.")

    supported_extensions = (
        IMAGE_EXTENSIONS | {PDF_EXTENSION} | WORD_EXTENSIONS | EXCEL_EXTENSIONS
    )
    for source in sources:
        if source.suffix.lower() not in supported_extensions:
            raise ValueError(f"Unsupported file type: {source.name}")
        if not source.is_file():
            raise ValueError(f"File does not exist: {source}")

    writer = PdfWriter()
    readers = []
    with tempfile.TemporaryDirectory(prefix="merge_pdf_") as temporary_directory:
        temporary_path = Path(temporary_directory)
        try:
            for index, source in enumerate(sources):
                suffix = source.suffix.lower()
                if suffix == PDF_EXTENSION:
                    pdf_path = source
                else:
                    pdf_path = temporary_path / f"{index}.pdf"
                    if suffix in IMAGE_EXTENSIONS:
                        with Image.open(source) as image:
                            rgba = image.convert("RGBA")
                            background = Image.new("RGB", rgba.size, "white")
                            background.paste(rgba, mask=rgba.getchannel("A"))
                            background.save(pdf_path, format="PDF")
                    else:
                        convert_office_to_pdf(source, pdf_path)

                reader = PdfReader(str(pdf_path))
                readers.append(reader)
                writer.append(reader)

            directory = sources[0].parent
            index = 0
            while True:
                filename = (
                    "Merged_Output.pdf"
                    if index == 0
                    else f"Merged_Output ({index}).pdf"
                )
                output = directory / filename
                try:
                    with output.open("xb") as result:
                        writer.write(result)
                    return output
                except FileExistsError:
                    index += 1
                except Exception:
                    output.unlink(missing_ok=True)
                    raise
        finally:
            for reader in readers:
                reader.close()


def show_merge_window(sources: list[Path]) -> None:
    root = tk.Tk()
    root.title("Merge as PDF")
    root.geometry("520x360")
    root.minsize(420, 280)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    listbox = tk.Listbox(root, selectmode=tk.SINGLE, exportselection=False)
    listbox.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=12, pady=12)
    for source in sources:
        listbox.insert(tk.END, source.name)

    def move_selected(offset: int) -> None:
        selected = listbox.curselection()
        if not selected:
            return
        index = selected[0]
        new_index = index + offset
        if not 0 <= new_index < len(sources):
            return
        sources[index], sources[new_index] = sources[new_index], sources[index]
        name = listbox.get(index)
        listbox.delete(index)
        listbox.insert(new_index, name)
        listbox.selection_set(new_index)
        listbox.activate(new_index)

    def remove_selected() -> None:
        selected = listbox.curselection()
        if not selected:
            return
        index = selected[0]
        listbox.delete(index)
        sources.pop(index)
        if sources:
            listbox.selection_set(min(index, len(sources) - 1))

    def merge_selected() -> None:
        try:
            output = merge_files(sources)
        except Exception as exc:
            messagebox.showerror("Merge failed", str(exc), parent=root)
            return
        messagebox.showinfo("Success", f"PDF created:\n{output}", parent=root)
        root.destroy()

    ttk.Button(root, text="Move Up", command=lambda: move_selected(-1)).grid(
        row=1, column=0, sticky="ew", padx=(12, 4), pady=(0, 12)
    )
    ttk.Button(root, text="Move Down", command=lambda: move_selected(1)).grid(
        row=1, column=1, sticky="ew", padx=4, pady=(0, 12)
    )
    ttk.Button(root, text="Remove", command=remove_selected).grid(
        row=1, column=2, sticky="ew", padx=(4, 12), pady=(0, 12)
    )
    ttk.Button(root, text="Merge as PDF", command=merge_selected).grid(
        row=2, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12)
    )
    root.mainloop()


def main() -> int:
    sources = [Path(argument).resolve() for argument in sys.argv[1:]]
    if not sources:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Merge as PDF",
            "Pass PDF, JPG/PNG, Word, or Excel files to merge.",
        )
        root.destroy()
        return 1

    show_merge_window(sources)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
