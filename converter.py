import argparse
import os
import sys
from pathlib import Path


class ConversionError(Exception):
    pass


def output_path_for(source: Path, target: str) -> Path:
    return source.with_suffix(".md" if target == "md" else ".pdf")


def convert_to_markdown(source: Path, output: Path) -> None:
    try:
        from markitdown import MarkItDown
    except ImportError as exc:
        raise ConversionError(
            "MarkItDown is not installed. Run: py -m pip install -r requirements.txt"
        ) from exc

    try:
        result = MarkItDown().convert(str(source))
        output.write_text(result.text_content, encoding="utf-8")
    except Exception as exc:
        raise ConversionError(f"Markdown conversion failed: {exc}") from exc


def choose_pdf_quality() -> str | None:
    import tkinter as tk
    from tkinter import ttk

    dialog = tk.Tk()
    dialog.title("PDF quality")
    dialog.resizable(False, False)
    quality = tk.StringVar(value="Normal")
    result: dict[str, str | None] = {"quality": None}

    ttk.Label(dialog, text="Select PDF quality:").pack(padx=16, pady=(14, 6), anchor="w")
    for option in ("Lowest", "Low", "Normal", "Best"):
        ttk.Radiobutton(dialog, text=option, value=option, variable=quality).pack(
            padx=16, anchor="w"
        )

    def accept() -> None:
        result["quality"] = quality.get()
        dialog.destroy()

    buttons = ttk.Frame(dialog)
    buttons.pack(fill="x", padx=16, pady=14)
    ttk.Button(buttons, text="Convert", command=accept).pack(side="right")
    dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() - width) // 2
    y = (dialog.winfo_screenheight() - height) // 2
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    dialog.grab_set()
    dialog.wait_window()
    return result["quality"]


def convert_image_to_pdf(source: Path, output: Path, quality: str) -> None:
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise ConversionError(
            "Pillow is not installed. Run: py -m pip install -r requirements.txt"
        ) from exc

    try:
        with Image.open(source) as image:
            image = ImageOps.exif_transpose(image)
            scale = {"Lowest": 0.5, "Low": 0.75}.get(quality, 1)
            if scale < 1:
                size = (
                    max(1, round(image.width * scale)),
                    max(1, round(image.height * scale)),
                )
                image = image.resize(size, Image.Resampling.LANCZOS)
            save_options = {"resolution": 36} if quality == "Lowest" else {}
            if "A" in image.getbands():
                rgba = image.convert("RGBA")
                background = Image.new("RGB", rgba.size, "white")
                background.paste(rgba, mask=rgba.getchannel("A"))
                background.save(output, "PDF", **save_options)
            else:
                image.convert("RGB").save(output, "PDF", **save_options)
    except Exception as exc:
        raise ConversionError(f"Image-to-PDF conversion failed: {exc}") from exc


def convert_office_to_pdf(
    source: Path, output: Path, application: str, quality: str
) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ConversionError(
            "Office PDF conversion requires pywin32 and Microsoft Office. "
            "Run: py -m pip install -r requirements.txt"
        ) from exc

    pythoncom.CoInitialize()
    app = None
    document = None
    optimize_for_screen = quality in {"Lowest", "Low"}
    try:
        app = win32com.client.DispatchEx(application)
        if application == "Word.Application":
            app.Visible = False
            app.DisplayAlerts = 0
            document = app.Documents.Open(
                str(source), ReadOnly=True, AddToRecentFiles=False
            )
            document.ExportAsFixedFormat(
                str(output), 17, False, 1 if optimize_for_screen else 0
            )
        elif application == "Excel.Application":
            app.Visible = False
            app.DisplayAlerts = False
            document = app.Workbooks.Open(
                str(source), ReadOnly=True, UpdateLinks=0, AddToMru=False
            )
            document.ExportAsFixedFormat(
                0, str(output), 1 if optimize_for_screen else 0
            )
        else:
            app.Visible = True
            document = app.Presentations.Open(
                str(source), ReadOnly=True, Untitled=False, WithWindow=True
            )
            document.ExportAsFixedFormat(
                str(output),
                2,
                1 if optimize_for_screen else 2,
                False,
                1,
                0,
                False,
                None,
                0,
                "",
                True,
                True,
                True,
                True,
                False,
            )
    except Exception as exc:
        raise ConversionError(f"Office PDF conversion failed: {exc}") from exc
    finally:
        try:
            if document is not None:
                document.Close(False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


def convert_to_pdf(source: Path, output: Path, quality: str) -> None:
    suffix = source.suffix.lower()
    if suffix in {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}:
        convert_image_to_pdf(source, output, quality)
        return

    if suffix in {".doc", ".docx", ".docm", ".rtf", ".txt"}:
        application = "Word.Application"
    elif suffix in {".xls", ".xlsx", ".xlsm", ".xlsb", ".csv"}:
        application = "Excel.Application"
    elif suffix in {".ppt", ".pptx", ".pptm", ".pps", ".ppsx"}:
        application = "PowerPoint.Application"
    else:
        raise ConversionError(
            f"PDF conversion is not supported for '{suffix or 'files without an extension'}'. "
            "Supported inputs are common image and Microsoft Office formats."
        )

    convert_office_to_pdf(source, output, application, quality)


def show_error(message: str) -> None:
    if sys.stderr is not None:
        print(message, file=sys.stderr)
        return

    if os.name == "nt":
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, message, "File conversion failed", 0x10)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("--target", choices=("md", "pdf"), required=True)
    args = parser.parse_args()

    try:
        source = args.source.resolve(strict=True)
        if not source.is_file():
            raise ConversionError(f"Source is not a file: {source}")

        output = output_path_for(source, args.target)
        quality = "Normal"
        if args.target == "pdf":
            selected_quality = choose_pdf_quality()
            if selected_quality is None:
                return 0
            quality = selected_quality
        if os.path.normcase(str(source)) == os.path.normcase(str(output)):
            return 0
        if output.exists():
            raise ConversionError(f"Output already exists; refusing to overwrite: {output}")

        if args.target == "md":
            convert_to_markdown(source, output)
        else:
            convert_to_pdf(source, output, quality)
    except Exception as exc:
        show_error(str(exc))
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
