import tkinter as tk
from tkinter import filedialog
import zipfile
import shutil
import os
import tempfile
import argparse

ANONYMIZED_SUFFIX = " - Anonymized.docx"

def select_file(title):
    # Open a file dialog to select a .docx file
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Word Documents", "*.docx")])
    return file_path

def prompt_if_missing(value, prompt_text, param=None, input_path=None):
    if value:
        return value
    if param == "input":
        file_path = select_file("Select the input .docx file")
        if file_path:
            return file_path
        # Fallback to prompt if user cancels
        return input(f"{prompt_text}: ").strip()
    elif param == "output":
        # If input_path is provided, default output to same directory with ' - Anonymized.docx'
        if input_path:
            base, _ = os.path.splitext(input_path)
            default_output = base + ANONYMIZED_SUFFIX
            print(f"Output file not specified. Defaulting to: {default_output}")
            return default_output
        return input(f"{prompt_text}: ").strip()
    else:
        return input(f"{prompt_text}: ").strip()

def replace_in_docx(docx_path, output_path, old_name, new_name):
    tmp_dir = tempfile.mkdtemp()

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_dir)

    targets = [
        "word/document.xml",
        "word/comments.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/styles.xml",
        "word/settings.xml",
        "docProps/core.xml",
        "word/header1.xml",
        "word/header2.xml",
        "word/footer1.xml",
        "word/footer2.xml"
    ]

    for target in targets:
        file_path = os.path.join(tmp_dir, target)
        if os.path.exists(file_path):
            with open(file_path, 'rb') as f:
                xml = f.read()

            xml = xml.replace(bytes(old_name, encoding='utf-8'), bytes(new_name, encoding='utf-8'))

            with open(file_path, 'wb') as f:
                f.write(xml)

    base_dir = os.path.dirname(output_path)
    tmp_zip = os.path.join(base_dir, "temp_anonymized.zip")
    shutil.make_archive(tmp_zip.replace(".zip", ""), 'zip', tmp_dir)
    shutil.move(tmp_zip, output_path)

    shutil.rmtree(tmp_dir)
    print(f"✅ Anonymized file saved to:\n{output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Anonymize author name in .docx files. "
                    "You can use the GUI to select files if --input is not provided. "
                    "For bulk processing, provide --folder to anonymize all .docx files in the folder."
    )
    parser.add_argument("--input", help="Path to input .docx file (or use GUI if not provided)")
    parser.add_argument("--output", help="Path to output anonymized .docx file (defaults to ' - Anonymized.docx')")
    parser.add_argument("--old-name", help="Original author name to find")
    parser.add_argument("--new-name", help="Replacement name")
    parser.add_argument("--folder", help="Directory containing .docx files to anonymize in bulk (overrides --input/--output)")

    args = parser.parse_args()

    if args.folder:
        folder = args.folder
        old_name = prompt_if_missing(args.old_name, "Enter author name to find", param="old-name")
        new_name = prompt_if_missing(args.new_name, "Enter replacement name", param="new-name")
        if not os.path.isdir(folder):
            print(f"❌ Folder does not exist: {folder}")
            exit(1)
        files = [
            f for f in os.listdir(folder)
            if f.lower().endswith(".docx") and not f.endswith(ANONYMIZED_SUFFIX)
        ]
        if not files:
            print(f"❌ No .docx files found in folder: {folder}")
            exit(1)
        print(f"Processing {len(files)} .docx files in: {folder}")
        for fname in files:
            input_path = os.path.join(folder, fname)
            base, _ = os.path.splitext(fname)
            output_path = os.path.join(folder, base + ANONYMIZED_SUFFIX)
            print(f"Anonymizing '{fname}'...")
            try:
                replace_in_docx(input_path, output_path, old_name, new_name)
            except Exception as e:
                print(f"❌ Failed to anonymize '{fname}': {e}")
    else:
        input_path = prompt_if_missing(args.input, "Enter path to input .docx", param="input")
        output_path = prompt_if_missing(args.output, "Enter path to output .docx", param="output", input_path=input_path)
        old_name = prompt_if_missing(args.old_name, "Enter author name to find", param="old-name")
        new_name = prompt_if_missing(args.new_name, "Enter replacement name", param="new-name")
        replace_in_docx(input_path, output_path, old_name, new_name)
