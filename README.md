# AnonymizeReviewer

A Python script to anonymize tracked changes and comment authors in `.docx` Word documents. This is useful when you want to preserve edits/comments but remove or replace the original author's name before sharing the document.

---

## Features

- Replaces author names in:
  - Comments
  - Tracked changes
  - Style definitions
  - Headers/footers
  - Core document metadata
- Supports interactive prompts OR command-line arguments
- Supports batch processing of multiple `.docx` files in a folder
- GUI file selector available for Mac/Linux/Windows

---

## Requirements

- Python 3.7+
- `lxml` library

---

## Setup Instructions

1. **Clone the project or download the files**:
   ```bash
   git clone https://github.com/gregvarghese/AnonymizeReviewer.git
   cd AnonymizeReviewer
   ```

2. **Create a virtual environment (optional but recommended)**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate   # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

---

## How to Use

### Option 1: Interactive Mode (No Arguments)

Simply run:
```bash
python anonymize_docx.py
```

You'll be prompted to:
- Select the input file via GUI
- Provide old name and new replacement name
- Confirm output file path

---

### Option 2: Command-line Arguments

```bash
python anonymize_docx.py \
  --input "/path/to/your.docx" \
  --output "/path/to/output.docx" \
  --old-name "John Doe" \
  --new-name "Reviewer"
```

---

### Option 3: Bulk Anonymize All `.docx` in a Folder

```bash
python anonymize_docx.py --folder "/path/to/folder" --old-name "John Doe" --new-name "Reviewer"
```

This creates anonymized copies of each file in the same folder.

---

## Testing

Open the resulting `.docx` file and:
- Hover over comments or changes
- Confirm the author's name has been replaced

---

## License

This project is licensed under the MIT License.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
