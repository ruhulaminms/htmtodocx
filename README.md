# HTML to DOCX Converter (Python)

This Python script automatically converts all `.html` and `.htm` files in a folder into `.docx` (Microsoft Word) format using the Windows COM interface.

## 📌 Features

- Converts multiple HTML files in bulk
- Automatically detects `.html` and `.htm` files
- Saves output as `.docx`
- Uses Microsoft Word for accurate rendering
- Runs silently (Word window stays hidden)

## ⚙️ Requirements

Before running the script, make sure you have:

- Python installed (3.x recommended)
- **Microsoft Word 2016 (or later) installed**
- Required Python package:

```bash
pip install pywin32
