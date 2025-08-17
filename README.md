# XFile-Tracker
Designed for Lazy Geniuses

Word Document Comparison Tool

A simple Python Tkinter tool to compare two Word documents (.docx) and show the differences:

- Detects added and removed paragraphs.
- Maps changes to their nearest heading (using Word’s heading styles).
- Shows the true page number (via PDF conversion).
- Displays results in a scrollable Tkinter window.

This tool is handy when you want to quickly see what’s changed between two versions of a Word document without relying on Word’s “Track Changes”.

✨ Features

1. Upload original and modified Word documents.
2. Converts Word to PDF to preserve accurate pagination.
3. Compares content at paragraph level to detect:
    ✅ Added text
    ✅ Removed text
4. Shows changes under the correct Word heading.
5. Lightweight Tkinter GUI with a scrollable results window.



Install dependencies:
pip install python-docx pdfplumber docx2pdf

🚀 Usage

1. Run the script:
python bin.py
2. Click Upload Original Document and select the baseline .docx.
3. Click Upload Modified Document and select the updated .docx.
4. Click Compare Documents.
A new window will pop up with a report of all changes, including page numbers and headings.


📝 Requirements

Python 3.7+
Microsoft Word installed (for docx2pdf on Windows)
Packages:
python-docx
pdfplumber
docx2pdf

⚠️ Limitations

- Page detection relies on PDF text extraction, so very long or oddly formatted paragraphs may occasionally map to “Unknown” page.
- Word’s “Track Changes” feature offers deeper diffs — this tool is a lightweight alternative for quick checks.
- Currently detects paragraph-level changes (not character-level).


🤝 Contributing

Contributions are welcome! Feel free to fork the repo, open issues, and submit pull requests.

📜 License

This project is open source under the MIT License.
