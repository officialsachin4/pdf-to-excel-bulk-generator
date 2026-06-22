# 📄 PDF to Excel Bulk Generator

A Python desktop application that converts **multiple PDF files containing structured tables** into **Excel spreadsheets** while preserving the original data as accurately as possible.

This tool is designed for users who need to process large batches of PDFs without manually copying data into Excel.

---

## ✨ Features

* 📁 Bulk PDF processing
* 📊 Extracts structured tables from PDF files
* 📄 Converts each PDF into an Excel workbook
* ⚡ Fast batch conversion
* 🖥️ Simple and user-friendly GUI (Tkinter)
* 📂 Select input and output folders
* 📈 Progress bar with conversion status
* 📝 Error logging for failed files
* 💾 Supports `.xlsx` output format

---

## 📷 Application Workflow

1. Launch the application.
2. Select the folder containing PDF files.
3. Choose the output folder.
4. Click **Convert**.
5. The application processes all PDFs automatically.
6. Converted Excel files are saved in the selected output directory.

---

## 🛠 Technologies Used

* Python 3.11+
* Tkinter
* pdfplumber
* Camelot
* Tabula-py
* pandas
* openpyxl

---

## 📦 Installation

Clone the repository:

```bash
git clone https://github.com/yourusername/pdf-to-excel-bulk-generator.git
cd pdf-to-excel-bulk-generator
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the application:

```bash
python app.py
```

---

## 📁 Project Structure

```
PDF-To-Excel-Bulk-Generator/
│
├── app.py
├── converter.py
├── gui.py
├── utils.py
├── requirements.txt
├── README.md
│
├── input/
├── output/
├── logs/
└── assets/
```

---

## 📄 Supported Input

* Text-based PDF files
* PDFs containing structured tables
* Multiple PDF files in a single folder

---

## 📊 Output

Each PDF is converted into an Excel workbook.

Example:

```
Invoice1.pdf
        ↓
Invoice1.xlsx

Report.pdf
        ↓
Report.xlsx
```

---

## ⚠ Limitations

* Scanned PDFs require OCR before conversion.
* Highly complex layouts may require manual verification.
* Password-protected PDFs are not supported.

---

## 🚀 Future Improvements

* OCR support using Tesseract
* Drag-and-drop PDF upload
* Merge all PDFs into a single Excel workbook
* Multi-threaded processing
* Export to CSV
* AI-powered table detection
* Executable (.exe) version for Windows

---

## 🤝 Contributing

Contributions are welcome!

1. Fork the repository.
2. Create a feature branch.
3. Commit your changes.
4. Open a Pull Request.

---

## 📜 License

This project is licensed under the MIT License.

---

## 👨‍💻 Author

**Sachin Sharma**

Python Automation Engineer | Playwright | Desktop Automation | PDF Processing | Java | Spring Boot | .NET

If you found this project useful, consider giving it a ⭐ on GitHub!
