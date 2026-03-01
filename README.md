# 📄 AI-Based Smart PDF Summarizer & Auto Highlighter

An AI-powered document processing system that automatically summarizes documents and highlights important entities such as dates, numbers, names, organizations, locations, and technical terms.

---

## 🚀 Features

- Multi-format support (PDF, DOCX, PPTX)
- OCR for scanned & handwritten PDFs
- Extractive summarization (35% compression)
- Intelligent auto-highlighting (6 entity types)
- Color-coded highlighting
- Section-wise summaries
- Keyword extraction (Top 10 keywords)
- Professional PDF output generation
- 100% local processing (privacy focused)

---

## 🎨 Highlighting Categories

 --> Entity Type      

 Dates         
 Numbers          
 Names             
 Technical Terms   
 Locations        
 Organizations     

---

## 🧠 How It Works

1. Upload document  
2. Text extraction (PyMuPDF / python-docx / python-pptx)  
3. OCR pipeline (EasyOCR + Tesseract fallback)  
4. Entity detection (Regex + NLP)  
5. Auto-highlighting with color coding  
6. Extractive summarization with entity boosting  
7. PDF generation using ReportLab  
8. Download highlighted summary  

---

## 🛠 Tech Stack

- Python  
- Django  
- PyMuPDF  
- spaCy  
- NLTK  
- EasyOCR  
- Tesseract OCR  
- ReportLab  
- python-docx  
- python-pptx  

---

## ⚙ Installation

```bash
git clone https://github.com/dipendrakumar07/AI-PDF-Summarizer-Auto-Highlighter.git
# Navigate into project
cd AI-PDF-Summarizer-Auto-Highlighter

# Create virtual environment
python -m venv .venv

# Activate (Windows)
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run server
python manage.py runserver
```

Open in browser:

```
http://127.0.0.1:8000
```

## 📊 Performance

- Typed PDF (10 pages): 2–3 seconds  
- OCR-based PDF (5 pages): 30–60 seconds  
- Highlighting Accuracy: 92–95%  
- Reduces manual review time by ~70%  

---

## 👨‍💻 Developed By

- Dipendra Kumar Chaudhary  
- Zafran Tariq

## Guided by:
- Prof. Dr. Vikas Somani

---

## 📌 Future Enhancements

- Abstractive summarization  
- Multi-language support  
- REST API  
- Batch document processing  

---

## 📄 License

This project is developed for academic and educational purposes.
