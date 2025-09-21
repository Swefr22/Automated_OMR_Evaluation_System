# Automated OMR Evaluation System

## Description
This is a Python-based Streamlit application that automatically evaluates OMR sheets. Users can upload an **answer key (Excel)** and **OMR sheets (JPEG)**, and the system calculates the total score, percentage, and subject-wise marks. It also shows the detected bubbles for visual verification.

---

## Features
- Upload answer key in Excel format (`xls` or `xlsx`)  
- Upload multiple OMR sheet images (`JPEG`)  
- Automatic evaluation and scoring  
- Subject-wise score calculation  
- Visual debugging of detected bubbles  
- Download results as Excel

---

## Requirements
- Python 3.x  
- Libraries:
  - streamlit
  - pandas
  - opencv-python
  - numpy
  - openpyxl

Install dependencies using:

```bash
pip install -r requirements.txt
