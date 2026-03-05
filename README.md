# 🏫 Student Exam Data Consolidator

A Streamlit web application for schools and colleges to consolidate student exam marks across multiple exams, compute final results, assign ranks, and generate printable PDF result slips.

---

## ✨ Features

- **Multi-Faculty Support** — Arts, Commerce, and Science faculties with correct subject configurations
- **Optional Subject Handling** — Auto-detects whether a student took Sociology/Vocational (Arts) or Biology/Maths (Science)
- **Internal Marks Entry** — Enter INT/Practical marks directly in the app per student per subject
- **Consolidated Marksheet** — Automatically computes:
  - Total out of 200 per subject (across all 5 exam components)
  - Average out of 100 per subject
  - Grand Total (out of 600), Percentage, Pass/Fail, Rank
- **Excel Download** — Fully formatted Excel file with **live formulas** (SUM, ROUND, IF, COUNTIF) that recalculate offline
- **PDF Result Slips** — Exam-wise result slips for every student, 2 per page, ready to print

---

## 🏗️ Project Structure

```
├── app.py               # Main Streamlit application
├── requirements.txt     # Python dependencies
├── README.md            # This file
└── .streamlit/
    └── config.toml      # Streamlit theme configuration
```

---

## 🚀 Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/your-username/student-marksheet-app.git
cd student-marksheet-app
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the app
```bash
streamlit run app.py
```

---

## 📁 Excel File Format

Your uploaded Excel file must have **separate sheets** for each exam. Sheet names must match exactly (case-insensitive):

| Sheet Name       | Exam                  |
|------------------|-----------------------|
| `FIRST UNIT TEST`| First Unit Test (25)  |
| `FIRST TERM`     | First Term Exam (50)  |
| `SECOND UNIT TEST`| Second Unit Test (25)|
| `ANNUAL EXAM`    | Annual Exam (70/80)   |

Each sheet must have these columns:

| Column        | Description                        |
|---------------|------------------------------------|
| `ROLL NO.`    | Student roll number                |
| `STUDENT NAME`| Full name of student               |
| `ENG`         | English marks                      |
| `MAR`         | Marathi marks                      |
| *(subject abbrs)* | Other subject marks per faculty |
| `TOTAL`       | Total marks (optional — recalculated)|
| `%`           | Percentage (optional — recalculated)|
| `RESULT`      | Result (optional — recalculated)   |

### Subject Abbreviations by Faculty

**Arts:** `ENG`, `MAR`, `GEO`, `PSY`, `ECO`, `SOC` or `VOC`  
**Commerce:** `ENG`, `MAR`, `ECO`, `ACC`, `O.C.`, `S.P.`  
**Science:** `ENG`, `MAR`, `GEO`, `PHY`, `CHE`, `BIO` or `MATH`

---

## 📐 Marking Scheme

### Individual Exam Pass Marks (per subject)

| Exam                    | Max Marks | Pass Mark |
|-------------------------|-----------|-----------|
| First Unit Test         | 25        | 9         |
| First Term Exam         | 50        | 18        |
| Second Unit Test        | 25        | 9         |
| Annual Exam (Arts/Commerce) | 80    | 28        |
| Annual Exam (Science Sub 4-6) | 70  | 25        |

### Internal / Practical Marks

| Faculty         | Subjects       | Internal Max |
|-----------------|----------------|--------------|
| Arts / Commerce | All 6          | 20           |
| Science         | Sub 1–3        | 20           |
| Science         | Sub 4–6 (BIO/MATH) | 30       |

### Final Consolidated Result

- **Total per subject** = UT1 + FT + UT2 + Annual + Internal = **200**
- **Average per subject** = Total ÷ 2 = **out of 100**
- **Pass criteria** = All 6 subjects average ≥ **35 marks**
- **Grand Total** = Sum of all 6 averages = **out of 600**
- **Rank** = Dense rank among PASS students by Grand Total (highest = Rank 1)

---

## 📄 PDF Result Slips

After generating the final report, scroll down to the **PDF Result Slips** section:

1. Enter your **School / College Name**
2. Select which exam(s) to generate PDFs for
3. Click **Generate PDF Result Slips**
4. Download one PDF per exam — each PDF has **2 students per page**

Each slip includes:
- School name, exam name, faculty
- Student roll number and name
- Marks table (subject, max marks, obtained marks)
- Total and percentage
- PASS ✓ / FAIL ✗ result
- Class Teacher and Principal signature lines

---

## 🌐 Deploy on Streamlit Cloud

1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repository
4. Set **Main file path** to `app.py`
5. Click **Deploy**

---

## 📦 Dependencies

| Package      | Purpose                        |
|--------------|--------------------------------|
| `streamlit`  | Web application framework      |
| `pandas`     | Data manipulation              |
| `numpy`      | Numerical calculations         |
| `openpyxl`   | Excel file generation          |
| `reportlab`  | PDF generation                 |

---

## 📝 License

This project is for educational use. Feel free to adapt it for your school or institution.
