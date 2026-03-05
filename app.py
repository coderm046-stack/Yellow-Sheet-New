import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

st.set_page_config(page_title="Consolidated Marksheet Pro", layout="wide")
st.title("🏫 Student Exam Data Consolidator")

# ── Faculty / Subject Definitions ─────────────────────────────────────────────
# Each faculty defines:
#   'subjects'  : ordered list of (abbr, full_name, annual_max, internal_max)
#   'optional'  : list of abbrs that are optional (student picks exactly one)
#                 the app auto-detects which optional the student took
#   'core'      : always-present subjects (first 5 or first 6 excl optional group)

FACULTY_CONFIG = {
    "Arts": {
        "core":     ["ENG", "MAR", "GEO", "PSY", "ECO"],
        "optional": ["SOC", "VOC"],          # student takes exactly ONE
        "subjects": {
            "ENG": ("English",     80, 20),
            "MAR": ("Marathi",     80, 20),
            "GEO": ("Geography",   80, 20),
            "PSY": ("Psychology",  80, 20),
            "ECO": ("Economics",   80, 20),
            "SOC": ("Sociology",   80, 20),
            "VOC": ("Vocational",  80, 20),
        },
    },
    "Commerce": {
        "core":     ["ENG", "MAR", "ECO", "ACC", "O.C.", "S.P."],
        "optional": [],
        "subjects": {
            "ENG":  ("English",    80, 20),
            "MAR":  ("Marathi",    80, 20),
            "ECO":  ("Economics",  80, 20),
            "ACC":  ("Accounts",   80, 20),
            "O.C.": ("O.C.",       80, 20),
            "S.P.": ("S.P.",       80, 20),
        },
    },
    "Science": {
        "core":     ["ENG", "MAR", "GEO", "PHY", "CHE"],
        "optional": ["BIO", "MATH"],         # student takes exactly ONE
        "subjects": {
            "ENG":  ("English",    80, 20),
            "MAR":  ("Marathi",    80, 20),
            "GEO":  ("Geography",  80, 20),
            "PHY":  ("Physics",    80, 20),
            "CHE":  ("Chemistry",  80, 20),
            "BIO":  ("Biology",    70, 30),
            "MATH": ("Maths",      70, 30),
        },
    },
}

# Passing marks reference (individual exams — informational only)
# Final overall pass = average/100 >= 35 in ALL 6 subjects
EXAM_PASS = {
    "FIRST UNIT TEST (25)":  9,
    "FIRST TERM EXAM (50)": 18,
    "SECOND UNIT TEST (25)": 9,
    "ANNUAL EXAM (70/80)":  28,   # 28 for /80, 25 for /70 — shown only
}

def custom_round(x):
    try:
        return int(np.floor(float(x) + 0.5))
    except:
        return 0

def clean_marks(val):
    if isinstance(val, str):
        v = val.strip().upper()
        if v in ("AB", ""):
            return 0.0
    try:
        return float(val)
    except:
        return 0.0

# ── Detect which 6 subjects a student actually has (handles optional) ─────────
def detect_student_subjects(faculty, df_row, cfg):
    """
    Returns ordered list of 6 abbrs for this student.
    For optional groups, picks whichever column has a non-zero / non-blank value.
    If both blank, defaults to first optional.
    """
    core = cfg["core"]
    opt  = cfg["optional"]
    if not opt:
        return core  # exactly 6 core subjects

    # Pick optional: whichever has a value in this row
    chosen_opt = opt[0]  # default
    for o in opt:
        val = str(df_row.get(o, "")).strip()
        if val and val.upper() not in ("", "NAN", "0"):
            chosen_opt = o
            break
    return core + [chosen_opt]

# ── PDF Generation ────────────────────────────────────────────────────────────
def build_exam_pdf(school_name, faculty_name, exam_label, student_results,
                   cfg, pos_cols, selected_exam_data):
    """
    Landscape A4, 2 result slips side by side per page.
    Each slip is ~135mm wide — large, readable fonts suitable for printing.
    """
    from reportlab.lib.pagesizes import landscape

    exam_meta = {
        "FIRST UNIT TEST (25)":  {"max_per_sub": 25,  "pass_mark": 9,  "total_max": 150},
        "FIRST TERM EXAM (50)":  {"max_per_sub": 50,  "pass_mark": 18, "total_max": 300},
        "SECOND UNIT TEST (25)": {"max_per_sub": 25,  "pass_mark": 9,  "total_max": 150},
        "ANNUAL EXAM (70/80)":   {"max_per_sub": None,"pass_mark": 28, "total_max": None},
    }
    meta = exam_meta.get(exam_label, {"max_per_sub": None, "pass_mark": None, "total_max": None})

    PAGE      = landscape(A4)           # 297mm × 210mm
    L_MARGIN  = 8*mm
    R_MARGIN  = 8*mm
    T_MARGIN  = 8*mm
    B_MARGIN  = 8*mm
    USABLE_W  = PAGE[0] - L_MARGIN - R_MARGIN   # ~281mm
    SLIP_W    = USABLE_W / 2                     # ~140mm each
    DIVIDER   = 4*mm                             # gap between slips

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=PAGE,
        leftMargin=L_MARGIN, rightMargin=R_MARGIN,
        topMargin=T_MARGIN,  bottomMargin=B_MARGIN,
    )

    # ── Styles (larger since landscape gives us room) ─────────────────────────
    school_style = ParagraphStyle("sch", fontSize=16, fontName="Helvetica-Bold",
                                  alignment=TA_CENTER, spaceAfter=2,
                                  textColor=colors.Color(0.12, 0.31, 0.49))
    exam_style   = ParagraphStyle("exm", fontSize=12, fontName="Helvetica-Bold",
                                  alignment=TA_CENTER, spaceAfter=2,
                                  textColor=colors.Color(0.8, 0.2, 0.0))
    fac_style    = ParagraphStyle("fac", fontSize=10, fontName="Helvetica",
                                  alignment=TA_CENTER, spaceAfter=3)
    lbl_style    = ParagraphStyle("lbl", fontSize=10, fontName="Helvetica-Bold")
    val_style    = ParagraphStyle("val", fontSize=10, fontName="Helvetica")
    res_pass     = ParagraphStyle("rp",  fontSize=14, fontName="Helvetica-Bold",
                                  textColor=colors.green, alignment=TA_CENTER)
    res_fail     = ParagraphStyle("rf",  fontSize=14, fontName="Helvetica-Bold",
                                  textColor=colors.red,   alignment=TA_CENTER)

    # Subject col, Max col, Marks col widths inside each slip
    S_COL = SLIP_W * 0.52
    M_COL = SLIP_W * 0.24
    O_COL = SLIP_W * 0.24

    def make_slip(sr):
        roll   = sr["roll"]
        name   = sr["name"]
        subj_6 = sr["subj_6"]
        exam_d = selected_exam_data.get(roll, {})
        is_ann = (exam_label == "ANNUAL EXAM (70/80)")
        elems  = []

        # Header
        elems.append(Paragraph(school_name, school_style))
        elems.append(Paragraph(exam_label,  exam_style))
        elems.append(Paragraph(f"{faculty_name} Faculty", fac_style))
        elems.append(HRFlowable(width="100%", thickness=1.5,
                                color=colors.Color(0.12, 0.31, 0.49)))
        elems.append(Spacer(1, 3*mm))

        # Roll / Name row
        info = Table(
            [[Paragraph("Roll No. :", lbl_style), Paragraph(str(roll), val_style),
              Paragraph("Name :", lbl_style),      Paragraph(str(name), val_style)]],
            colWidths=[SLIP_W*0.18, SLIP_W*0.18, SLIP_W*0.14, SLIP_W*0.50],
        )
        info.setStyle(TableStyle([
            ("BOTTOMPADDING", (0,0),(-1,-1), 3),
            ("TOPPADDING",    (0,0),(-1,-1), 3),
            ("FONTSIZE",      (0,0),(-1,-1), 10),
        ]))
        elems.append(info)
        elems.append(Spacer(1, 3*mm))

        # Marks table
        tbl_data      = [["Subject", "Max Marks", "Obtained"]]
        total_obt     = 0
        total_max_val = 0

        for abbr in subj_6:
            subj_name, ann_max, _ = cfg["subjects"][abbr]
            max_m = ann_max if is_ann else meta["max_per_sub"]
            raw   = exam_d.get(abbr, "")
            try:
                obt = int(float(raw)) if str(raw).strip().upper() != "AB" else "AB"
            except:
                obt = str(raw)
            if isinstance(obt, int):
                total_obt     += obt
                total_max_val += max_m
            tbl_data.append([subj_name, str(max_m), str(obt)])

        tbl_data.append(["TOTAL", str(total_max_val),
                          str(total_obt) if isinstance(total_obt, int) else "-"])
        pct = round(total_obt / total_max_val * 100, 2) \
              if isinstance(total_obt, int) and total_max_val else "-"
        tbl_data.append(["Percentage", "", f"{pct} %" if pct != "-" else "-"])

        n_rows = len(tbl_data)
        marks_tbl = Table(tbl_data, colWidths=[S_COL, M_COL, O_COL])
        marks_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0),    colors.Color(0.12,0.31,0.49)),
            ("TEXTCOLOR",     (0,0), (-1,0),    colors.white),
            ("FONTNAME",      (0,0), (-1,0),    "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1),   10),
            ("ALIGN",         (1,0), (-1,-1),   "CENTER"),
            ("ALIGN",         (0,0), (0,-1),    "LEFT"),
            ("ROWBACKGROUNDS",(0,1), (-1,n_rows-3),
                              [colors.white, colors.Color(0.94,0.96,0.99)]),
            ("BACKGROUND",    (0,-2),(-1,-2),   colors.Color(0.82,0.82,0.82)),
            ("FONTNAME",      (0,-2),(-1,-2),   "Helvetica-Bold"),
            ("BACKGROUND",    (0,-1),(-1,-1),   colors.Color(0.94,0.96,0.99)),
            ("SPAN",          (0,-1),(1,-1)),
            ("FONTNAME",      (0,-1),(-1,-1),   "Helvetica-Bold"),
            ("GRID",          (0,0), (-1,-1),   0.5, colors.grey),
            ("BOTTOMPADDING", (0,0), (-1,-1),   4),
            ("TOPPADDING",    (0,0), (-1,-1),   4),
            ("LEFTPADDING",   (0,0), (-1,-1),   4),
        ]))
        elems.append(marks_tbl)
        elems.append(Spacer(1, 4*mm))

        # Pass / Fail
        indiv_pass = True
        for abbr in subj_6:
            raw = exam_d.get(abbr, "")
            try:
                m   = float(raw)
                req = (28 if cfg["subjects"][abbr][1] == 80 else 25) \
                      if is_ann else meta["pass_mark"]
                if m < req:
                    indiv_pass = False; break
            except:
                indiv_pass = False

        elems.append(Paragraph(
            "✓   PASS" if indiv_pass else "✗   FAIL",
            res_pass if indiv_pass else res_fail
        ))
        elems.append(Spacer(1, 6*mm))

        # Signature line
        sig = Table(
            [["Class Teacher", "Principal"]],
            colWidths=[SLIP_W * 0.5, SLIP_W * 0.5],
        )
        sig.setStyle(TableStyle([
            ("FONTSIZE",    (0,0),(-1,-1), 9),
            ("ALIGN",       (0,0),(-1,-1), "CENTER"),
            ("LINEABOVE",   (0,0),(0,0),   0.7, colors.black),
            ("LINEABOVE",   (1,0),(1,0),   0.7, colors.black),
            ("TOPPADDING",  (0,0),(-1,-1), 14),
        ]))
        elems.append(sig)
        return elems

    # ── Assemble: 2 slips per page side by side ───────────────────────────────
    story = []
    for i in range(0, len(student_results), 2):
        left  = make_slip(student_results[i])
        right = make_slip(student_results[i+1]) \
                if i+1 < len(student_results) else [Spacer(1, 1)]

        page_tbl = Table(
            [[left, right]],
            colWidths=[SLIP_W - DIVIDER/2, SLIP_W - DIVIDER/2],
        )
        page_tbl.setStyle(TableStyle([
            ("VALIGN",       (0,0),(-1,-1), "TOP"),
            ("LINEAFTER",    (0,0),(0,-1),  1.0, colors.Color(0.7,0.7,0.7)),
            ("LEFTPADDING",  (0,0),(-1,-1), 4),
            ("RIGHTPADDING", (0,0),(-1,-1), 4),
            ("TOPPADDING",   (0,0),(-1,-1), 0),
            ("BOTTOMPADDING",(0,0),(-1,-1), 0),
        ]))
        story.append(page_tbl)
        if i + 2 < len(student_results):
            story.append(PageBreak())

    doc.build(story)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
uploaded_file = st.file_uploader("Upload Excel Marksheet", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)

        # ── Faculty selection ──────────────────────────────────────────────────
        st.markdown("---")
        faculty = st.selectbox("🎓 Select Faculty", list(FACULTY_CONFIG.keys()))
        cfg     = FACULTY_CONFIG[faculty]

        # Show subject table
        st.markdown("**Subject Configuration:**")
        sub_info = []
        for abbr in cfg["core"] + cfg["optional"]:
            name, am, im = cfg["subjects"][abbr]
            tag = ""
            if abbr in cfg["optional"]:
                tag = " *(optional)*"
            sub_info.append({
                "Abbr": abbr, "Subject": name + tag,
                "Annual Max": am, "Internal Max": im,
                "Total": f"25+50+25+{am}+{im}=200"
            })
        st.dataframe(pd.DataFrame(sub_info), hide_index=True, use_container_width=True)

        # ── Parse exam sheets ──────────────────────────────────────────────────
        exam_configs = [
            {"label": "FIRST UNIT TEST (25)",  "sheets": ["FIRST UNIT TEST"]},
            {"label": "FIRST TERM EXAM (50)",  "sheets": ["FIRST TERM"]},
            {"label": "SECOND UNIT TEST (25)", "sheets": ["SECOND UNIT TEST"]},
            {"label": "ANNUAL EXAM (70/80)",   "sheets": ["ANNUAL EXAM"]},
        ]

        all_students = {}   # roll -> {Name, Exams, subjects (list of 6 abbrs)}

        for config in exam_configs:
            sheet_name = next(
                (s for s in xl.sheet_names if s.strip().upper() in config["sheets"]), None
            )
            if not sheet_name:
                continue

            df = xl.parse(sheet_name)
            df.columns = df.columns.astype(str).str.strip().str.upper()

            # Normalise O.C. and S.P. column names (dots may be stripped)
            col_map = {}
            for c in df.columns:
                cn = c.replace(" ", "")
                if cn in ("OC", "O.C"):  col_map[c] = "O.C."
                if cn in ("SP", "S.P"):  col_map[c] = "S.P."
            if col_map:
                df = df.rename(columns=col_map)

            t_col = next((c for c in df.columns if "TOTAL" in c), None)
            p_col = next((c for c in df.columns if "%" in c or "PERCENT" in c), None)
            r_col = next((c for c in df.columns if "RESULT" in c), None)

            for _, row in df.iterrows():
                roll = str(row.get("ROLL NO.", "")).strip()
                if not roll or roll.lower() == "nan":
                    continue

                if roll not in all_students:
                    subj_6 = detect_student_subjects(faculty, row, cfg)
                    all_students[roll] = {
                        "Name":     str(row.get("STUDENT NAME", "Unknown")),
                        "Exams":    {},
                        "subjects": subj_6,
                    }

                subj_6 = all_students[roll]["subjects"]
                marks  = {}
                for abbr in subj_6:
                    raw = str(row.get(abbr, "0")).strip()
                    marks[abbr] = raw if raw.upper() == "AB" else row.get(abbr, 0)

                try:
                    marks["Grand Total"] = str(row.get(t_col, "")) if t_col else ""
                except:
                    marks["Grand Total"] = ""
                try:
                    raw_p = row.get(p_col, "")
                    marks["%"] = str(round(float(raw_p), 2)) if str(raw_p).strip() else ""
                except:
                    marks["%"] = ""
                marks["Result"] = str(row.get(r_col, ""))
                all_students[roll]["Exams"][config["label"]] = marks

        if not all_students:
            st.error("No student data found. Check sheet names match: FIRST UNIT TEST, FIRST TERM, SECOND UNIT TEST, ANNUAL EXAM")
            st.stop()

        student_rolls = sorted(
            all_students.keys(),
            key=lambda x: float(x) if x.replace(".", "", 1).isdigit() else 0
        )

        categories = [
            "FIRST UNIT TEST (25)",
            "FIRST TERM EXAM (50)",
            "SECOND UNIT TEST (25)",
            "ANNUAL EXAM (70/80)",
            "INT/PRACTICAL (20/30)",
            "Total Marks Out of 200",
            "Average Marks 200/2=100",
        ]
        result_cols = ["Grand Total", "%", "Result", "Remark", "Rank"]

        # ── Internal Marks Input ───────────────────────────────────────────────
        st.markdown("---")
        st.subheader("📝 Enter Internal / Practical Marks")
        st.info("Enter marks for each student. Subject columns match each student's chosen optional subject.")

        if "internal_marks" not in st.session_state:
            st.session_state.internal_marks = {
                roll: {abbr: "0" for abbr in all_students[roll]["subjects"]}
                for roll in student_rolls
            }

        # Header row
        hdr = st.columns([0.6, 1.8] + [0.9]*6)
        hdr[0].markdown("**Roll**")
        hdr[1].markdown("**Name**")

        for roll in student_rolls:
            subj_6 = all_students[roll]["subjects"]
            name   = all_students[roll]["Name"]
            cols   = st.columns([0.6, 1.8] + [0.9]*6)
            cols[0].write(roll)
            cols[1].write(name)
            for i, abbr in enumerate(subj_6):
                _, am, im = cfg["subjects"][abbr]
                val = cols[i+2].text_input(
                    label=f"{roll}-{abbr}",
                    value=st.session_state.internal_marks[roll].get(abbr, "0"),
                    key=f"int_{roll}_{abbr}",
                    label_visibility="collapsed",
                    placeholder=f"{abbr} /{im}",
                )
                st.session_state.internal_marks[roll][abbr] = val

        # ── Build base_df (one universal subject slot per position) ────────────
        # Since different students may have different optional subjects,
        # we use positional column names Sub1-Sub6 for the dataframe
        # but store actual abbr names in subject headers for display

        # Determine display subject headers — use most common subject set
        # (since all students in one faculty have same core, only optional differs)
        display_subj = cfg["core"] + [cfg["optional"][0]] if cfg["optional"] else cfg["core"]

        pos_cols  = [f"Sub{i+1}" for i in range(6)]   # positional df columns
        base_rows = []

        for roll in student_rolls:
            s      = all_students[roll]
            subj_6 = s["subjects"]

            for cat in categories:
                row_data = {
                    "Roll No.": roll if cat == "FIRST UNIT TEST (25)" else "",
                    "Column1":  s["Name"] if cat == "FIRST UNIT TEST (25)" else "",
                    "Column2":  cat,
                    "_subjects": "|".join(subj_6),   # hidden metadata
                }
                for pos, pc in enumerate(pos_cols):
                    row_data[pc] = ""
                for rc in result_cols:
                    row_data[rc] = ""

                if cat in s["Exams"]:
                    exam_marks = s["Exams"][cat]
                    for pos, abbr in enumerate(subj_6):
                        row_data[pos_cols[pos]] = str(exam_marks.get(abbr, ""))
                    row_data["Grand Total"] = exam_marks.get("Grand Total", "")
                    row_data["%"]           = exam_marks.get("%", "")
                    row_data["Result"]      = exam_marks.get("Result", "")

                elif cat == "INT/PRACTICAL (20/30)":
                    for pos, abbr in enumerate(subj_6):
                        row_data[pos_cols[pos]] = st.session_state.internal_marks[roll].get(abbr, "0")

                base_rows.append(row_data)

        base_df = pd.DataFrame(base_rows)
        for col in base_df.columns:
            if col != "_subjects":
                base_df[col] = base_df[col].astype(str).replace("nan", "")

        # Inject latest internal marks
        for i, roll in enumerate(student_rolls):
            subj_6 = all_students[roll]["subjects"]
            for pos, abbr in enumerate(subj_6):
                base_df.at[i*7 + 4, pos_cols[pos]] = \
                    st.session_state.internal_marks[roll].get(abbr, "0")

        # Display with actual subject names as column headers
        st.markdown("---")
        st.subheader("📊 Marks Preview & Edit")

        display_df = base_df.drop(columns=["_subjects"]).copy()
        # Rename Sub1-Sub6 to actual subject names for display
        rename_map = {pos_cols[i]: display_subj[i] if i < len(display_subj) else pos_cols[i]
                      for i in range(6)}
        display_df = display_df.rename(columns=rename_map)
        edited_display = st.data_editor(display_df, hide_index=True, use_container_width=True)

        # Map edited values back to positional columns
        rev_rename = {v: k for k, v in rename_map.items()}
        edited_df  = edited_display.rename(columns=rev_rename)

        # ── Generate Report ────────────────────────────────────────────────────
        if st.button("🚀 Generate Final Report & Rank"):

            student_results = []

            for s_idx, roll in enumerate(student_rolls):
                subj_6 = all_students[roll]["subjects"]
                block  = edited_df.iloc[s_idx*7 : s_idx*7+7].copy().reset_index(drop=True)

                raw = {}
                for row_i in range(5):
                    raw[row_i] = {pc: clean_marks(block.at[row_i, pc]) for pc in pos_cols}

                t200 = {pc: sum(raw[r][pc] for r in range(5)) for pc in pos_cols}
                a100 = {pc: custom_round(t200[pc] / 2) for pc in pos_cols}
                gt   = sum(a100.values())
                pc_  = round((gt / 600) * 100, 2)
                isp  = all(a100[pc] >= 35 for pc in pos_cols)

                student_results.append({
                    "roll":   roll,
                    "name":   all_students[roll]["Name"],
                    "subj_6": subj_6,
                    "t200":   t200,
                    "a100":   a100,
                    "gt":     gt,
                    "pc":     pc_,
                    "pass":   isp,
                    "rank":   "",
                })

            # Dense rank — PASS students only
            pass_gts = sorted(set(sr["gt"] for sr in student_results if sr["pass"]), reverse=True)
            rank_map = {gt_val: r+1 for r, gt_val in enumerate(pass_gts)}
            for sr in student_results:
                sr["rank"] = rank_map[sr["gt"]] if sr["pass"] else ""

            # Rebuild final_df
            processed = []
            for s_idx, sr in enumerate(student_results):
                block = edited_df.iloc[s_idx*7 : s_idx*7+7].copy().reset_index(drop=True)
                for pc in pos_cols:
                    block.at[5, pc] = str(int(sr["t200"][pc]))
                    block.at[6, pc] = str(int(sr["a100"][pc]))
                block.at[6, "Grand Total"] = str(sr["gt"])
                block.at[6, "%"]           = str(sr["pc"])
                block.at[6, "Result"]      = "PASS" if sr["pass"] else "FAIL"
                block.at[6, "Rank"]        = str(sr["rank"])
                processed.append(block)

            final_df = pd.concat(processed).reset_index(drop=True)

            # ── Save everything to session_state so PDF section persists ──────
            st.session_state.report_ready       = True
            st.session_state.student_results    = student_results
            st.session_state.final_df           = final_df
            st.session_state.all_students_snap  = all_students   # snapshot for PDF
            st.session_state.faculty_snap       = faculty
            st.session_state.cfg_snap           = cfg
            st.session_state.pos_cols_snap      = pos_cols
            st.session_state.exam_configs_snap  = exam_configs
            st.session_state.xl_sheet_names     = xl.sheet_names
            st.session_state.display_subj_snap  = display_subj

        # ── Show report if already generated ──────────────────────────────────
        if st.session_state.get("report_ready"):
            student_results = st.session_state.student_results
            final_df        = st.session_state.final_df
            cfg_s           = st.session_state.cfg_snap
            display_subj_s  = st.session_state.display_subj_snap

            passed = sum(1 for sr in student_results if sr["pass"])
            st.success(f"✅ {passed} PASS  |  {len(student_results)-passed} FAIL")

            summary_rows = []
            for sr in student_results:
                row_s = {"Roll No.": sr["roll"], "Name": sr["name"]}
                for i, abbr in enumerate(sr["subj_6"]):
                    row_s[f"{abbr} /100"] = sr["a100"][pos_cols[i]]
                row_s["Grand Total"] = sr["gt"]
                row_s["%"]           = sr["pc"]
                row_s["Result"]      = "PASS" if sr["pass"] else "FAIL"
                row_s["Rank"]        = sr["rank"]
                summary_rows.append(row_s)

            st.subheader("📋 Result Summary")
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

            # ── Build Excel ────────────────────────────────────────────────────
            wb   = Workbook()
            ws   = wb.active
            ws.title = "Consolidated"

            ws_h = wb.create_sheet("_RankHelper")
            ws_h.sheet_state = "hidden"
            ws_h.cell(row=1, column=1, value="GT")
            ws_h.cell(row=1, column=2, value="IsPass")

            hdr_font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
            hdr_fill = PatternFill("solid", start_color="1F4E79")
            cat_fill = {
                "FIRST UNIT TEST (25)":    PatternFill("solid", start_color="DDEBF7"),
                "FIRST TERM EXAM (50)":    PatternFill("solid", start_color="E2EFDA"),
                "SECOND UNIT TEST (25)":   PatternFill("solid", start_color="FFF2CC"),
                "ANNUAL EXAM (70/80)":     PatternFill("solid", start_color="FCE4D6"),
                "INT/PRACTICAL (20/30)":   PatternFill("solid", start_color="EAD1DC"),
                "Total Marks Out of 200":  PatternFill("solid", start_color="D9D9D9"),
                "Average Marks 200/2=100": PatternFill("solid", start_color="BDD7EE"),
            }
            thin = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"),  bottom=Side(style="thin"),
            )
            ctr = Alignment(horizontal="center", vertical="center")

            col_headers = (
                ["Roll No.", "Student Name", "Exam Type"]
                + [cfg_s["subjects"][a][0] if a in cfg_s["subjects"] else a
                   for a in display_subj_s]
                + result_cols
            )
            for ci, h in enumerate(col_headers, 1):
                c = ws.cell(row=1, column=ci, value=h)
                c.font=hdr_font; c.fill=hdr_fill; c.alignment=ctr; c.border=thin

            ws.column_dimensions["A"].width = 10
            ws.column_dimensions["B"].width = 22
            ws.column_dimensions["C"].width = 28
            for i in range(6):
                ws.column_dimensions[get_column_letter(4+i)].width = 11
            for i in range(len(result_cols)):
                ws.column_dimensions[get_column_letter(10+i)].width = 13

            SUB_S = 4
            GT_C  = SUB_S + 6
            PCT_C = GT_C  + 1
            RES_C = PCT_C + 1
            REM_C = RES_C + 1
            RNK_C = REM_C + 1

            sub_lets = [get_column_letter(SUB_S + i) for i in range(6)]
            gt_let   = get_column_letter(GT_C)
            res_let  = get_column_letter(RES_C)
            n        = len(student_results)
            h_gt_rng = f"_RankHelper!$A$2:$A${n+1}"
            avg_excel_rows = []

            for s_idx, roll in enumerate([sr["roll"] for sr in student_results]):
                sr   = student_results[s_idx]
                brow = 2 + s_idx * 7

                for cat_idx, cat in enumerate(categories):
                    erow = brow + cat_idx
                    fl   = cat_fill.get(cat, PatternFill("solid", start_color="FFFFFF"))

                    ws.cell(row=erow, column=1, value=roll if cat_idx == 0 else "")
                    ws.cell(row=erow, column=2, value=sr["name"] if cat_idx == 0 else "")
                    ws.cell(row=erow, column=3, value=cat)

                    if cat == "Total Marks Out of 200":
                        r1, r5 = brow, brow + 4
                        for i, sl in enumerate(sub_lets):
                            c = ws.cell(row=erow, column=SUB_S+i,
                                        value=f"=SUM({sl}{r1}:{sl}{r5})")
                            c.fill=fl; c.border=thin; c.alignment=ctr
                            c.font=Font(name="Arial", bold=True)
                        for ri in range(len(result_cols)):
                            c = ws.cell(row=erow, column=GT_C+ri, value="")
                            c.fill=fl; c.border=thin

                    elif cat == "Average Marks 200/2=100":
                        trow = erow - 1
                        for i, sl in enumerate(sub_lets):
                            c = ws.cell(row=erow, column=SUB_S+i,
                                        value=f"=ROUND({sl}{trow}/2,0)")
                            c.fill=fl; c.border=thin; c.alignment=ctr
                            c.font=Font(name="Arial", bold=True)
                        c = ws.cell(row=erow, column=GT_C,
                                    value=f"=SUM({sub_lets[0]}{erow}:{sub_lets[-1]}{erow})")
                        c.fill=fl; c.border=thin; c.alignment=ctr
                        c.font=Font(name="Arial", bold=True, color="1F4E79")
                        c = ws.cell(row=erow, column=PCT_C,
                                    value=f"=ROUND({gt_let}{erow}/600*100,2)")
                        c.fill=fl; c.border=thin; c.alignment=ctr
                        pass_chk = ",".join([f"{sl}{erow}>=35" for sl in sub_lets])
                        c = ws.cell(row=erow, column=RES_C,
                                    value=f'=IF(AND({pass_chk}),"PASS","FAIL")')
                        c.fill=fl; c.border=thin; c.alignment=ctr
                        c.font=Font(name="Arial", bold=True)
                        ws.cell(row=erow, column=REM_C, value="").border = thin
                        c = ws.cell(row=erow, column=RNK_C, value="")
                        c.fill=fl; c.border=thin; c.alignment=ctr
                        c.font=Font(name="Arial", bold=True, color="C00000")
                        h_row = s_idx + 2
                        ws_h.cell(row=h_row, column=1,
                                  value=f"=Consolidated!{gt_let}{erow}")
                        ws_h.cell(row=h_row, column=2,
                                  value=f'=IF(Consolidated!{res_let}{erow}="PASS",1,0)')
                        avg_excel_rows.append((erow, h_row))

                    else:
                        frow = final_df.iloc[s_idx*7 + cat_idx]
                        for i, pc in enumerate(pos_cols):
                            v = frow.get(pc, "")
                            try: v = float(v)
                            except: pass
                            c = ws.cell(row=erow, column=SUB_S+i, value=v)
                            c.fill=fl; c.border=thin; c.alignment=ctr
                        for ri, rc in enumerate(result_cols):
                            v = "" if rc == "Rank" else frow.get(rc, "")
                            c = ws.cell(row=erow, column=GT_C+ri, value=v)
                            c.fill=fl; c.border=thin; c.alignment=ctr

                    for ci in [1, 2, 3]:
                        c = ws.cell(row=erow, column=ci)
                        c.fill=fl; c.border=thin
                        c.font=Font(name="Arial", bold=(ci == 2 and cat_idx == 0))

            for (erow, h_row) in avg_excel_rows:
                rank_formula = (
                    f"=IF(_RankHelper!$B${h_row}=1,"
                    f"COUNTIF({h_gt_rng},\">\"&_RankHelper!$A${h_row})+1,\"\")"
                )
                fl = cat_fill["Average Marks 200/2=100"]
                c  = ws.cell(row=erow, column=RNK_C, value=rank_formula)
                c.fill=fl; c.border=thin; c.alignment=ctr
                c.font=Font(name="Arial", bold=True, color="C00000")

            ws.freeze_panes = "A2"
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.download_button(
                "📥 Download Excel (with Live Formulas)",
                output.getvalue(),
                "Final_Consolidated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # ── PDF Result Slips — ALWAYS visible once report is generated ─────────
        if st.session_state.get("report_ready"):
            student_results = st.session_state.student_results
            all_students_s  = st.session_state.all_students_snap
            faculty_s       = st.session_state.faculty_snap
            cfg_s           = st.session_state.cfg_snap
            pos_cols_s      = st.session_state.pos_cols_snap
            exam_configs_s  = st.session_state.exam_configs_snap
            sheet_names_s   = st.session_state.xl_sheet_names

            st.markdown("---")
            st.subheader("📄 Generate Exam-wise Result Slip PDFs")

            school_name = st.text_input(
                "🏫 School / College Name",
                value=st.session_state.get("school_name_input", "Your School Name"),
                key="school_name_input",
                help="This appears at the top of every result slip",
            )

            exam_options = [
                ec["label"] for ec in exam_configs_s
                if any(s.strip().upper() in ec["sheets"] for s in sheet_names_s)
            ]
            sel_exams = st.multiselect(
                "Select Exam(s) to generate PDFs for",
                options=exam_options,
                default=exam_options,
                key="sel_exams_pdf",
            )

            if st.button("📄 Generate PDF Result Slips", key="gen_pdf_btn"):
                if not sel_exams:
                    st.warning("Please select at least one exam.")
                else:
                    st.session_state.pdf_results = {}
                    for exam_label in sel_exams:
                        exam_data_for_pdf = {}
                        for sr in student_results:
                            roll   = sr["roll"]
                            subj_6 = sr["subj_6"]
                            ed     = all_students_s[roll]["Exams"].get(exam_label, {})
                            exam_data_for_pdf[roll] = {abbr: ed.get(abbr, "") for abbr in subj_6}
                            exam_data_for_pdf[roll]["Grand Total"] = ed.get("Grand Total", "")
                            exam_data_for_pdf[roll]["%"]           = ed.get("%", "")
                            exam_data_for_pdf[roll]["Result"]      = ed.get("Result", "")

                        pdf_buf = build_exam_pdf(
                            school_name        = school_name,
                            faculty_name       = faculty_s,
                            exam_label         = exam_label,
                            student_results    = student_results,
                            cfg                = cfg_s,
                            pos_cols           = pos_cols_s,
                            selected_exam_data = exam_data_for_pdf,
                        )
                        st.session_state.pdf_results[exam_label] = pdf_buf.getvalue()

            # Show download buttons for all generated PDFs
            if st.session_state.get("pdf_results"):
                st.markdown("#### 📥 Download PDFs")
                cols = st.columns(len(st.session_state.pdf_results))
                for col, (exam_label, pdf_bytes) in zip(
                    cols, st.session_state.pdf_results.items()
                ):
                    safe_name = exam_label.replace("/", "-").replace(" ", "_")
                    n_students = len(student_results)
                    col.download_button(
                        label=f"📥 {exam_label}",
                        data=pdf_bytes,
                        file_name=f"Results_{safe_name}.pdf",
                        mime="application/pdf",
                        key=f"dl_pdf_{safe_name}",
                    )
                    col.caption(f"{n_students} students · {-(-n_students//2)} pages")

    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())
