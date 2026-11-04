import io
import random
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# ==============================================================================
# 1. PAGE CONFIG & SESSION STATE INITIALIZATION
# ==============================================================================
st.set_page_config(
    page_title="Innoinf Technologies & Envahan — Salary Simulator & Offer Studio",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_STATE = {
    "company_name": "INNOINF TECHNOLOGIES PRIVATE LIMITED",
    "company_address": "3-684 Sakkavayal, Karaikudi, Sivaganga - 630108, Tamil Nadu",
    "brand_unit": "Envahan Electric Vehicles",
    "emp_name": "Arjun Natarajan",
    "designation": "Senior Lead Engineer - EV Systems",
    "grade": "L3 - Mid-Management",
    "work_state": "Tamil Nadu",
    "city_type": "Metro (50% Basic)",
    "joining_date": "01-Nov-2026",
    "calc_mode": "Forward (Annual CTC ➔ Monthly In-Hand)",
    "annual_ctc": 1500000,
    "target_monthly_net": 105000,
    "var_pct": 10.0,
    "basic_pct": 50.0,
    "hra_pct": 50.0,
    # PF Mode & Custom Minimum Value Input
    "pf_mode": "Statutory Ceiling (₹15k Basic ➔ ₹1,800/m)",
    "custom_pf_min": 1800.0,
    "include_gratuity": True,
    "nps_emp_pct": 5.0,
    "regime_pref": "Auto-Recommend Best",
    # Flexi Benefit Basket (Annual)
    "fbp_lta": 36000,
    "fbp_telecom": 24000,
    "fbp_meal": 26400,
    "fbp_books": 12000,
    # Tax Declarations (Old Regime / Comparison)
    "monthly_rent": 25000,
    "sec_80c_other": 120000,
    "sec_80ccd1b": 50000,
    "sec_80d": 35000,
    "sec_24b": 0,
    # Hike & Appraisal
    "hike_pct": 20.0,
    "rev_var_pct": 10.0,
    "arrear_months": 2,
    "perf_rating": "4 - Exceeds Expectations",
    # Offer Letter Terms
    "probation_months": 6,
    "notice_period_days": 60,
    "signatory_name": "Karthik Subramanian",
    "signatory_title": "Director & Head of Engineering",
}

for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ==============================================================================
# 2. CALLBACKS & GENERATORS
# ==============================================================================
def randomize_ctc_only():
    if random.random() < 0.65:
        random_ctc = random.randint(240000, 5500000)
    else:
        random_ctc = random.randrange(250000, 6000000, 25000)
    st.session_state.annual_ctc = int(random_ctc)
    max_safe_fbp = int(random_ctc * 0.08)
    if max_safe_fbp < 98400:
        st.session_state.fbp_lta = int(max_safe_fbp * 0.35)
        st.session_state.fbp_telecom = int(max_safe_fbp * 0.25)
        st.session_state.fbp_meal = int(max_safe_fbp * 0.25)
        st.session_state.fbp_books = int(max_safe_fbp * 0.15)


def randomize_full_profile():
    sample_names = [
        ("Arjun Natarajan", "Senior Lead Engineer - EV Systems", "L3 - Mid-Management"),
        ("Divya Krishnan", "Embedded Firmware Architect", "L4 - Leadership & Exec"),
        ("Prashanth V", "Supply Chain & Sourcing Executive", "L2 - Individual Contributor"),
        ("Meenakshi Sundaram", "Graduate Engineer Trainee (GET)", "L1 - Entry & Support"),
        ("Rohan Deshmukh", "Regional Sales Manager", "L3 - Mid-Management"),
    ]
    name, role, grade = random.choice(sample_names)
    st.session_state.emp_name = name
    st.session_state.designation = role
    st.session_state.grade = grade
    st.session_state.work_state = random.choice(["Tamil Nadu", "Karnataka", "Maharashtra", "Telangana", "Delhi / NCR / Haryana"])
    st.session_state.city_type = random.choice(["Metro (50% Basic)", "Non-Metro (40% Basic)"])
    st.session_state.annual_ctc = random.randint(320000, 4850000)
    st.session_state.var_pct = float(random.choice([0, 5, 10, 15, 20]))
    st.session_state.basic_pct = float(random.choice([50, 50, 55]))
    st.session_state.hra_pct = 50.0 if "Metro" in st.session_state.city_type else 40.0
    st.session_state.pf_mode = random.choice([
        "Statutory Ceiling (₹15k Basic ➔ ₹1,800/m)",
        "Uncapped (12% of Actual Basic)",
        "Custom Minimum Value Threshold",
    ])
    st.session_state.nps_emp_pct = float(random.choice([0, 5, 10, 14]))
    fixed_ctc_est = st.session_state.annual_ctc * (1 - st.session_state.var_pct / 100.0)
    fbp_budget = min(120000, int(fixed_ctc_est * 0.08))
    st.session_state.fbp_lta = round(fbp_budget * 0.40, -2)
    st.session_state.fbp_telecom = round(fbp_budget * 0.25, -2)
    st.session_state.fbp_meal = round(fbp_budget * 0.20, -2)
    st.session_state.fbp_books = round(fbp_budget * 0.15, -2)
    st.session_state.monthly_rent = round((fixed_ctc_est / 12) * 0.22, -2)


def set_preset_ctc(val: int):
    st.session_state.annual_ctc = int(val)
    if val <= 400000:
        st.session_state.fbp_lta = 0
        st.session_state.fbp_telecom = 0
        st.session_state.fbp_meal = 0
        st.session_state.fbp_books = 0


# ==============================================================================
# 3. CORE PAYROLL & TAX ENGINE (FY 2026-27 / AY 2027-28)
# ==============================================================================
def compute_new_regime_tax(taxable_inc: float) -> dict:
    ti = max(0.0, round(taxable_inc, -1))
    raw_tax = (
        max(0.0, ti - 400000) * 0.05
        + max(0.0, ti - 800000) * 0.05
        + max(0.0, ti - 1200000) * 0.05
        + max(0.0, ti - 1600000) * 0.05
        + max(0.0, ti - 2000000) * 0.05
        + max(0.0, ti - 2400000) * 0.05
    )
    raw_tax = round(raw_tax)
    rebate_87a = raw_tax if ti <= 1200000 else 0
    marginal_relief = 0
    if ti > 1200000 and (raw_tax - rebate_87a) > (ti - 1200000):
        marginal_relief = (raw_tax - rebate_87a) - (ti - 1200000)

    tax_after_relief = max(0, raw_tax - rebate_87a - marginal_relief)
    cess = round(tax_after_relief * 0.04)
    total_tax = tax_after_relief + cess
    return {
        "taxable_income": ti,
        "raw_tax": raw_tax,
        "rebate_87a": rebate_87a,
        "marginal_relief": marginal_relief,
        "cess": cess,
        "total_tax": total_tax,
    }


def compute_old_regime_tax(taxable_inc: float) -> dict:
    ti = max(0.0, round(taxable_inc, -1))
    raw_tax = (
        max(0.0, ti - 250000) * 0.05
        + max(0.0, ti - 500000) * 0.15
        + max(0.0, ti - 1000000) * 0.10
    )
    raw_tax = round(raw_tax)
    rebate_87a = raw_tax if ti <= 500000 else 0
    tax_after_relief = max(0, raw_tax - rebate_87a)
    cess = round(tax_after_relief * 0.04)
    total_tax = tax_after_relief + cess
    return {
        "taxable_income": ti,
        "raw_tax": raw_tax,
        "rebate_87a": rebate_87a,
        "marginal_relief": 0,
        "cess": cess,
        "total_tax": total_tax,
    }


def calculate_salary_structure(
    annual_ctc: float,
    var_pct: float,
    basic_pct: float,
    hra_pct: float,
    pf_mode: str,
    custom_pf_min: float,
    include_gratuity: bool,
    nps_emp_pct: float,
    work_state: str,
    city_type: str,
    fbp_dict: dict,
    tax_inputs: dict,
    regime_pref: str,
) -> dict:
    total_ctc = float(annual_ctc)
    variable_pay = round(total_ctc * (var_pct / 100.0))
    fixed_ctc = total_ctc - variable_pay

    basic_ann = round(fixed_ctc * (basic_pct / 100.0))
    basic_mo = round(basic_ann / 12.0)

    hra_ann = round(basic_ann * (hra_pct / 100.0))
    hra_mo = round(hra_ann / 12.0)

    # PF Calculation with Custom Minimum Threshold or 12% Toggle
    if pf_mode.startswith("Statutory"):
        emp_pf_mo = round(min(basic_mo, 15000) * 0.12)
    elif pf_mode.startswith("Revised"):
        emp_pf_mo = round(min(basic_mo, 25000) * 0.12)
    elif pf_mode.startswith("Uncapped"):
        emp_pf_mo = round(basic_mo * 0.12)
    elif pf_mode.startswith("Custom"):
        emp_pf_mo = round(max(custom_pf_min, min(basic_mo * 0.12, basic_mo)))
    else:
        emp_pf_mo = 0
    emp_pf_ann = emp_pf_mo * 12

    gratuity_ann = round(basic_ann * (15.0 / 26.0) / 12.0) if include_gratuity else 0
    nps_ann = round(basic_ann * (nps_emp_pct / 100.0))

    prelim_gross_ann = fixed_ctc - (emp_pf_ann + gratuity_ann + nps_ann)
    prelim_gross_mo = prelim_gross_ann / 12.0

    if prelim_gross_mo <= 21000:
        gross_ann = round(prelim_gross_ann / 1.0325)
        employer_esi_ann = round(gross_ann * 0.0325)
        employee_esi_ann = round(gross_ann * 0.0075)
        esic_applicable = True
    else:
        gross_ann = prelim_gross_ann
        employer_esi_ann = 0
        employee_esi_ann = 0
        esic_applicable = False

    gross_mo = round(gross_ann / 12.0)
    employer_retirals_ann = emp_pf_ann + employer_esi_ann + gratuity_ann + nps_ann

    requested_fbp = sum(fbp_dict.values())
    available_for_fbp_and_special = gross_ann - (basic_ann + hra_ann)

    fbp_scaled_warning = False
    if requested_fbp > max(0, available_for_fbp_and_special):
        fbp_scaled_warning = True
        scale_ratio = max(0.0, available_for_fbp_and_special) / requested_fbp if requested_fbp > 0 else 0.0
        lta_ann = round(fbp_dict["lta"] * scale_ratio)
        telecom_ann = round(fbp_dict["telecom"] * scale_ratio)
        meal_ann = round(fbp_dict["meal"] * scale_ratio)
        books_ann = round(fbp_dict["books"] * scale_ratio)
    else:
        lta_ann = fbp_dict["lta"]
        telecom_ann = fbp_dict["telecom"]
        meal_ann = fbp_dict["meal"]
        books_ann = fbp_dict["books"]

    total_fbp_ann = lta_ann + telecom_ann + meal_ann + books_ann
    special_ann = gross_ann - (basic_ann + hra_ann + total_fbp_ann)

    if gross_mo <= 15000 or work_state == "Delhi / NCR / Haryana":
        pt_ann = 0
    elif work_state in ["Tamil Nadu", "Maharashtra"]:
        pt_ann = 2500
    else:
        pt_ann = 2400

    gross_taxable_salary = gross_ann + variable_pay + nps_ann

    nps_ded_new = min(nps_ann, round(basic_ann * 0.14))
    ded_new = 75000 + nps_ded_new
    new_tax_res = compute_new_regime_tax(gross_taxable_salary - ded_new)

    ann_rent = tax_inputs["monthly_rent"] * 12
    hra_rule_1 = hra_ann
    hra_rule_2 = max(0, ann_rent - round(basic_ann * 0.10))
    hra_rule_3 = round(basic_ann * (0.50 if "Metro" in city_type else 0.40))
    hra_exempt = min(hra_rule_1, hra_rule_2, hra_rule_3) if ann_rent > 0 else 0

    sec_80c_eff = min(150000, emp_pf_ann + tax_inputs["sec_80c_other"])
    sec_80ccd1b_eff = min(50000, tax_inputs["sec_80ccd1b"])
    sec_80d_eff = min(100000, tax_inputs["sec_80d"])
    sec_24b_eff = min(200000, tax_inputs["sec_24b"])
    nps_ded_old = min(nps_ann, round(basic_ann * 0.10))

    ded_old = (
        50000
        + hra_exempt
        + total_fbp_ann
        + pt_ann
        + nps_ded_old
        + sec_80c_eff
        + sec_80ccd1b_eff
        + sec_80d_eff
        + sec_24b_eff
    )
    old_tax_res = compute_old_regime_tax(gross_taxable_salary - ded_old)

    if regime_pref.startswith("New"):
        active_regime = "New Tax Regime"
        active_tax_ann = new_tax_res["total_tax"]
    elif regime_pref.startswith("Old"):
        active_regime = "Old Tax Regime"
        active_tax_ann = old_tax_res["total_tax"]
    else:
        if new_tax_res["total_tax"] <= old_tax_res["total_tax"]:
            active_regime = "New Tax Regime (Auto-Best)"
            active_tax_ann = new_tax_res["total_tax"]
        else:
            active_regime = "Old Tax Regime (Auto-Best)"
            active_tax_ann = old_tax_res["total_tax"]

    total_emp_ded_ann = emp_pf_ann + employee_esi_ann + pt_ann + active_tax_ann
    net_fixed_ann = gross_ann - total_emp_ded_ann
    net_fixed_mo = round(net_fixed_ann / 12.0)

    return {
        "total_ctc": total_ctc,
        "variable_pay": variable_pay,
        "fixed_ctc": fixed_ctc,
        "basic_ann": basic_ann,
        "basic_mo": basic_mo,
        "hra_ann": hra_ann,
        "hra_mo": hra_mo,
        "lta_ann": lta_ann,
        "telecom_ann": telecom_ann,
        "meal_ann": meal_ann,
        "books_ann": books_ann,
        "total_fbp_ann": total_fbp_ann,
        "special_ann": special_ann,
        "special_mo": round(special_ann / 12.0),
        "gross_ann": gross_ann,
        "gross_mo": gross_mo,
        "emp_pf_ann": emp_pf_ann,
        "emp_pf_mo": emp_pf_mo,
        "employer_esi_ann": employer_esi_ann,
        "employee_esi_ann": employee_esi_ann,
        "esic_applicable": esic_applicable,
        "gratuity_ann": gratuity_ann,
        "nps_ann": nps_ann,
        "employer_retirals_ann": employer_retirals_ann,
        "pt_ann": pt_ann,
        "pt_mo": round(pt_ann / 12.0),
        "hra_exempt": hra_exempt,
        "hra_rules": (hra_rule_1, hra_rule_2, hra_rule_3),
        "sec_80c_eff": sec_80c_eff,
        "ded_new": ded_new,
        "ded_old": ded_old,
        "new_tax": new_tax_res,
        "old_tax": old_tax_res,
        "active_regime": active_regime,
        "active_tax_ann": active_tax_ann,
        "active_tax_mo": round(active_tax_ann / 12.0),
        "total_emp_ded_ann": total_emp_ded_ann,
        "net_fixed_ann": net_fixed_ann,
        "net_fixed_mo": net_fixed_mo,
        "net_total_ann": net_fixed_ann + variable_pay,
        "fbp_scaled_warning": fbp_scaled_warning,
        "wage_code_ok": (basic_ann >= 0.50 * fixed_ctc - 10),
    }


def solve_ctc_from_target_net(target_mo_net: float, **kwargs) -> int:
    low, high = 120000.0, 20000000.0
    for _ in range(28):
        mid = (low + high) / 2.0
        res = calculate_salary_structure(annual_ctc=mid, **kwargs)
        if res["net_fixed_mo"] < target_mo_net:
            low = mid
        else:
            high = mid
    return int(round((low + high) / 2.0))


# ==============================================================================
# 4. WORD (.DOCX) OFFER LETTER & EXCEL WORKBOOK GENERATORS
# ==============================================================================
def build_word_offer_letter_bytes(s: dict, res: dict) -> bytes:
    doc = docx.Document()

    # Page Margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Company Header Banner
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_comp = p_head.add_run(f"{s['company_name']}\n")
    run_comp.bold = True
    run_comp.font.size = Pt(14)
    run_comp.font.color.rgb = RGBColor(27, 54, 93)

    run_sub = p_head.add_run(f"{s['company_address']}\nBrand Unit: {s['brand_unit']}\n")
    run_sub.font.size = Pt(9.5)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_div = p_div.add_run("____________________________________________________________________")
    run_div.font.color.rgb = RGBColor(0, 102, 102)

    # Date & Candidate Info
    doc.add_paragraph(f"Date: {s['joining_date']}")
    p_to = doc.add_paragraph()
    p_to.add_run(f"To,\n{s['emp_name']}\nDesignation: {s['designation']}\nGrade: {s['grade']}").bold = True

    doc.add_heading("Sub: Offer of Employment & Compensation Structure", level=2)

    p_body = doc.add_paragraph(
        f"Dear {s['emp_name']},\n\n"
        f"We are delighted to offer you employment with {s['company_name']} ({s['brand_unit']}) "
        f"in the position of {s['designation']}. Your annual Cost to Company (CTC) will be "
        f"₹{res['total_ctc']:,.0f} (Rupees {number_to_words(res['total_ctc'])} Only), payable on a monthly basis "
        f"effective from your date of joining on or before {s['joining_date']}.\n\n"
        f"Your employment will be subject to a probationary period of {s['probation_months']} months and "
        f"governed by the standard terms of employment, confidentiality, and company policies."
    )

    doc.add_heading("Annexure A: Salary Structure Breakup", level=3)

    # Add Salary Table
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Salary Component"
    hdr_cells[1].text = "Monthly (₹)"
    hdr_cells[2].text = "Annual (₹)"

    table_data = [
        ("Basic Salary (50% Wage Code Anchor)", f"₹{res['basic_mo']:,.0f}", f"₹{res['basic_ann']:,.0f}"),
        ("House Rent Allowance (HRA)", f"₹{res['hra_mo']:,.0f}", f"₹{res['hra_ann']:,.0f}"),
        ("Leave Travel Allowance (LTA)", f"₹{round(res['lta_ann']/12):,.0f}", f"₹{res['lta_ann']:,.0f}"),
        ("Telecom & Broadband Allowance", f"₹{round(res['telecom_ann']/12):,.0f}", f"₹{res['telecom_ann']:,.0f}"),
        ("Food Vouchers & Periodicals", f"₹{round((res['meal_ann']+res['books_ann'])/12):,.0f}", f"₹{res['meal_ann']+res['books_ann']:,.0f}"),
        ("Special Allowance (Balancing Plug)", f"₹{res['special_mo']:,.0f}", f"₹{res['special_ann']:,.0f}"),
        ("GROSS FIXED SALARY (A)", f"₹{res['gross_mo']:,.0f}", f"₹{res['gross_ann']:,.0f}"),
        ("Employer Provident Fund (EPF)", f"₹{res['emp_pf_mo']:,.0f}", f"₹{res['emp_pf_ann']:,.0f}"),
        ("Employer ESIC Contribution", f"₹{round(res['employer_esi_ann']/12):,.0f}", f"₹{res['employer_esi_ann']:,.0f}"),
        ("Employer Gratuity Accrual", f"₹{round(res['gratuity_ann']/12):,.0f}", f"₹{res['gratuity_ann']:,.0f}"),
        ("Employer Corporate NPS u/s 80CCD(2)", f"₹{round(res['nps_ann']/12):,.0f}", f"₹{res['nps_ann']:,.0f}"),
        ("FIXED COST TO COMPANY (Fixed CTC = A + Retirals)", f"₹{round(res['fixed_ctc']/12):,.0f}", f"₹{res['fixed_ctc']:,.0f}"),
        ("Performance Linked Variable Pay", f"₹{round(res['variable_pay']/12):,.0f}", f"₹{res['variable_pay']:,.0f}"),
        ("TOTAL COST TO COMPANY (TOTAL CTC)", f"₹{round(res['total_ctc']/12):,.0f}", f"₹{res['total_ctc']:,.0f}"),
    ]

    for comp, m_val, a_val in table_data:
        row_cells = table.add_row().cells
        row_cells[0].text = comp
        row_cells[1].text = m_val
        row_cells[2].text = a_val

    # Table styling
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9.5)
            if i == 0:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    doc.add_paragraph("\n")
    p_sign = doc.add_paragraph()
    p_sign.add_run(
        f"For {s['company_name']}\n\n\n"
        f"{s['signatory_name']}\n"
        f"{s['signatory_title']}\n\n\n"
        f"---------------------------------------------------\n"
        f"I accept the terms and conditions of this offer:\n"
        f"Candidate Signature & Date: ______________________"
    )

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def number_to_words(num: int) -> str:
    # Helper for formal letter wording
    if num >= 10000000:
        return f"{num // 10000000} Crore {number_to_words(num % 10000000)}"
    elif num >= 100000:
        return f"{num // 100000} Lakh {number_to_words(num % 100000)}"
    elif num >= 1000:
        return f"{num // 1000} Thousand {number_to_words(num % 1000)}"
    else:
        return str(num)


def build_excel_workbook_bytes(s: dict, res: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws_offer = wb.active
    ws_offer.title = "1_Offer_Letter_Annexure"
    ws_tax = wb.create_sheet("2_Tax_Comparison_FY26_27")
    ws_reg = wb.create_sheet("3_12M_Payroll_Schedule")

    navy_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    teal_fill = PatternFill(start_color="006666", end_color="006666", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    total_fill = PatternFill(start_color="E6FFFA", end_color="E6FFFA", fill_type="solid")

    f_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    f_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_bold = Font(name="Calibri", size=10, bold=True, color="1A202C")
    f_reg = Font(name="Calibri", size=10, bold=False, color="1A202C")
    thin_bdr = Border(
        left=Side(style="thin", color="CBD5E0"),
        right=Side(style="thin", color="CBD5E0"),
        top=Side(style="thin", color="CBD5E0"),
        bottom=Side(style="thin", color="CBD5E0"),
    )
    inr_fmt = "₹#,##0"

    # --- Sheet 1: Offer Letter Annexure ---
    ws_offer.merge_cells("A1:D1")
    ws_offer["A1"] = f"{s['company_name']} | {s['company_address']}"
    ws_offer["A1"].font = f_title
    ws_offer["A1"].fill = navy_fill
    ws_offer["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_offer.row_dimensions[1].height = 25

    ws_offer.merge_cells("A2:D2")
    ws_offer["A2"] = "ANNEXURE A: COMPENSATION STRUCTURE BREAKUP"
    ws_offer["A2"].font = f_title
    ws_offer["A2"].fill = teal_fill
    ws_offer["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_offer.row_dimensions[2].height = 22

    meta_rows = [
        ("Candidate Name", s["emp_name"], "Designation", s["designation"]),
        ("Grade / Band", s["grade"], "Work State & City", f"{s['work_state']} ({s['city_type']})"),
        ("Effective Date", s["joining_date"], "Active Tax Regime", res["active_regime"]),
    ]
    for idx, (k1, v1, k2, v2) in enumerate(meta_rows, start=4):
        ws_offer[f"A{idx}"], ws_offer[f"B{idx}"] = k1, v1
        ws_offer[f"C{idx}"], ws_offer[f"D{idx}"] = k2, v2
        for c in ["A", "B", "C", "D"]:
            ws_offer[f"{c}{idx}"].border = thin_bdr
            ws_offer[f"{c}{idx}"].font = f_bold if c in ["A", "C"] else f_reg

    headers = ["Salary Component", "Monthly (₹)", "Annual (₹)", "Statutory & Tax Notes"]
    for col_i, h_text in enumerate(headers, start=1):
        c_cell = ws_offer.cell(row=8, column=col_i, value=h_text)
        c_cell.font = f_hdr
        c_cell.fill = teal_fill
        c_cell.border = thin_bdr

    offer_data = [
        ("Basic Salary", round(res["basic_ann"] / 12), res["basic_ann"], "50% Wage Code Compliant"),
        ("House Rent Allowance (HRA)", round(res["hra_ann"] / 12), res["hra_ann"], "Sec 10(13A) Tax Exempt in Old Regime"),
        ("Leave Travel Allowance (LTA)", round(res["lta_ann"] / 12), res["lta_ann"], "Flexi-Basket Component"),
        ("Telecom & Internet Allowance", round(res["telecom_ann"] / 12), res["telecom_ann"], "Flexi-Basket Bill Reimbursement"),
        ("Food Coupons & Books", round((res["meal_ann"] + res["books_ann"]) / 12), res["meal_ann"] + res["books_ann"], "Tax-Exempt Meal Vouchers & Books"),
        ("Special Allowance", round(res["special_ann"] / 12), res["special_ann"], "Balancing Fixed Component"),
        ("GROSS FIXED SALARY (A)", "=SUM(B9:B14)", "=SUM(C9:C14)", "Monthly & Annual Fixed Gross"),
        ("Employer EPF Contribution", round(res["emp_pf_ann"] / 12), res["emp_pf_ann"], s["pf_mode"]),
        ("Employer ESIC (3.25%)", round(res["employer_esi_ann"] / 12), res["employer_esi_ann"], "Applicable if Gross <= ₹21,000/m"),
        ("Employer Gratuity Accrual (4.81%)", round(res["gratuity_ann"] / 12), res["gratuity_ann"], "Payment of Gratuity Act"),
        ("Employer Corporate NPS u/s 80CCD(2)", round(res["nps_ann"] / 12), res["nps_ann"], "Deductible up to 14% Basic in New Regime"),
        ("FIXED COST TO COMPANY (Fixed CTC)", "=B15+SUM(B16:B19)", "=C15+SUM(C16:C19)", "Guaranteed Annual Fixed CTC"),
        ("Performance Linked Variable Pay", round(res["variable_pay"] / 12), res["variable_pay"], f"{s['var_pct']:.1f}% of Total CTC"),
        ("TOTAL COST TO COMPANY (TOTAL CTC)", "=B20+B21", "=C20+C21", "100% Total Package"),
        ("Less: Employee EPF + ESIC + PT", round((res["emp_pf_ann"] + res["employee_esi_ann"] + res["pt_ann"]) / 12), res["emp_pf_ann"] + res["employee_esi_ann"] + res["pt_ann"], "Statutory Deductions"),
        ("Less: Income Tax (Monthly TDS)", res["active_tax_mo"], res["active_tax_ann"], res["active_regime"]),
        ("ESTIMATED NET MONTHLY TAKE-HOME (FIXED)", "=B15-B23-B24", "=C15-C23-C24", "Net Bank Credit (Excl. Variable)"),
    ]

    for r_idx, (comp, m_val, a_val, note) in enumerate(offer_data, start=9):
        ws_offer.cell(row=r_idx, column=1, value=comp)
        ws_offer.cell(row=r_idx, column=2, value=m_val).number_format = inr_fmt
        ws_offer.cell(row=r_idx, column=3, value=a_val).number_format = inr_fmt
        ws_offer.cell(row=r_idx, column=4, value=note)
        is_tot = "GROSS" in comp or "COST TO COMPANY" in comp or "NET MONTHLY" in comp
        for c_i in range(1, 5):
            cell = ws_offer.cell(row=r_idx, column=c_i)
            cell.border = thin_bdr
            cell.font = f_bold if is_tot else f_reg
            cell.fill = total_fill if is_tot else zebra_fill

    # --- Sheet 2: Tax Planner Summary ---
    ws_tax["A1"] = f"{s['company_name']} — DUAL-REGIME TAX COMPARISON (FY 2026-27)"
    ws_tax["A1"].font = f_bold
    tax_rows = [
        ("Parameter", "New Tax Regime (₹)", "Old Tax Regime (₹)"),
        ("Gross Taxable Salary (Incl. Bonus & NPS)", res["gross_ann"] + res["variable_pay"] + res["nps_ann"], res["gross_ann"] + res["variable_pay"] + res["nps_ann"]),
        ("Total Exemptions & Deductions", res["ded_new"], res["ded_old"]),
        ("Net Taxable Income", res["new_tax"]["taxable_income"], res["old_tax"]["taxable_income"]),
        ("Gross Slab Tax", res["new_tax"]["raw_tax"], res["old_tax"]["raw_tax"]),
        ("Less: Rebate u/s 87A & Marginal Relief", res["new_tax"]["rebate_87a"] + res["new_tax"]["marginal_relief"], res["old_tax"]["rebate_87a"]),
        ("Health & Education Cess (4%)", res["new_tax"]["cess"], res["old_tax"]["cess"]),
        ("TOTAL ANNUAL INCOME TAX (TDS)", res["new_tax"]["total_tax"], res["old_tax"]["total_tax"]),
        ("Monthly TDS Deduction", round(res["new_tax"]["total_tax"] / 12), round(res["old_tax"]["total_tax"] / 12)),
    ]
    for r_i, row_vals in enumerate(tax_rows, start=3):
        for c_i, val in enumerate(row_vals, start=1):
            c = ws_tax.cell(row=r_i, column=c_i, value=val)
            c.border = thin_bdr
            c.font = f_bold if r_i in [3, 10] else f_reg
            if c_i > 1 and isinstance(val, (int, float)):
                c.number_format = inr_fmt

    # --- Sheet 3: 12-Month Payroll Schedule ---
    months = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
    ws_reg.cell(row=1, column=1, value="Pay Component").font = f_bold
    for m_i, m_name in enumerate(months, start=2):
        ws_reg.cell(row=1, column=m_i, value=m_name).font = f_bold
    ws_reg.cell(row=1, column=14, value="Annual Total").font = f_bold

    sched_items = [
        ("Basic Salary", res["basic_mo"]),
        ("HRA", res["hra_mo"]),
        ("Flexi-Basket Allowances", round(res["total_fbp_ann"] / 12)),
        ("Special Allowance", res["special_mo"]),
        ("Gross Monthly Salary", res["gross_mo"]),
        ("Employee EPF + ESIC + PT", round((res["emp_pf_ann"] + res["employee_esi_ann"] + res["pt_ann"]) / 12)),
        ("Monthly Income Tax (TDS)", res["active_tax_mo"]),
        ("Net Monthly In-Hand Salary", res["net_fixed_mo"]),
    ]
    for r_i, (lbl, m_val) in enumerate(sched_items, start=2):
        ws_reg.cell(row=r_i, column=1, value=lbl).font = f_bold if "Net" in lbl or "Gross" in lbl else f_reg
        for m_col in range(2, 14):
            c = ws_reg.cell(row=r_i, column=m_col, value=m_val)
            c.number_format = inr_fmt
            c.border = thin_bdr
        tot_c = ws_reg.cell(row=r_i, column=14, value=f"=SUM(B{r_i}:M{r_i})")
        tot_c.number_format = inr_fmt
        tot_c.font = f_bold

    for ws in [ws_offer, ws_tax, ws_reg]:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = 28

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ==============================================================================
# 5. SIDEBAR: MASTER CONTROLS & OFFER STUDIO SETTINGS
# ==============================================================================
with st.sidebar:
    st.markdown("### 🏢 Innoinf | Envahan HRMS")
    st.caption("Official Salary & Offer Letter Studio")

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        st.button("🎲 Random CTC", on_click=randomize_ctc_only, use_container_width=True)
    with col_btn2:
        st.button("🔀 Random Profile", on_click=randomize_full_profile, use_container_width=True)

    st.markdown("**Quick CTC Presets:**")
    p_cols = st.columns(4)
    p_cols[0].button("₹3.6L", on_click=set_preset_ctc, args=(360000,), use_container_width=True)
    p_cols[1].button("₹12L", on_click=set_preset_ctc, args=(1200000,), use_container_width=True)
    p_cols[2].button("₹18L", on_click=set_preset_ctc, args=(1800000,), use_container_width=True)
    p_cols[3].button("₹32L", on_click=set_preset_ctc, args=(3200000,), use_container_width=True)

    st.divider()

    st.radio(
        "Calculation Mode",
        ["Forward (Annual CTC ➔ Monthly In-Hand)", "Reverse (Target Monthly Net ➔ Required CTC)"],
        key="calc_mode",
    )

    if st.session_state.calc_mode.startswith("Forward"):
        st.number_input(
            "Annual Total CTC (₹) — Type Any Value",
            min_value=100000,
            max_value=100000000,
            step=1,
            key="annual_ctc",
        )
    else:
        st.number_input(
            "Target Monthly Net In-Hand (₹)",
            min_value=10000,
            max_value=5000000,
            step=1,
            key="target_monthly_net",
        )

    st.slider("Variable / Bonus (% of CTC)", 0.0, 40.0, step=0.5, key="var_pct")
    st.slider("Basic Salary (% of Fixed CTC)", 40.0, 70.0, step=1.0, key="basic_pct")
    st.slider("HRA (% of Basic Salary)", 30.0, 60.0, step=5.0, key="hra_pct")

    st.selectbox(
        "PF Calculation Policy",
        [
            "Statutory Ceiling (₹15k Basic ➔ ₹1,800/m)",
            "Uncapped (12% of Actual Basic)",
            "Custom Minimum Value Threshold",
        ],
        key="pf_mode",
    )

    if st.session_state.pf_mode.startswith("Custom"):
        st.number_input("Custom Minimum PF Value (₹/month)", 500, 20000, step=100, key="custom_pf_min")

    st.checkbox("Include Gratuity (4.81% of Basic) in CTC", key="include_gratuity")
    st.slider("Employer Corporate NPS u/s 80CCD(2) (% Basic)", 0.0, 14.0, step=1.0, key="nps_emp_pct")
    st.selectbox(
        "Work State (Professional Tax)",
        ["Tamil Nadu", "Karnataka", "Maharashtra", "Telangana", "Delhi / NCR / Haryana"],
        key="work_state",
    )
    st.selectbox("City Category (HRA Rule)", ["Metro (50% Basic)", "Non-Metro (40% Basic)"], key="city_type")
    st.selectbox(
        "Preferred Tax Regime",
        ["Auto-Recommend Best", "New Tax Regime (FY 26-27)", "Old Tax Regime"],
        key="regime_pref",
    )


# ==============================================================================
# 6. RUN ENGINE
# ==============================================================================
common_kwargs = dict(
    var_pct=st.session_state.var_pct,
    basic_pct=st.session_state.basic_pct,
    hra_pct=st.session_state.hra_pct,
    pf_mode=st.session_state.pf_mode,
    custom_pf_min=st.session_state.custom_pf_min,
    include_gratuity=st.session_state.include_gratuity,
    nps_emp_pct=st.session_state.nps_emp_pct,
    work_state=st.session_state.work_state,
    city_type=st.session_state.city_type,
    fbp_dict={
        "lta": st.session_state.fbp_lta,
        "telecom": st.session_state.fbp_telecom,
        "meal": st.session_state.fbp_meal,
        "books": st.session_state.fbp_books,
    },
    tax_inputs={
        "monthly_rent": st.session_state.monthly_rent,
        "sec_80c_other": st.session_state.sec_80c_other,
        "sec_80ccd1b": st.session_state.sec_80ccd1b,
        "sec_80d": st.session_state.sec_80d,
        "sec_24b": st.session_state.sec_24b,
    },
    regime_pref=st.session_state.regime_pref,
)

if st.session_state.calc_mode.startswith("Reverse"):
    solved_ctc = solve_ctc_from_target_net(st.session_state.target_monthly_net, **common_kwargs)
    st.session_state.annual_ctc = solved_ctc

res = calculate_salary_structure(annual_ctc=st.session_state.annual_ctc, **common_kwargs)


# ==============================================================================
# 7. MAIN DASHBOARD HEADER & LETTERHEAD BRANDING
# ==============================================================================
st.markdown(
    f"""
    <div style="background-color: #1B365D; padding: 16px; border-radius: 8px; color: white; text-align: center; margin-bottom: 20px;">
        <h2 style="margin: 0; color: #FFFFFF; font-size: 22px;">{st.session_state.company_name}</h2>
        <p style="margin: 4px 0 0 0; font-size: 13px; color: #CBD5E0;">{st.session_state.company_address} | Brand Unit: {st.session_state.brand_unit}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Annual Total CTC", f"₹{res['total_ctc']:,.0f}", f"Fixed: ₹{res['fixed_ctc']:,.0f}")
kpi2.metric("Monthly Gross Salary", f"₹{res['gross_mo']:,.0f}", f"Annual: ₹{res['gross_ann']:,.0f}")
kpi3.metric("Monthly Net Take-Home (Fixed)", f"₹{res['net_fixed_mo']:,.0f}", f"{res['net_fixed_ann'] / res['fixed_ctc'] * 100:.1f}% of Fixed")
tax_diff = abs(res["new_tax"]["total_tax"] - res["old_tax"]["total_tax"])
kpi4.metric("Annual Income Tax (TDS)", f"₹{res['active_tax_ann']:,.0f}", f"Saves ₹{tax_diff:,.0f} ({res['active_regime']})")

c_a1, c_a2, c_a3 = st.columns(3)
with c_a1:
    if res["wage_code_ok"]:
        st.success("✅ **Code on Wages (50% Rule):** Compliant")
    else:
        st.warning("⚠️ **Code on Wages Alert:** Basic < 50% of Fixed CTC")
with c_a2:
    if res["esic_applicable"]:
        st.info("🏥 **ESIC Applicable:** Gross ≤ ₹21k/mo")
    else:
        st.success("✅ **ESIC Exempt:** Gross > ₹21k/mo")
with c_a3:
    if res["fbp_scaled_warning"]:
        st.warning("⚠️ **FBP Auto-Scaled:** Allowances adjusted to prevent negative Special Allowance.")
    else:
        st.success(f"✅ **Special Allowance Plug:** ₹{res['special_mo']:,.0f}/mo")


# ==============================================================================
# 8. FOUR WORKSPACE TABS (INCLUDING DEDICATED OFFER LETTER STUDIO)
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Structure & FBP Designer",
    "⚖️ 2. Dual-Regime Tax Planner",
    "📈 3. Appraisal & Hike Modeler",
    "📄 4. Offer Letter Studio & Word/Excel Export",
])

# --- TAB 1: STRUCTURE & FBP ---
with tab1:
    col_l, col_r = st.columns([1, 2])
    with col_l:
        st.subheader("Flexi-Benefit Basket (Annual ₹)")
        st.number_input("Leave Travel Allowance (LTA)", 0, 300000, step=1000, key="fbp_lta")
        st.number_input("Telecom & Broadband", 0, 120000, step=1000, key="fbp_telecom")
        st.number_input("Food / Meal Coupons", 0, 50000, step=600, key="fbp_meal")
        st.number_input("Books & Periodicals", 0, 120000, step=1000, key="fbp_books")

    with col_r:
        st.subheader("Live Master CTC Breakdown")
        struct_df = pd.DataFrame([
            ("Basic Salary", res["basic_mo"], res["basic_ann"], "Earnings"),
            ("House Rent Allowance (HRA)", res["hra_mo"], res["hra_ann"], "Earnings"),
            ("Leave Travel Allowance (LTA)", round(res["lta_ann"] / 12), res["lta_ann"], "Flexi-Basket"),
            ("Telecom & Internet Allowance", round(res["telecom_ann"] / 12), res["telecom_ann"], "Flexi-Basket"),
            ("Food Coupons & Books", round((res["meal_ann"] + res["books_ann"]) / 12), res["meal_ann"] + res["books_ann"], "Flexi-Basket"),
            ("Special Allowance (Balancing Plug)", res["special_mo"], res["special_ann"], "Earnings"),
            ("GROSS FIXED SALARY (A)", res["gross_mo"], res["gross_ann"], "Subtotal"),
            ("Employer EPF Contribution", res["emp_pf_mo"], res["emp_pf_ann"], "Employer Retiral"),
            ("Employer ESIC (3.25%)", round(res["employer_esi_ann"] / 12), res["employer_esi_ann"], "Employer Retiral"),
            ("Employer Gratuity (4.81%)", round(res["gratuity_ann"] / 12), res["gratuity_ann"], "Employer Retiral"),
            ("Employer Corporate NPS (80CCD(2))", round(res["nps_ann"] / 12), res["nps_ann"], "Employer Retiral"),
            ("FIXED COST TO COMPANY (Fixed CTC)", round(res["fixed_ctc"] / 12), res["fixed_ctc"], "Subtotal"),
            ("Variable Performance Pay", round(res["variable_pay"] / 12), res["variable_pay"], "Variable"),
            ("TOTAL COST TO COMPANY (TOTAL CTC)", round(res["total_ctc"] / 12), res["total_ctc"], "Total"),
            ("Employee EPF + ESIC + Prof. Tax", round((res["emp_pf_ann"] + res["employee_esi_ann"] + res["pt_ann"]) / 12), res["emp_pf_ann"] + res["employee_esi_ann"] + res["pt_ann"], "Deduction"),
            ("Income Tax (Monthly TDS)", res["active_tax_mo"], res["active_tax_ann"], "Deduction"),
            ("NET MONTHLY TAKE-HOME (FIXED)", res["net_fixed_mo"], res["net_fixed_ann"], "Net In-Hand"),
        ], columns=["Component", "Monthly (₹)", "Annual (₹)", "Category"])

        st.dataframe(
            struct_df.style.format({"Monthly (₹)": "₹{:,.0f}", "Annual (₹)": "₹{:,.0f}"}),
            use_container_width=True,
            hide_index=True,
            height=630,
        )

# --- TAB 2: DUAL-REGIME TAX PLANNER ---
with tab2:
    t_col1, t_col2 = st.columns([1, 2])
    with t_col1:
        st.subheader("Old Regime Tax Declarations")
        st.number_input("Actual Monthly Rent Paid (₹/mo)", 0, 500000, step=1000, key="monthly_rent")
        st.caption(f"Auto-Included Employee EPF in 80C: **₹{res['emp_pf_ann']:,.0f}**")
        st.number_input("Other Sec 80C (PPF, ELSS, LIC)", 0, 150000, step=5000, key="sec_80c_other")
        st.number_input("Sec 80CCD(1B) Self NPS (Max ₹50k)", 0, 50000, step=5000, key="sec_80ccd1b")
        st.number_input("Sec 80D Medical Insurance", 0, 100000, step=5000, key="sec_80d")
        st.number_input("Sec 24(b) Home Loan Interest", 0, 200000, step=10000, key="sec_24b")

        st.markdown("#### Sec 10(13A) HRA Exemption Working")
        r1, r2, r3 = res["hra_rules"]
        st.write(f"1. Actual HRA Received: **₹{r1:,.0f}**")
        st.write(f"2. Rent Paid − 10% Basic: **₹{r2:,.0f}**")
        st.write(f"3. 50%/40% of Basic Salary: **₹{r3:,.0f}**")
        st.success(f"**Exempt HRA Allowed: ₹{res['hra_exempt']:,.0f}**")

    with t_col2:
        st.subheader("Side-by-Side Tax Computation (FY 2026–27)")
        gross_taxable = res["gross_ann"] + res["variable_pay"] + res["nps_ann"]
        tax_df = pd.DataFrame([
            ("Gross Taxable Salary (Incl. Variable & NPS)", gross_taxable, gross_taxable),
            ("Less: Standard Deduction u/s 16(ia)", 75000, 50000),
            ("Less: HRA Exemption u/s 10(13A)", 0, res["hra_exempt"]),
            ("Less: Flexi-Benefit Basket (LTA, Telecom, Meal, Books)", 0, res["total_fbp_ann"]),
            ("Less: Professional Tax u/s 16(iii)", 0, res["pt_ann"]),
            ("Less: Employer Corporate NPS u/s 80CCD(2)", min(res["nps_ann"], round(res["basic_ann"] * 0.14)), min(res["nps_ann"], round(res["basic_ann"] * 0.10))),
            ("Less: Sec 80C (EPF + Other, Capped at ₹1.5L)", 0, res["sec_80c_eff"]),
            ("Less: Sec 80CCD(1B) + Sec 80D + Sec 24(b)", 0, min(50000, st.session_state.sec_80ccd1b) + min(100000, st.session_state.sec_80d) + min(200000, st.session_state.sec_24b)),
            ("Total Exemptions & Deductions", res["ded_new"], res["ded_old"]),
            ("NET TAXABLE INCOME", res["new_tax"]["taxable_income"], res["old_tax"]["taxable_income"]),
            ("Gross Slab Tax Before Rebate", res["new_tax"]["raw_tax"], res["old_tax"]["raw_tax"]),
            ("Less: Sec 87A Tax Rebate (Up to ₹12L New / ₹5L Old)", res["new_tax"]["rebate_87a"], res["old_tax"]["rebate_87a"]),
            ("Less: Marginal Relief u/s 87A", res["new_tax"]["marginal_relief"], 0),
            ("Add: Health & Education Cess @ 4%", res["new_tax"]["cess"], res["old_tax"]["cess"]),
            ("TOTAL ANNUAL INCOME TAX PAYABLE", res["new_tax"]["total_tax"], res["old_tax"]["total_tax"]),
            ("MONTHLY TDS DEDUCTION", round(res["new_tax"]["total_tax"] / 12), round(res["old_tax"]["total_tax"] / 12)),
        ], columns=["Tax Computation Line Item", "New Tax Regime (₹)", "Old Tax Regime (₹)"])

        st.dataframe(
            tax_df.style.format({"New Tax Regime (₹)": "₹{:,.0f}", "Old Tax Regime (₹)": "₹{:,.0f}"}),
            use_container_width=True,
            hide_index=True,
            height=580,
        )

# --- TAB 3: APPRAISAL & HIKE MODELER ---
with tab3:
    h_col1, h_col2 = st.columns([1, 2])
    with h_col1:
        st.subheader("Appraisal Inputs")
        st.selectbox("Performance Rating", ["5 - Outstanding", "4 - Exceeds Expectations", "3 - Meets Expectations"], key="perf_rating")
        st.slider("Proposed CTC Hike (%)", 0.0, 100.0, step=0.5, key="hike_pct")
        st.slider("Revised Variable Pay (% of New CTC)", 0.0, 40.0, step=0.5, key="rev_var_pct")
        st.slider("Backdated Arrears Months", 0, 12, step=1, key="arrear_months")

    with h_col2:
        rev_ctc = round(res["total_ctc"] * (1.0 + st.session_state.hike_pct / 100.0))
        rev_kwargs = dict(common_kwargs)
        rev_kwargs["var_pct"] = st.session_state.rev_var_pct
        rev_res = calculate_salary_structure(annual_ctc=rev_ctc, **rev_kwargs)

        real_net_hike_pct = ((rev_res["net_fixed_mo"] - res["net_fixed_mo"]) / res["net_fixed_mo"] * 100.0) if res["net_fixed_mo"] > 0 else 0.0
        arrear_gross = (rev_res["gross_mo"] - res["gross_mo"]) * st.session_state.arrear_months

        hk1, hk2, hk3 = st.columns(3)
        hk1.metric("Revised Annual CTC", f"₹{rev_ctc:,.0f}", f"+₹{rev_ctc - res['total_ctc']:,.0f} ({st.session_state.hike_pct:.1f}%)")
        hk2.metric("Revised Monthly In-Hand", f"₹{rev_res['net_fixed_mo']:,.0f}", f"+₹{rev_res['net_fixed_mo'] - res['net_fixed_mo']:,.0f}/mo ({real_net_hike_pct:.1f}% Real)")
        hk3.metric(f"Backdated Arrears ({st.session_state.arrear_months}m)", f"₹{arrear_gross:,.0f}", "One-Time Gross Payout")

        hike_df = pd.DataFrame([
            ("Total Annual CTC", res["total_ctc"], rev_ctc, rev_ctc - res["total_ctc"]),
            ("Fixed Annual CTC", res["fixed_ctc"], rev_res["fixed_ctc"], rev_res["fixed_ctc"] - res["fixed_ctc"]),
            ("Monthly Gross Salary", res["gross_mo"], rev_res["gross_mo"], rev_res["gross_mo"] - res["gross_mo"]),
            ("Monthly Income Tax (TDS)", res["active_tax_mo"], rev_res["active_tax_mo"], rev_res["active_tax_mo"] - res["active_tax_mo"]),
            ("Monthly Net Take-Home (Fixed)", res["net_fixed_mo"], rev_res["net_fixed_mo"], rev_res["net_fixed_mo"] - res["net_fixed_mo"]),
        ], columns=["Metric", "Current (₹)", "Revised (₹)", "Net Delta (+₹)"])

        st.dataframe(
            hike_df.style.format({"Current (₹)": "₹{:,.0f}", "Revised (₹)": "₹{:,.0f}", "Net Delta (+₹)": "+₹{:,.0f}"}),
            use_container_width=True,
            hide_index=True,
        )

# --- TAB 4: OFFER LETTER STUDIO & WORD/EXCEL EXPORT ---
with tab4:
    st.subheader("📄 Offer Letter Studio & Professional Document Generation")
    st.caption("Customize candidate details, probationary terms, and corporate signatories below, then download the editable Microsoft Word offer letter or multi-tab Excel workbook.")

    o_col1, o_col2 = st.columns(2)
    with o_col1:
        st.text_input("Company Header Name", key="company_name")
        st.text_input("Registered Address", key="company_address")
        st.text_input("Brand Unit / Subsidiary", key="brand_unit")
        st.text_input("Candidate Full Name", key="emp_name")
        st.text_input("Designation / Title", key="designation")
    with o_col2:
        st.text_input("Job Grade / Band", key="grade")
        st.text_input("Joining Date", key="joining_date")
        st.number_input("Probation Period (Months)", 1, 12, step=1, key="probation_months")
        st.text_input("Authorized Signatory Name", key="signatory_name")
        st.text_input("Signatory Designation", key="signatory_title")

    st.divider()

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        word_bytes = build_word_offer_letter_bytes(st.session_state, res)
        st.download_button(
            label="📝 Download Editable Offer Letter (.docx)",
            data=word_bytes,
            file_name=f"Offer_Letter_{st.session_state.emp_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    with dl_col2:
        excel_bytes = build_excel_workbook_bytes(st.session_state, res)
        st.download_button(
            label="📊 Download Multi-Tab Excel Tax & CTC Workbook (.xlsx)",
            data=excel_bytes,
            file_name=f"Salary_Structure_{st.session_state.emp_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
