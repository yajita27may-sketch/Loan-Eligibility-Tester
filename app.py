import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Loan Eligibility System", layout="wide")

FILE_NAME = "loan_reports.xlsx"

# -----------------------------
# CUSTOM CSS
# -----------------------------
st.markdown("""
    <style>
    body {
        background: linear-gradient(135deg, #1f4037, #99f2c8);
    }
    .card {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
        margin-bottom: 15px;
    }
    .success { color: green; font-weight: bold; font-size: 18px; }
    .error { color: red; font-weight: bold; font-size: 18px; }
    </style>
""", unsafe_allow_html=True)

st.title("🏦 Loan Eligibility & Bank Comparison System")

# -----------------------------
# FUNCTIONS
# -----------------------------

def calculate_credit_score(income, existing_emi, employment):
    score = 650

    if income > 150000:
        score += 120
    elif income > 75000:
        score += 80
    elif income > 40000:
        score += 40

    if existing_emi > income * 0.4:
        score -= 80

    if employment in ["Self-employed", "Business Owner"]:
        score -= 20

    return max(300, min(score, 900))


def calculate_emi(P, annual_rate, n):
    r = annual_rate / (12 * 100)
    return P * r * (1 + r)**n / ((1 + r)**n - 1)


def calculate_total_interest(emi, principal, n):
    total = emi * n
    return total, total - principal


def check_eligibility(age, income, existing_emi):
    if age < 21:
        return False, 0, 0

    max_emi = income * 0.45
    available_emi = max_emi - existing_emi

    if available_emi <= 0:
        return False, 0, 0

    return True, max_emi, available_emi


def estimate_loan_amount(emi, rate, n):
    r = rate / (12 * 100)
    return emi * ((1 + r)**n - 1) / (r * (1 + r)**n)


def get_required_documents(loan_type, employment):
    docs = ["Aadhaar Card", "PAN Card", "Photo"]

    if employment == "Salaried":
        docs += ["Salary Slips", "Bank Statement", "Form 16"]
    else:
        docs += ["ITR", "Business Proof", "Bank Statement"]

    if loan_type == "Home Loan":
        docs += ["Property Papers"]
    elif loan_type in ["CC (Cash Credit)", "OD (Overdraft)"]:
        docs += ["Business Financials"]
    elif loan_type == "TL (Term Loan)":
        docs += ["Project Report"]

    return docs


def compare_banks(available_emi, tenure_months):
    banks = {
        "SBI": 8.5,
        "HDFC": 9.0,
        "ICICI": 9.2,
        "Axis": 9.5
    }

    data = []

    for bank, rate in banks.items():
        loan = estimate_loan_amount(available_emi, rate, tenure_months)
        emi = calculate_emi(loan, rate, tenure_months)
        total, interest = calculate_total_interest(emi, loan, tenure_months)

        data.append({
            "Bank": bank,
            "Interest Rate (%)": rate,
            "Loan Amount (₹)": round(loan, 2),
            "EMI (₹)": round(emi, 2),
            "Total Interest (₹)": round(interest, 2),
            "Total Payment (₹)": round(total, 2)
        })

    return pd.DataFrame(data)

# -----------------------------
# SAVE TO EXCEL (NEW FEATURE)
# -----------------------------

def save_to_excel(record):
    df_new = pd.DataFrame([record])

    if os.path.exists(FILE_NAME):
        df_old = pd.read_excel(FILE_NAME)
        df_final = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_final = df_new

    df_final.to_excel(FILE_NAME, index=False)


def download_excel():
    if os.path.exists(FILE_NAME):
        with open(FILE_NAME, "rb") as file:
            st.download_button(
                label="📥 Download Full Excel Report",
                data=file,
                file_name="Loan_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# -----------------------------
# INPUT
# -----------------------------

st.sidebar.header("📋 Enter Details")

name = st.sidebar.text_input("Name")
age = st.sidebar.number_input("Age", 18, 70, 25)
income = st.sidebar.number_input("Monthly Income", value=50000)

employment = st.sidebar.selectbox(
    "Employment Type",
    ["Salaried", "Self-employed", "Business Owner", "Professional", "Freelancer"]
)

loan_type = st.sidebar.selectbox(
    "Loan Type",
    ["Personal Loan", "Home Loan", "Car Loan", "CC (Cash Credit)", "OD (Overdraft)", "TL (Term Loan)"]
)

existing_emi = st.sidebar.number_input("Existing EMI", value=0)
tenure = st.sidebar.slider("Tenure (Years)", 1, 30, 5)

months = tenure * 12

# -----------------------------
# MAIN
# -----------------------------

if st.sidebar.button("🚀 Check Eligibility"):

    score = calculate_credit_score(income, existing_emi, employment)
    eligible, max_emi, avail_emi = check_eligibility(age, income, existing_emi)

    st.markdown(f"<div class='card'>📊 Credit Score: <b>{score}</b></div>", unsafe_allow_html=True)

    if not eligible:
        st.markdown("<div class='error'>❌ Not Eligible</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='success'>✅ Eligible</div>", unsafe_allow_html=True)

        df = compare_banks(avail_emi, months)

        st.subheader("🏦 Bank Comparison")
        st.dataframe(df)

        best = df.loc[df["EMI (₹)"].idxmin()]

        st.markdown(f"""
        <div class='card'>
        🏆 <b>Best Bank:</b> {best['Bank']} <br>
        📊 <b>Interest Rate:</b> {best['Interest Rate (%)']}% <br>
        💰 <b>Loan Amount:</b> ₹{best['Loan Amount (₹)']} <br>
        📉 <b>EMI:</b> ₹{best['EMI (₹)']}
        </div>
        """, unsafe_allow_html=True)

        st.subheader("📄 Required Documents")
        docs = get_required_documents(loan_type, employment)
        for d in docs:
            st.write("✔", d)

        st.subheader("📊 EMI Chart")
        st.bar_chart(df.set_index("Bank")["EMI (₹)"])

        # -----------------------------
        # SAVE DATA TO EXCEL (NEW)
        # -----------------------------
        record = {
            "Name": name,
            "Age": age,
            "Income": income,
            "Employment": employment,
            "Loan Type": loan_type,
            "Credit Score": score,
            "Eligible": "Yes" if eligible else "No",
            "Max EMI": max_emi,
            "Available EMI": avail_emi,
            "Best Bank": best["Bank"],
            "Interest Rate": best["Interest Rate (%)"],
            "Loan Amount": best["Loan Amount (₹)"],
            "EMI": best["EMI (₹)"]
        }

        save_to_excel(record)

# -----------------------------
# DOWNLOAD SECTION
# -----------------------------

st.subheader("📊 Reports")
download_excel()