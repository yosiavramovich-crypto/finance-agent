import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import pdfplumber
import re
import os
from openpyxl import Workbook

# ===== CONFIG =====
st.set_page_config(page_title="Finance Agent", layout="wide")

# ===== EXCEL HANDLING =====
EXCEL_FILE = None

for file in os.listdir():
    if file.endswith(".xlsx"):
        EXCEL_FILE = file
        break

if EXCEL_FILE is None:
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Transactions"
    ws1.append(["Date", "Description", "Category", "Type", "Amount"])

    ws2 = wb.create_sheet(title="Balance")
    ws2.append(["Date", "Balance"])

    wb.save("finance.xlsx")
    EXCEL_FILE = "finance.xlsx"

# ===== CATEGORY ENGINE =====
def categorize(desc):
    desc = desc.lower()

    if "כרטיס" in desc:
        return "צריכה"
    if "מקס" in desc:
        return "אשראי"
    if "משכנת" in desc:
        return "דיור"
    if "פנסיה" in desc:
        return "חיסכון"
    if "גמל" in desc:
        return "השקעה"
    if "משכורת" in desc:
        return "הכנסה"
    if "קצבה" in desc:
        return "הכנסה"
    if "ביטוח" in desc:
        return "הכנסה"

    return "אחר"

# ===== LOAD DATA =====
def load_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Transactions", engine="openpyxl")
    balance = pd.read_excel(EXCEL_FILE, sheet_name="Balance", engine="openpyxl")

    if not df.empty:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    return df, balance

# ===== SAVE =====
def save_data(df, balance):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Transactions", index=False)
        balance.to_excel(writer, sheet_name="Balance", index=False)

# ===== 🔥 PARSER מותאם לדוח שלך =====
def parse_pdf_local(uploaded_file):

    text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"

    if len(text.strip()) == 0:
        st.error("❌ לא הצלחנו לקרוא את ה-PDF")
        return None

    lines = text.split("\n")

    transactions = []
    balance = {"date": None, "amount": 0}

    for line in lines:

        # ===== תאריך =====
        date_match = re.search(r"(\d{2}\.\d{2}\.\d{4})", line)
        if not date_match:
            continue

        date = pd.to_datetime(date_match.group(1), dayfirst=True)

        # ===== סכומים =====
        amounts = re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}", line)
        if len(amounts) < 1:
            continue

        try:
            if len(amounts) >= 2:
                amount = float(amounts[-2].replace(",", ""))
                balance_amount = float(amounts[-1].replace(",", ""))
            else:
                amount = float(amounts[0].replace(",", ""))
                balance_amount = 0
        except:
            continue

        # ===== סוג =====
        if any(word in line for word in ["משכורת", "קצבה", "ביטוח", "תגמול", "זכות"]):
            t_type = "income"
        else:
            t_type = "expense"

        # ===== ניקוי תיאור =====
        desc = line
        desc = re.sub(r"\d{2}\.\d{2}\.\d{4}", "", desc)
        desc = re.sub(r"\d{1,3}(?:,\d{3})*\.\d{2}", "", desc)
        desc = desc.replace("₪", "")
        desc = desc.strip()

        if len(desc) < 3:
            continue

        transactions.append({
            "date": str(date.date()),
            "description": desc,
            "amount": amount,
            "type": t_type
        })

        # עדכון יתרה (האחרונה תישמר)
        balance = {
            "date": str(date.date()),
            "amount": balance_amount
        }

    return {
        "transactions": transactions,
        "balance": balance
    }

# ===== UI =====
st.title("📊 Finance Agent Dashboard")

uploaded_file = st.file_uploader("📄 העלה דוח בנק (PDF)", type="pdf")

if uploaded_file and st.button("🚀 עבד דוח"):

    df, balance_df = load_data()
    data = parse_pdf_local(uploaded_file)

    if data is None:
        st.stop()

    for t in data["transactions"]:
        df.loc[len(df)] = [
            t["date"],
            t["description"],
            categorize(t["description"]),
            t["type"],
            t["amount"]
        ]

    balance_df.loc[len(balance_df)] = [
        data["balance"]["date"],
        data["balance"]["amount"]
    ]

    save_data(df, balance_df)

    st.success("✅ הנתונים עודכנו בהצלחה")

# ===== DASHBOARD =====
df, balance_df = load_data()

income = df[df["Type"] == "income"]["Amount"].sum()
expense = df[df["Type"] == "expense"]["Amount"].sum()
savings = income - expense

col1, col2, col3 = st.columns(3)
col1.metric("💰 Income", f"₪{income:,.0f}")
col2.metric("💸 Expenses", f"₪{expense:,.0f}")
col3.metric("📊 Savings", f"₪{savings:,.0f}")

# ===== גרף קטגוריות =====
st.subheader("📊 הוצאות לפי קטגוריה")
expense_by_cat = df[df["Type"] == "expense"].groupby("Category")["Amount"].sum()
st.bar_chart(expense_by_cat)

# ===== גרף הכנסות מול הוצאות =====
st.subheader("⚖️ הכנסות מול הוצאות")

fig, ax = plt.subplots()
ax.bar(["Income", "Expenses"], [income, expense], color=["green", "red"])
st.pyplot(fig)

# ===== מגמה חודשית =====
st.subheader("📅 מגמה חודשית")

if not df.empty:
    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    monthly = df.groupby(["Month", "Type"])["Amount"].sum().unstack().fillna(0)
    st.line_chart(monthly)

# ===== חיסכון =====
st.subheader("💰 חיסכון חודשי")

if not df.empty:
    monthly["Savings"] = monthly.get("income", 0) - monthly.get("expense", 0)
    st.line_chart(monthly["Savings"])

# ===== תובנות =====
st.subheader("🧠 Insights")

if expense > 30000:
    st.warning("⚠️ הוצאות גבוהות")

if savings < 0:
    st.error("❌ גירעון")

if savings > 20000:
    st.success("✅ חיסכון גבוה")
