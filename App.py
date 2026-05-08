import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import pdfplumber
import json
import os
from openpyxl import Workbook
from openai import OpenAI

# ====== CONFIG ======
st.set_page_config(page_title="Finance Agent", layout="wide")

# ====== FIND OR CREATE EXCEL ======
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

    wb.save("finance_automation_template.xlsx")
    EXCEL_FILE = "finance_automation_template.xlsx"

# ====== CATEGORY ENGINE ======
def categorize(desc):
    if "כרטיס" in desc:
        return "צריכה"
    if "מקס" in desc:
        return "אשראי"
    if "משכנת" in desc:
        return "דיור"
    if "פנסיה" in desc or "גמל" in desc:
        return "חיסכון"
    if "משכורת" in desc:
        return "הכנסה"
    return "אחר"

# ====== LOAD DATA ======
def load_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Transactions", engine="openpyxl")
    balance = pd.read_excel(EXCEL_FILE, sheet_name="Balance", engine="openpyxl")
    df["Date"] = pd.to_datetime(df["Date"])
    return df, balance

# ====== SAVE DATA ======
def save_data(df, balance):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Transactions", index=False)
        balance.to_excel(writer, sheet_name="Balance", index=False)

# ====== PDF → AI ======
def parse_pdf_with_ai(uploaded_file):

    text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                text += txt + "\n"

    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

    prompt = f"""
    Extract all bank transactions from this text:

    {text}

    Return ONLY JSON:

    {{
      "transactions": [
        {{
          "date": "YYYY-MM-DD",
          "description": "text",
          "amount": number,
          "type": "income or expense"
        }}
      ],
      "balance": {{
        "date": "YYYY-MM-DD",
        "amount": number
      }}
    }}

    Rules:
    - זכות = income
    - חובה = expense
    - numbers must be positive
    - clean Hebrew text
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}]
    )

    try:
        result = response.choices[0].message.content
        return json.loads(result)
    except:
        st.error("❌ שגיאה בניתוח ה‑PDF")
        return None

# ====== UI ======
st.title("📊 Finance Agent Dashboard")

uploaded_file = st.file_uploader("📄 העלה דוח בנק (PDF)", type="pdf")

if uploaded_file and st.button("🚀 Process PDF"):
    df, balance_df = load_data()

    data = parse_pdf_with_ai(uploaded_file)

    if data:

        transactions = data["transactions"]
        balance = data["balance"]

        # עדכון נתונים
        for t in transactions:
            df.loc[len(df)] = [
                t["date"],
                t["description"],
                categorize(t["description"]),
                t["type"],
                t["amount"]
            ]

        balance_df.loc[len(balance_df)] = [
            balance["date"],
            balance["amount"]
        ]

        save_data(df, balance_df)

        st.success("✅ הנתונים עודכנו בהצלחה")

# ====== DASHBOARD ======
df, balance_df = load_data()

income = df[df["Type"] == "income"]["Amount"].sum()
expense = df[df["Type"] == "expense"]["Amount"].sum()
savings = income - expense

# ====== METRICS ======
col1, col2, col3 = st.columns(3)
col1.metric("💰 Income", f"₪{income:,.0f}")
col2.metric("💸 Expenses", f"₪{expense:,.0f}")
col3.metric("📊 Savings", f"₪{savings:,.0f}")

# ====== GRAPH 1 ======
st.subheader("📊 הוצאות לפי קטגוריה")
expenses_by_cat = df[df["Type"] == "expense"].groupby("Category")["Amount"].sum()
st.bar_chart(expenses_by_cat)

# ====== GRAPH 2 ======
st.subheader("⚖️ הכנסות מול הוצאות")

labels = ["Income", "Expenses"]
values = [income, expense]

fig, ax = plt.subplots()
colors = ["green", "red"]
bars = ax.bar(labels, values, color=colors)

for bar in bars:
    yval = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, yval, f"{int(yval):,}", ha='center', va='bottom')

st.pyplot(fig)

# ====== GRAPH 3 ======
st.subheader("📅 מגמה חודשית")

df["Month"] = df["Date"].dt.to_period("M").astype(str)
monthly = df.groupby(["Month", "Type"])["Amount"].sum().unstack().fillna(0)

st.line_chart(monthly)

# ====== GRAPH 4 ======
st.subheader("💰 חיסכון חודשי")

monthly["Savings"] = monthly.get("income", 0) - monthly.get("expense", 0)
st.line_chart(monthly["Savings"])

# ====== INSIGHTS ======
st.subheader("🧠 Insights")

if expense > 30000:
    st.warning("⚠️ הוצאות גבוהות")

if savings < 0:
    st.error("❌ גירעון")

if savings > 20000:
    st.success("✅ חיסכון חזק")

