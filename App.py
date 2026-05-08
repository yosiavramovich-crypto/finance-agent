import pdfplumber
import openai
import json
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

EXCEL_FILE = "finance_automation_template.xlsx"

def load_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Transactions", engine="openpyxl")
    df['Date'] = pd.to_datetime(df['Date'])
    return df

st.title("📊 Finance Dashboard")

df = load_data()

income = df[df["Type"] == "income"]["Amount"].sum()
expense = df[df["Type"] == "expense"]["Amount"].sum()
savings = income - expense

col1, col2, col3 = st.columns(3)
col1.metric("Income", f"₪{income:,.0f}")
col2.metric("Expenses", f"₪{expense:,.0f}")
col3.metric("Savings", f"₪{savings:,.0f}")

st.subheader("📊 הוצאות לפי קטגוריה")
expense_by_cat = df[df["Type"] == "expense"].groupby("Category")["Amount"].sum()
st.bar_chart(expense_by_cat)

st.subheader("⚖️ הכנסות מול הוצאות")
labels = ['Income', 'Expenses']
values = [income, expense]
fig, ax = plt.subplots()
ax.bar(labels, values, color=['green', 'red'])
st.pyplot(fig)

st.subheader("📅 מגמה חודשית")
df["Month"] = df["Date"].dt.to_period("M").astype(str)
monthly = df.groupby(["Month", "Type"])["Amount"].sum().unstack().fillna(0)
st.line_chart(monthly)

st.subheader("💰 חיסכון חודשי")
monthly["Savings"] = monthly.get("income", 0) - monthly.get("expense", 0)
st.line_chart(monthly["Savings"])
def parse_pdf_with_ai(uploaded_file):
uploaded_file = st.file_uploader("Upload PDF", type="pdf")

if uploaded_file and st.button("Process"):

    data = parse_pdf_with_ai(uploaded_file)

    transactions = data["transactions"]
    balance = data["balance"]

    text = ""

    # קריאת PDF
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    # קריאה ל‑AI
    client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

    prompt = f"""
    Extract all bank transactions from this text:

    {text}

    Return JSON only:

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
    זכות = income
    חובה = expense
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}]
    )

    result = response.choices[0].message.content

    return json.loads(result)
