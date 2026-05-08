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
