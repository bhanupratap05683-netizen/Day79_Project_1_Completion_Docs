# 📈 Automated Financial Dashboard - Python + Excel

An end-to-end automated financial dashboard that fetches live market data, analyzes it with pandas, and generates a professionally formatted Excel report with charts.

This project was built as Project 1 for my 84-Day Python & Excel Mastery.

### 🚀 Demo
![Dashboard Screenshot](charts/sample_dashboard.png)
*Auto-generated Excel report with KPIs, formatted tables, and embedded charts.*

### ✨ Key Features
- **API Data Pipeline:** Fetches stock data (with fallback to sample data if API key not present)
- **Pandas Analysis:** Calculates daily returns, moving averages, volatility
- **Excel Automation:** Generates multi-sheet report with openpyxl styling (colors, filters, currency format)
- **Live Excel Integration:** Optional xlwings integration to update live open workbook
- **Visualization:** Matplotlib charts saved and embedded into Excel
- **Production Ready:** Logging, error handling, PEP 8 compliant

### 🛠️ Tech Stack
Python, pandas, openpyxl, matplotlib, requests, xlwings, logging

### 📁 Folder Structure
