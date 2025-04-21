import pandas as pd
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import requests
import yfinance as yf
import numpy as np
from fpdf import FPDF
import smtplib
from email.message import EmailMessage

class FinanceManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("\U0001F4C8 Stock Analyzer - Finance Manager")
        self.root.geometry("1280x850")
        self.root.configure(bg="white")

        self.create_widgets()
        self.stock_data = []
        self.create_table()
        self.create_tabs()

    def create_widgets(self):
        title_frame = tb.Frame(self.root, bootstyle="light")
        title_frame.pack(fill="x", pady=(10, 0))

        title = tb.Label(title_frame, text="\U0001F4CA Smart Finance Manager", font=("Segoe UI Emoji", 22, "bold"), bootstyle="primary")
        title.pack(side="left", padx=20, pady=5)

        control_frame = tb.Frame(self.root)
        control_frame.pack(pady=10, fill="x")

        left_controls = tb.Frame(control_frame)
        left_controls.pack(side="left")

        tb.Button(left_controls, text="\U0001F4E5 Load Excel", command=self.load_excel, bootstyle="info-outline").grid(row=0, column=0, padx=8)
        tb.Button(left_controls, text="\U0001F501 Refresh View", command=self.refresh_view, bootstyle="success-outline").grid(row=0, column=1, padx=8)
        tb.Button(left_controls, text="\U0001F4F0 Get Live Price", command=self.get_live_price, bootstyle="warning-outline").grid(row=0, column=2, padx=8)
        tb.Button(left_controls, text="\U0001F4CE Export PDF", command=self.export_pdf, bootstyle="secondary-outline").grid(row=0, column=3, padx=8)
        tb.Button(left_controls, text="\U0001F4E7 Email Report", command=self.send_email_report, bootstyle="light-outline").grid(row=0, column=4, padx=8)
        tb.Button(left_controls, text="❌ Exit", command=self.root.quit, bootstyle="danger-outline").grid(row=0, column=5, padx=8)

        right_frame = tb.Frame(control_frame)
        right_frame.pack(side="right", padx=(10, 20))

        tb.Label(right_frame, text="Currency:").grid(row=0, column=0, padx=(0, 5))
        self.currency_var = tk.StringVar(value="USD")
        self.currency_dropdown = tb.Combobox(right_frame, textvariable=self.currency_var, values=["USD", "AED", "INR"], width=10, state="readonly")
        self.currency_dropdown.grid(row=0, column=1)
        self.currency_dropdown.bind("<<ComboboxSelected>>", self.on_currency_change)

        tb.Label(right_frame, text="Data Source:").grid(row=0, column=2, padx=(20, 5))
        self.source_var = tk.StringVar(value="Yahoo")
        self.source_dropdown = tb.Combobox(right_frame, textvariable=self.source_var, values=["Yahoo", "Capital.com"], width=12, state="readonly")
        self.source_dropdown.grid(row=0, column=3)
        self.source_dropdown.bind("<<ComboboxSelected>>", self.on_source_change)

        tb.Label(right_frame, text="Select Stock:").grid(row=0, column=4, padx=(20, 5))
        self.symbol_var = tk.StringVar()
        self.symbol_dropdown = tb.Combobox(right_frame, textvariable=self.symbol_var, values=[], width=10, state="readonly")
        self.symbol_dropdown.grid(row=0, column=5)
        self.symbol_dropdown.bind("<<ComboboxSelected>>", self.on_symbol_select)

        self.status_box = scrolledtext.ScrolledText(self.root, height=6, font=("Segoe UI Emoji", 10))
        self.status_box.pack(fill="x", padx=20, pady=(5, 10))
        self.status_box.insert("end", "Welcome to your Smart Finance Manager Dashboard!\n")

    def create_tabs(self):
        self.tab_control = ttk.Notebook(self.root)
        self.tab_dashboard = tb.Frame(self.tab_control)
        self.tab_control.add(self.tab_dashboard, text='Dashboard')
        self.tab_control.pack(expand=1, fill='both', padx=20, pady=10)
        self.create_dashboard_table()

    def create_table(self):
        pass

    def create_dashboard_table(self):
        table_frame = tb.Labelframe(self.tab_dashboard, text="\U0001F4CA Portfolio Overview", bootstyle="info")
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)
        columns = ("Stock Symbol", "Advice", "Profit", "RSI", "Live Advice", "Alert", "Purchased", "Data Source")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=130)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.show_stock_details)

    def on_currency_change(self, event):
        self.status_box.insert("end", f"\U0001F4B1 Currency switched to: {self.currency_var.get()}\n")

    def on_source_change(self, event):
        self.status_box.insert("end", f"\U0001F310 Data source changed to: {self.source_var.get()}\n")

    def on_symbol_select(self, event):
        selected = self.symbol_var.get()
        if selected:
            self.status_box.insert("end", f"\U0001F4CC Selected stock: {selected}\n")
            self.show_stock_chart(selected)

    def load_excel(self):
        self.status_box.insert("end", "\U0001F4C2 Loading Stock_Analyzer_Multi_Output.xlsx...\n")
        self.display_data("Stock_Analyzer_Multi_Output.xlsx")

    def refresh_view(self):
        self.status_box.insert("end", f"\U0001F501 Refreshing with currency: {self.currency_var.get()} and source: {self.source_var.get()}\n")
        self.display_data("Stock_Analyzer_Multi_Output.xlsx")

    def display_data(self, filepath):
        try:
            df = pd.read_excel(filepath, sheet_name="Dashboard")
            df = df.dropna(subset=["Stock Symbol"])
            df = df[~df["Stock Symbol"].str.lower().isin(["buy", "sell", "hold", "advice", "count", "nan", ""])]

            self.tree.delete(*self.tree.get_children())
            self.symbol_dropdown["values"] = df["Stock Symbol"].tolist()

            for _, row in df.iterrows():
                symbol = row.get("Stock Symbol", "")
                advice = row.get("Advice", "–")
                profit = row.get("Profit", 0)
                rsi = row.get("RSI", "–")
                alert = row.get("Alert", "–")
                source = row.get("Data Source", self.source_var.get())
                purchased = "Yes" if advice and advice != "–" else "No"

                if pd.isna(source) or source == "N/A":
                    source = self.source_var.get()

                if pd.isna(rsi):
                    live_advice = "–"
                elif rsi < 30 and profit > 0:
                    live_advice = "\U0001F525 Strong Buy"
                elif rsi > 70 and profit > 0:
                    live_advice = "\U0001F53A Consider Selling"
                elif abs(rsi - 50) <= 10:
                    live_advice = "⚖️ Wait"
                else:
                    live_advice = "\U0001F4CA Review"

                self.tree.insert("", "end", values=(
                    symbol,
                    advice,
                    f"{profit:.2f}",
                    round(rsi, 2) if not pd.isna(rsi) else "–",
                    live_advice,
                    alert,
                    purchased,
                    source
                ))

            self.status_box.insert("end", "✅ Excel Loaded and Displayed.\n")

        except Exception as e:
            self.status_box.insert("end", f"❌ Error loading file: {e}\n")

    def get_live_price(self):
        symbol = self.symbol_var.get()
        if not symbol:
            self.status_box.insert("end", "⚠️ Please select a stock symbol first.\n")
            return

        source = self.source_var.get()
        if source == "Capital.com":
            self.status_box.insert("end", f"\U0001F310 Capital.com - Fetching live price for {symbol}...\n")
        else:
            self.status_box.insert("end", f"\U0001F310 Yahoo - Fetching live price for {symbol}...\n")

        try:
            ticker = yf.Ticker(symbol)
            hist = ticker.history(period="1d")
            price = hist["Close"].iloc[-1] if not hist.empty else "N/A"
        except Exception as e:
            price = f"Error: {e}"

        self.status_box.insert("end", f"\U0001F4B5 Live Price for {symbol}: {price}\n")

    def export_pdf(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Portfolio Summary", ln=1, align="C")
        for child in self.tree.get_children():
            values = self.tree.item(child)['values']
            line = ", ".join(str(v) for v in values)
            pdf.multi_cell(0, 10, txt=line)

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if save_path:
            pdf.output(save_path)
            messagebox.showinfo("Exported", f"Dashboard exported to PDF at:\n{save_path}")

    def send_email_report(self):
        try:
            msg = EmailMessage()
            msg['Subject'] = 'Smart Finance Manager - Portfolio Summary'
            msg['From'] = 'tuttu19@gmail.com'
            msg['To'] = 'tuttu19@gmail.com'
            msg.set_content('Please find attached your latest portfolio summary.')

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt="Portfolio Summary", ln=1, align="C")
            for child in self.tree.get_children():
                values = self.tree.item(child)['values']
                line = ", ".join(str(v) for v in values)
                pdf.multi_cell(0, 10, txt=line)
            pdf_path = "portfolio_temp.pdf"
            pdf.output(pdf_path)

            with open(pdf_path, "rb") as f:
                file_data = f.read()
                msg.add_attachment(file_data, maintype="application", subtype="pdf", filename="Portfolio_Summary.pdf")

            with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                smtp.starttls()
                smtp.login('tuttu19@gmail.com', 'fppw nygd ogzn panz')
                smtp.send_message(msg)

            os.remove(pdf_path)
            messagebox.showinfo("Email Sent", "Portfolio summary emailed successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email:\n{e}")

    def show_stock_chart(self, symbol):
        top = tk.Toplevel(self.root)
        top.title(f"📈 {symbol} Price Chart")
        top.geometry("600x400")

        try:
            data = yf.download(symbol, period="1mo", interval="1d")
            fig, ax = plt.subplots(figsize=(6, 4))
            ax.plot(data.index, data['Close'], label='Close Price')
            ax.set_title(f"{symbol} Last 1 Month")
            ax.set_xlabel("Date")
            ax.set_ylabel("Price")
            ax.grid(True)
            canvas = FigureCanvasTkAgg(fig, master=top)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
        except Exception as e:
            tk.Label(top, text=f"Error loading chart: {e}").pack(pady=10)

    def show_stock_details(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            item_data = self.tree.item(selected_item)
            symbol = item_data['values'][0]
            self.symbol_var.set(symbol)
            self.status_box.insert("end", f"\U0001F4CC Selected stock: {symbol}\n")
            self.show_stock_chart(symbol)

if __name__ == "__main__":
    app = tb.Window(themename="flatly")
    FinanceManagerApp(app)
    app.mainloop()