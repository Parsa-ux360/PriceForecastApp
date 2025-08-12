# ==============================================================================
# Module: main.py
# Author: Parsa Shahi
# Date: 2025-08-11
# Description:
#     This module implements the main graphical user interface (GUI) for the Price Forecast App.
#     It uses tkinter with ttk for a modern, responsive layout and styled widgets.
#
#     Main functionalities include:
#       - Adding, editing, and deleting product entries with price and forecast parameters
#       - Saving and loading product data to/from Excel files
#       - Invoking forecast calculations via the Calculate module
#       - Displaying logs and messages within the interface
#       - Persisting app state in a local JSON file for session continuity
#
# Usage:
#     Run this script to launch the GUI application.
#     Interact with the interface to manage products and perform price forecasts.
# ============================================================================== 

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import json
import os
import sys

# Import modules
import CreateExcel, ReadExcel, Calculate

APP_STATE_FILE = "app_state.json"


class PriceForecastApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Price Forecast App")
        self.geometry("1000x700")
        self.configure(bg="#f4f6f8")

        # Use ttk styles for a cleaner look
        style = ttk.Style()
        style.theme_use("clam")  # برای اینکه رنگ پس‌زمینه روی دکمه اعمال بشه

        style.configure('TButton',
                        background='#283593',  # سرمه‌ای
                        foreground='white',
                        padding=8,
                        font=('Segoe UI', 10, 'bold'),
                        relief='flat')

        style.map('TButton',
                background=[('active', '#5c6bc0'), ('pressed', '#1a237e')],
                relief=[('pressed', 'sunken'), ('!pressed', 'flat')])



        self.product_list = []
        self._load_state()

        self._build_ui()

    def _load_state(self):
        if os.path.exists(APP_STATE_FILE):
            try:
                with open(APP_STATE_FILE, 'r', encoding='utf-8') as f:
                    state = json.load(f)
                    self.product_list = state.get('products', [])
            except Exception:
                self.product_list = []

    def _save_state(self):
        try:
            with open(APP_STATE_FILE, 'w', encoding='utf-8') as f:
                json.dump({'products': self.product_list}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _build_ui(self):
        # Top frame with actions
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(side='top', fill='x')

        btn_add = ttk.Button(top_frame, text='➕ Add Product', command=self.add_product_popup)
        btn_save = ttk.Button(top_frame, text='Save to Excel', command=self.save_to_excel_action)
        btn_load = ttk.Button(top_frame, text='Load Excel', command=self.read_excel_action)
        btn_calc = ttk.Button(top_frame, text='🔎 Calculate Forecast', command=self.calculate_forecast_action)

        btn_add.pack(side='left', padx=6)
        btn_save.pack(side='left', padx=6)
        btn_load.pack(side='left', padx=6)
        btn_calc.pack(side='left', padx=6)

        # Main Paned window
        paned = ttk.Panedwindow(self, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=10, pady=10)

        # Left: table
        left_frame = ttk.Frame(paned, width=600)
        paned.add(left_frame, weight=3)

        columns = ("Product", "Current Price", "Forecast Months", "Country", "Currency")
        self.table = ttk.Treeview(left_frame, columns=columns, show='headings')
        for col in columns:
            self.table.heading(col, text=col)
            self.table.column(col, width=150, anchor='center')

        self.table.pack(fill='both', expand=True)

        # populate existing
        for p in self.product_list:
            self.table.insert('', 'end', values=(p.get('Product',''), p.get('Current Price',''), p.get('Forecast Months',''), p.get('Country',''), p.get('Currency','')))

        # Right: log and controls
        right_frame = ttk.Frame(paned, width=300)
        paned.add(right_frame, weight=1)

        ttk.Label(right_frame, text='Log').pack(anchor='w')
        self.log_box = tk.Text(right_frame, height=20, state='disabled', wrap='word')
        self.log_box.pack(fill='both', expand=True)

        # Footer
        footer = ttk.Label(self, text='Price Forecast App — Developed by Parsa Shahi', anchor='center')
        footer.pack(side='bottom', fill='x')

        # Context menu for table
        self.table.bind('<Button-3>', self._on_table_right_click)
        self._build_popup_menu()

    def _build_popup_menu(self):
        self.popup = tk.Menu(self, tearoff=0)
        self.popup.add_command(label='Edit', command=self._edit_selected)
        self.popup.add_command(label='Delete', command=self._delete_selected)

    def _on_table_right_click(self, event):
        iid = self.table.identify_row(event.y)
        if iid:
            self.table.selection_set(iid)
            self.popup.tk_popup(event.x_root, event.y_root)

    def _edit_selected(self):
        sel = self.table.selection()
        if not sel:
            return
        values = self.table.item(sel[0], 'values')
        self._open_edit_popup(values, sel[0])

    def _delete_selected(self):
        sel = self.table.selection()
        if not sel:
            return
        idx = self.table.index(sel[0])
        confirm = messagebox.askyesno('Confirm', 'Delete selected product?')
        if confirm:
            self.table.delete(sel[0])
            del self.product_list[idx]
            self._save_state()
            self.log_message('Product deleted.')

    def add_product_popup(self):
        popup = tk.Toplevel(self)
        popup.title('Add Product')
        popup.geometry('380x320')
        popup.transient(self)

        frm = ttk.Frame(popup, padding=12)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text='Product Name:').grid(row=0, column=0, sticky='w')
        entry_name = ttk.Entry(frm)
        entry_name.grid(row=0, column=1, sticky='ew')

        ttk.Label(frm, text='Current Price:').grid(row=1, column=0, sticky='w')
        entry_price = ttk.Entry(frm)
        entry_price.grid(row=1, column=1, sticky='ew')

        ttk.Label(frm, text='Forecast Months:').grid(row=2, column=0, sticky='w')
        entry_months = ttk.Entry(frm)
        entry_months.grid(row=2, column=1, sticky='ew')

        ttk.Label(frm, text='Country Code (ISO2/ISO3):').grid(row=3, column=0, sticky='w')
        entry_country = ttk.Entry(frm)
        entry_country.grid(row=3, column=1, sticky='ew')

        ttk.Label(frm, text='Currency Code (e.g. USD):').grid(row=4, column=0, sticky='w')
        entry_currency = ttk.Entry(frm)
        entry_currency.grid(row=4, column=1, sticky='ew')

        frm.columnconfigure(1, weight=1)

        def confirm():
            product = {
                'Product': entry_name.get().strip(),
                'Current Price': entry_price.get().strip(),
                'Forecast Months': entry_months.get().strip() or '0',
                'Country': entry_country.get().strip(),
                'Currency': entry_currency.get().strip().upper() or 'USD'
            }
            self.product_list.append(product)
            self.table.insert('', 'end', values=(product['Product'], product['Current Price'], product['Forecast Months'], product['Country'], product['Currency']))
            self._save_state()
            self.log_message(f"Product added: {product['Product']}")
            popup.destroy()

        ttk.Button(frm, text='Confirm', command=confirm).grid(row=5, column=0, columnspan=2, pady=12)

    def _open_edit_popup(self, values, iid):
        idx = self.table.index(iid)
        data = self.product_list[idx]

        popup = tk.Toplevel(self)
        popup.title('Edit Product')
        popup.geometry('380x320')
        popup.transient(self)

        frm = ttk.Frame(popup, padding=12)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text='Product Name:').grid(row=0, column=0, sticky='w')
        entry_name = ttk.Entry(frm); entry_name.insert(0, data.get('Product',''))
        entry_name.grid(row=0, column=1, sticky='ew')

        ttk.Label(frm, text='Current Price:').grid(row=1, column=0, sticky='w')
        entry_price = ttk.Entry(frm); entry_price.insert(0, data.get('Current Price',''))
        entry_price.grid(row=1, column=1, sticky='ew')

        ttk.Label(frm, text='Forecast Months:').grid(row=2, column=0, sticky='w')
        entry_months = ttk.Entry(frm); entry_months.insert(0, data.get('Forecast Months','0'))
        entry_months.grid(row=2, column=1, sticky='ew')

        ttk.Label(frm, text='Country Code (ISO2/ISO3):').grid(row=3, column=0, sticky='w')
        entry_country = ttk.Entry(frm); entry_country.insert(0, data.get('Country',''))
        entry_country.grid(row=3, column=1, sticky='ew')

        ttk.Label(frm, text='Currency Code (e.g. USD):').grid(row=4, column=0, sticky='w')
        entry_currency = ttk.Entry(frm); entry_currency.insert(0, data.get('Currency','USD'))
        entry_currency.grid(row=4, column=1, sticky='ew')

        frm.columnconfigure(1, weight=1)

        def save():
            data['Product'] = entry_name.get().strip()
            data['Current Price'] = entry_price.get().strip()
            data['Forecast Months'] = entry_months.get().strip() or '0'
            data['Country'] = entry_country.get().strip()
            data['Currency'] = entry_currency.get().strip().upper() or 'USD'
            # update table row
            self.table.item(iid, values=(data['Product'], data['Current Price'], data['Forecast Months'], data['Country'], data['Currency']))
            self._save_state()
            self.log_message('Product updated.')
            popup.destroy()

        ttk.Button(frm, text='Save', command=save).grid(row=5, column=0, columnspan=2, pady=12)

    def save_to_excel_action(self):
        if not self.product_list:
            messagebox.showwarning('Warning', 'No products to save.')
            return
        file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])
        if file_path:
            success, msg = CreateExcel.save_to_excel(self.product_list, file_path)
            self.log_message(msg)
            if success:
                messagebox.showinfo('Success', msg)

    def read_excel_action(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if file_path:
            success, msg = ReadExcel.read_excel_to_json_gui(file_path)
            self.log_message(msg)
            if success:
                # try load generated data.json
                try:
                    with open('data.json', 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        # clear table and load
                        for r in self.table.get_children():
                            self.table.delete(r)
                        self.product_list = data
                        for p in data:
                            self.table.insert('', 'end', values=(p.get('Product',''), p.get('Current Price',''), p.get('Forecast Months',''), p.get('Country',''), p.get('Currency','')))
                        self._save_state()
                except Exception as e:
                    self.log_message(f'Could not load data.json: {e}')
                messagebox.showinfo('Success', msg)

    def calculate_forecast_action(self):
        if not self.product_list:
            messagebox.showwarning('Warning', 'No products to forecast.')
            return
        # call Calculate.main
        success, msg = Calculate.main(self.product_list)
        self.log_message(msg)
        if success:
            messagebox.showinfo('Success', msg)

    def log_message(self, message):
        self.log_box.config(state='normal')
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_box.insert('end', f'[{ts}] {message}\n' + '-'*40 + '\n')
        self.log_box.config(state='disabled')
        self.log_box.see('end')


if __name__ == '__main__':
    import sys
    import os

    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    modules_dir = os.path.join(base_dir, 'Modules')
    if modules_dir not in sys.path:
        sys.path.insert(0, modules_dir)

    app = PriceForecastApp()
    app.mainloop()
