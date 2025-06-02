# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import openpyxl
import os
import tempfile
import webbrowser
import pandas as pd

class ExcelSheetNavigator:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sheet Navigator")
        self.root.geometry("1000x600")

        self.workbook = None
        self.sheet_names = []

        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')

        # --- Advanced Custom Styles ---
        # General background
        style.configure('TFrame', background='#eef2f7')
        style.configure('TLabelFrame', background='#eef2f7', foreground='#2c3e50', font=('Segoe UI', 11, 'bold'))
        style.configure('TLabel', background='#eef2f7', foreground='#34495e', font=('Segoe UI', 10))

        # Buttons
        style.configure('TButton', 
                        font=('Segoe UI', 10, 'bold'), 
                        padding=[15, 10], # Increased padding [width, height]
                        borderwidth=0, 
                        relief='flat', # Use flat relief
                        background='#3498db', # Primary blue
                        foreground='#ffffff')
        style.map('TButton', 
                  background=[('active', '#2980b9'), ('!disabled', '#3498db')], # Darker blue on active
                  foreground=[('active', '#ffffff'), ('!disabled', '#ffffff')]) # White text remains

        # Entry fields
        style.configure('TEntry', 
                        padding=[10, 8], # Increased padding
                        fieldbackground='#ffffff', 
                        foreground='#34495e', 
                        borderwidth=1, 
                        relief='solid')
        style.map('TEntry', 
                  fieldbackground=[('focus', '#ecf0f1')]) # Light grey on focus

        # Treeview (for sheet list and data)
        style.configure('Treeview', 
                        rowheight=30, # Increased row height
                        fieldbackground='#ffffff', 
                        background='#ffffff', 
                        foreground='#34495e', 
                        font=('Segoe UI', 10))
        style.configure('Treeview.Heading', 
                        font=('Segoe UI', 10, 'bold'), 
                        background='#bdc3c7', # Light grey header
                        foreground='#2c3e50',
                        relief='flat')
        style.map('Treeview.Heading', background=[('active', '#95a5a6')]) # Darker grey on active
        style.configure('Treeview', borderwidth=1, relief='solid') # Add border to treeview

        # Scrollbars
        style.configure('TScrollbar', 
                        troughcolor='#ecf0f1', # Light grey trough
                        bordercolor='#bdc3c7', # Medium grey border
                        background='#bdc3c7', # Medium grey background
                        arrowcolor='#34495e') # Dark grey arrows
        style.map('TScrollbar', background=[('active', '#95a5a6')]) # Darker grey on active

        # --- Layout Adjustments ---
        # Increase padding around frames
        self.root.configure(background='#eef2f7') # Root window background
        top_frame = ttk.Frame(self.root, padding="15")
        top_frame.pack(pady=15, padx=15, fill=tk.X)

        self.sheet_list_frame = ttk.Frame(self.root, padding="15")
        self.sheet_list_frame.pack(pady=5, padx=15, fill=tk.BOTH, expand=True)

        action_frame = ttk.Frame(self.root, padding="15")
        action_frame.pack(pady=5, padx=15, fill=tk.X)

        self.tree_frame = ttk.Frame(self.root, padding="15")
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        self.cell_edit_frame = ttk.Frame(self.root, padding="15")
        self.cell_edit_frame.pack(pady=5, padx=15, fill=tk.X)

        ttk.Label(self.cell_edit_frame, text="Edit Cell Value:").pack(side=tk.LEFT, padx=5)
        self.cell_value_entry = ttk.Entry(self.cell_edit_frame, width=50)
        self.cell_value_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.update_cell_btn = ttk.Button(self.cell_edit_frame, text="Update Cell", command=self.update_selected_cell, state=tk.DISABLED)
        self.update_cell_btn.pack(side=tk.LEFT, padx=5)

        self.selected_cell_coords = None # Store (item_id, col_index) of selected cell
        self.current_sheet = None # Store the currently displayed openpyxl sheet object

        ttk.Button(top_frame, text="Load Excel File", command=self.load_excel).pack(side=tk.LEFT, padx=10)

        ttk.Label(top_frame, text="Search Sheet:").pack(side=tk.LEFT, padx=10)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.update_listbox)
        search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)

        self.sheet_tree = ttk.Treeview(self.sheet_list_frame, columns=('Sheet Name',), show='headings')
        self.sheet_tree.heading('Sheet Name', text='Sheet Name')
        self.sheet_tree.column('Sheet Name', width=400, anchor='w')

        self.sheet_tree.tag_configure('match', background='yellow', foreground='black')

        sheet_vsb = ttk.Scrollbar(self.sheet_list_frame, orient="vertical", command=self.sheet_tree.yview)
        sheet_vsb.pack(side='right', fill='y')
        self.sheet_tree.configure(yscrollcommand=sheet_vsb.set)

        sheet_hsb = ttk.Scrollbar(self.sheet_list_frame, orient="horizontal", command=self.sheet_tree.xview)
        sheet_hsb.pack(side='bottom', fill='x')
        self.sheet_tree.configure(xscrollcommand=sheet_hsb.set)

        self.sheet_tree.pack(side='left', fill=tk.BOTH, expand=True)
        self.sheet_tree.bind('<<TreeviewSelect>>', self.display_sheet_data)

        self.open_excel_btn = ttk.Button(action_frame, text="Open in Excel", command=self.open_in_excel, state=tk.DISABLED)
        self.open_excel_btn.pack(side=tk.LEFT, padx=5)

        self.open_gsheets_btn = ttk.Button(action_frame, text="Open in Google Sheets", command=self.open_in_gsheets, state=tk.DISABLED)
        self.open_gsheets_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(action_frame, text="Export Sheet", command=self.export_sheet, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            self.workbook = openpyxl.load_workbook(file_path, data_only=True)
            self.sheet_names = self.workbook.sheetnames
            self.current_filepath = file_path # Store the current file path
            self.update_listbox()
            messagebox.showinfo("Success", f"Loaded {len(self.sheet_names)} sheets.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load workbook: {e}")

    def update_listbox(self, *args):
        search_term = self.search_var.get().lower()
        for item in self.sheet_tree.get_children():
            self.sheet_tree.delete(item)

        for name in self.sheet_names:
            if search_term and search_term in name.lower():
                self.sheet_tree.insert('', tk.END, values=(name,), tags=('match',))
            else:
                self.sheet_tree.insert('', tk.END, values=(name,))

    def display_sheet_data(self, event):
        selection = self.sheet_tree.selection()
        if not selection:
            self.open_excel_btn.config(state=tk.DISABLED)
            self.open_gsheets_btn.config(state=tk.DISABLED)
            self.export_btn.config(state=tk.DISABLED)
            return

        sheet_name = self.sheet_tree.item(selection[0], 'values')[0]

        self.open_excel_btn.config(state=tk.NORMAL)
        self.open_gsheets_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.NORMAL)

        sheet = self.workbook[sheet_name]
        self.current_sheet = sheet # Store the current sheet

        for widget in self.tree_frame.winfo_children():
            widget.destroy()

        tree = ttk.Treeview(self.tree_frame)
        tree.pack(fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side='bottom', fill='x')
        tree.configure(xscrollcommand=hsb.set)

        max_col = 0
        rows = []
        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            row_data = [str(cell) if cell is not None else "" for cell in row]
            rows.append(row_data)
            max_col = max(max_col, len(row_data))
            if i > 100:
                break

        tree["columns"] = list(range(max_col))
        tree["show"] = "headings"

        for i in range(max_col):
            tree.heading(i, text=f"Col {i+1}")
            tree.column(i, width=100, anchor='w')

        for row in rows:
            padded_row = row + [""] * (max_col - len(row))
            tree.insert("", "end", values=padded_row)

        self.data_tree = tree # Keep a reference to the data treeview
        self.data_tree.bind('<<TreeviewSelect>>', self.on_data_cell_select)

    def export_selected_sheet(self, file_path, sheet_name):
        from openpyxl import Workbook
        src_sheet = self.workbook[sheet_name]
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        for row in src_sheet.iter_rows(values_only=True):
            ws.append(list(row))
        wb.save(file_path)

    def open_in_excel(self):
        selection = self.sheet_tree.selection()
        if not selection:
            return
        sheet_name = self.sheet_tree.item(selection[0], 'values')[0]
        try:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                temp_filepath = tmp.name
                self.export_selected_sheet(temp_filepath, sheet_name)

            os.startfile(temp_filepath)
        except Exception as e:
             messagebox.showerror("Error", f"Failed to open in Excel: {e}")

    def open_in_gsheets(self):
        selection = self.sheet_tree.selection()
        if not selection:
            return
        sheet_name = self.sheet_tree.item(selection[0], 'values')[0]
        try:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                temp_filepath = tmp.name
                self.export_selected_sheet(temp_filepath, sheet_name)

            webbrowser.open(f'file://{os.path.abspath(temp_filepath)}')
            messagebox.showinfo("Google Sheets", "A copy of the sheet has been opened in your browser. Upload it to Google Sheets to view online.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open in Google Sheets: {e}")

    def export_sheet(self):
        selection = self.sheet_tree.selection()
        if not selection:
            return
        sheet_name = self.sheet_tree.item(selection[0], 'values')[0]

        file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")], initialfile=f"{sheet_name}.xlsx")
        if not file_path:
            return

        try:
            self.export_selected_sheet(file_path, sheet_name)
            messagebox.showinfo("Exported", f"Sheet '{sheet_name}' exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred during export: {e}")

    def lookup_door_link(self):
        # This method is removed as per the instructions
        pass

    def update_selected_cell(self):
        if not self.selected_cell_coords or not self.current_sheet:
            messagebox.showwarning("Warning", "No cell selected or no sheet loaded.")
            return

        item_id, col_index = self.selected_cell_coords
        new_value = self.cell_value_entry.get()

        try:
            # Find the actual row index in the openpyxl sheet
            # This is a simplified approach assuming Treeview order matches sheet order up to the loaded rows
            tree_items = self.data_tree.get_children()
            try:
                row_index_in_tree = tree_items.index(item_id)
            except ValueError:
                messagebox.showerror("Error", "Could not find selected item in the display.")
                return

            # openpyxl is 1-based index
            excel_row_index = row_index_in_tree + 1
            excel_col_index = col_index + 1

            # Update the openpyxl sheet
            # Need to handle potential data type conversions if necessary
            self.current_sheet.cell(row=excel_row_index, column=excel_col_index, value=new_value)

            # Update the Treeview display
            current_values = list(self.data_tree.item(item_id, 'values'))
            current_values[col_index] = new_value
            self.data_tree.item(item_id, values=current_values)

            # Save the workbook
            # Need the original file path. Store it when loading the file.
            if hasattr(self, 'current_filepath') and self.current_filepath:
                self.workbook.save(self.current_filepath)
                messagebox.showinfo("Saved", "Workbook saved successfully.")
            else:
                messagebox.showwarning("Warning", "Original file path not available. Cannot auto-save.")


        except Exception as e:
            messagebox.showerror("Error", f"Failed to update cell: {e}")

    def on_data_cell_select(self, event):
        selection = self.data_tree.selection()
        if not selection:
            self.cell_value_entry.delete(0, tk.END)
            self.update_cell_btn.config(state=tk.DISABLED)
            self.selected_cell_coords = None # Clear selected cell coordinates
            return

        # Get the item (row) and the column that was clicked
        item_id = selection[0]
        column_id = self.data_tree.identify_column(event.x)

        # The column_id is in the format #N, where N is the 1-based index
        # Convert it to a 0-based index for list access
        if column_id and column_id.startswith('#'):
            col_index = int(column_id[1:]) - 1
        else:
            col_index = 0 # Default to first column if identification fails

        # Get the row index (Treeview selection is based on item id, need to find its position)
        # This is not straightforward. For now, let's rely on the displayed row values.
        # A more robust solution might require storing row indices alongside data.

        # For simplicity and to proceed, we'll assume the order in the Treeview
        # matches the order of rows fetched from openpyxl and store the item_id.
        # We'll retrieve the actual row index when updating.

        item = self.data_tree.item(item_id)
        row_values = item['values']

        if row_values and 0 <= col_index < len(row_values):
            cell_value = str(row_values[col_index])
            self.cell_value_entry.delete(0, tk.END)
            self.cell_value_entry.insert(0, cell_value)
            self.update_cell_btn.config(state=tk.NORMAL)

            # Store the selected cell's item ID and column index
            self.selected_cell_coords = (item_id, col_index)
        else:
            self.cell_value_entry.delete(0, tk.END)
            self.update_cell_btn.config(state=tk.DISABLED)
            self.selected_cell_coords = None

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use('clam')

    app = ExcelSheetNavigator(root)
    root.mainloop()
