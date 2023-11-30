import tkcalendar as tkCal
import openpyxl as xlrw
import tkinter as tk
import AddExpense as addFrame
import os
from tkinter import messagebox
from openpyxl.styles import Font
from datetime import datetime
from datetime import date
from tkinter import ttk

class MyGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Expense Manager")
        self.root.geometry("1200x700")

        self.folder_path = './expenses/'
        self.performing_operation = False

        self.cell_font = Font(name='Calibri', size=11, bold=True)

        # self.setup_inputFrame()
        self.setup_excelExplorerFrame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.root.mainloop();
    
    def setup_inputFrame(self):
        self.input_frame = tk.Frame(self.root)

        self.input_frame.columnconfigure(0, weight=1)
        self.input_frame.columnconfigure(1, weight=1)

        # date input
        self.date_label = tk.Label(self.input_frame, text="Date: ", font=('Arial', 13))
        self.date_label.grid(row=0, column=0, sticky=tk.W+tk.E)
        self.date_inputCal = tkCal.Calendar(self.input_frame, date_pattern='dd/mm/yyyy', firstweekday="sunday")
        self.date_inputCal.grid(row=0, column=1, sticky=tk.W+tk.E)

        # ammount input
        self.ammount_label = tk.Label(self.input_frame, text="Ammount: ", font=('Arial', 13))
        self.ammount_label.grid(row=1, column=0, sticky=tk.W+tk.E)
        self.ammount_input = tk.Text(self.input_frame,height=1, width=2)
        self.ammount_input.grid(row=1, column=1, sticky=tk.W+tk.E)

        # purpose input
        self.purpose_label = tk.Label(self.input_frame, text="Purpose: ", font=('Arial', 13))
        self.purpose_label.grid(row=2, column=0, sticky=tk.W+tk.E)
        self.purpose_input = tk.Text(self.input_frame,height=1, width=2)
        self.purpose_input.grid(row=2, column=1, sticky=tk.W+tk.E)

        # description input
        self.description_label = tk.Label(self.input_frame, text="Description: ", font=('Arial', 13))
        self.description_label.grid(row=3, column=0, sticky=tk.W+tk.E)
        self.description_input = tk.Text(self.input_frame,height=3, width=2)
        self.description_input.grid(row=3, column=1, sticky=tk.W+tk.E)

        # add button
        self.add_button = tk.Button(self.input_frame, text="Add Expense", bg='light green', font=('Arial', 9), width=13, command=self.add_dataToExcel)
        self.add_button.grid(row=4, column=0)
        
        # clear button
        self.clear_button = tk.Button(self.input_frame, text="Clear", bg='red', font=('Arial', 9), width=13, command=self.clear_data)
        self.clear_button.grid(row=4, column=1)

        self.input_frame.pack(pady=15)

    def setup_excelExplorerFrame(self):
        self.books_names = os.listdir("./expenses")

        self.excel_explorerFrame = tk.Frame(self.root)
        self.excel_explorerFrame.columnconfigure(0, weight=1, pad=5)
        self.excel_explorerFrame.columnconfigure(1, weight=1, pad=5)

        self.year_boxlist = ttk.Combobox(self.excel_explorerFrame, value=self.books_names)
        self.year_boxlist.current(0)
        self.update_sheets(None)
        self.year_boxlist.grid(row=0, column=0, sticky=tk.E)
        self.year_boxlist.bind("<<ComboboxSelected>>", self.update_sheets)

        self.month_boxlist = ttk.Combobox(self.excel_explorerFrame, value=self.sheet_names)
        self.month_boxlist.current(0)
        self.month_boxlist.grid(row=0, column=1, sticky=tk.W)
        self.month_boxlist.bind("<<ComboboxSelected>>", self.update_expenses)

        self.update_expenses(None)
        self.expense_treeList = ttk.Treeview(self.excel_explorerFrame, columns=self.expenses_data[0], show='headings')
        for col in self.expenses_data[0]: self.expense_treeList.heading(col, text=col, anchor=tk.CENTER)
        self.populate_expensesTree()
        self.expense_treeList.grid(row=1, column=0, rowspan=3, columnspan=2, padx=5, pady=5, sticky=tk.NSEW)

        # add button
        self.add_button = tk.Button(self.excel_explorerFrame, text="Add Expense", bg='light green', font=('Arial', 9), width=13, command=self.add_dataToExcel)
        self.add_button.grid(row=1, column=2, padx=5, pady=5, sticky=tk.EW)

        # edit button
        self.edit_button = tk.Button(self.excel_explorerFrame, text="Edit Selection", bg='light blue', font=('Arial', 9), command=self.edit_selection)
        self.edit_button.grid(row=2, column=2, padx=5, pady=5, sticky=tk.EW)

        # delete button
        self.delete_button = tk.Button(self.excel_explorerFrame, text="Delete Selection", bg='red', font=('Arial', 9), command=self.delete_selection)
        self.delete_button.grid(row=3, column=2, padx=5, pady=5, sticky=tk.EW)

        self.excel_explorerFrame.pack()

    def update_sheets(self, event):
        # close the previous book
        if hasattr(self, 'workbook'):
            self.workbook.close()

        self.workbook = xlrw.load_workbook(self.folder_path + self.year_boxlist.get())
        self.sheet_names = self.workbook.sheetnames
        if hasattr(self, 'month_boxlist'):
            self.month_boxlist.config(value=self.sheet_names)
            self.month_boxlist.current(0)
            self.update_expenses(None)

    def update_expenses(self, event):
        self.row_nmb = 0
        if hasattr(self, 'current_sheet'):
            self.save_workbook(self.folder_path + self.year_boxlist.get()) # save information when changing worksheets
        self.current_sheet = self.get_current_sheet()
        self.expenses_data = []
        for row in self.current_sheet.iter_rows(min_row=1, max_row=self.current_sheet.max_row, max_col=4, values_only=True):
            if (None not in row):
                if type(row[0]) is not str: # to ignore the first row
                    row = (row[0].strftime("%d/%m/%Y"), row[1], row[2], row[3])
                self.expenses_data.append(row)
                self.row_nmb += 1
        self.populate_expensesTree()

    def populate_expensesTree(self):
        if hasattr(self, 'expense_treeList'):
            for item in self.expense_treeList.get_children():
                self.expense_treeList.delete(item)
            for i, row in enumerate(self.expenses_data[1:]):
                self.expense_treeList.insert("", i, values=row)

    def add_row_to_current_sheet(self, row):
        self.current_sheet.cell(row=self.row_nmb+1, column=1, value=row[0])
        try:
            self.current_sheet.cell(row=self.row_nmb+1, column=2, value=float(row[1])).number_format = '#,##0.00 €'
        except:
            messagebox.showwarning("Invalid number format", "The value must be seperated by \".\"!")
            return
        self.current_sheet.cell(row=self.row_nmb+1, column=3, value=str(row[2]))
        self.current_sheet.cell(row=self.row_nmb+1, column=4, value=str(row[3]))

        self.expenses_data.append(row)
        self.save_workbook(self.folder_path + self.year_boxlist.get())
        self.row_nmb += 1
        self.populate_expensesTree()

    def add_dataToExcel(self):
        self.perform_op()
        addFrame.AddExpense(self)
        # self.finish_op()
        # row = self.add_box.get_new_row()
        # print(row)
        # if row is not None:
        #     print(self.add_box.get_new_row())
        #     self.add_row_to_current_sheet(row)
        #     self.populate_expensesTree()
        # date = self.date_inputCal.selection_get().strftime("%d/%m/%Y")
        # ammount = self.ammount_input.get('1.0', tk.END).strip()
        # purpose = self.purpose_input.get('1.0', tk.END).strip()
        # description = self.description_input.get('1.0', tk.END).strip()

        # if date == '' or ammount == '' or purpose == '' or description == '':
        #     messagebox.showwarning("Missing Parameters", "There are empty parameters, cannot perform operation!")
        # else:
        #     row = (date, ammount, purpose, description)
        #     if not messagebox.askyesno("Confirm expense", "Are you sure you want to add this expense?"): return
        #     self.add_row_to_current_sheet(row)
        #     self.populate_expensesTree()

    # falta ainda editar realmente as coisas
    def edit_selection(self):
        selection = self.expense_treeList.selection()
        selection_size = len(selection)
        if selection_size == 0:
            messagebox.showwarning("Empty selection", "There are no items selected, cannot perform operation!")
            return
        if selection_size > 1:
            messagebox.showwarning("Too many items selected", "There are too many items selected, cannot perform operation! Edit can only be performed on one item.")
            return
        
        # set the data on the input frame
        self.date_inputCal.selection_set(datetime.strptime(self.expense_treeList.item(selection[0])['values'][0], "%d/%m/%Y"))
        self.ammount_input.replace('1.0', tk.END, self.expense_treeList.item(selection[0])['values'][1])
        self.purpose_input.replace('1.0', tk.END, self.expense_treeList.item(selection[0])['values'][2])
        self.description_input.replace('1.0', tk.END, self.expense_treeList.item(selection[0])['values'][3])

        # print(self.expense_treeList.item(selection[0])['values'])
        # print(datetime.strptime(self.expense_treeList.item(selection[0])['values'][0], "%d/%m/%Y"))
        # print(len(selection))
        # for row in selection:
        #     print(row)

    def delete_selection(self):
        selection = self.expense_treeList.selection()
        selection_size = len(selection)
        if selection_size <= 0:
            messagebox.showwarning("Empty selection", "There are no items selected, cannot perform operation!")        # selection = self.expense_treeList.selection()
        if messagebox.askyesno("Confirm deletion", "Are you sure you want to delete all the selected items? This will permanently delete the selected items!"):
            for row in selection:
                self.current_sheet.delete_rows(self.expense_treeList.index(row)+2, 1) # +2 bc of header and index starts at 0
                self.expense_treeList.delete(row)
                self.row_nmb -= 1

            self.update_expenses(None)
            self.populate_expensesTree()
            self.save_workbook(self.folder_path + self.year_boxlist.get())

    def get_current_sheet(self):
        sheet = self.workbook[self.month_boxlist.get()]
        self.current_sheet_start_date = sheet['H1'].value
        self.current_sheet_end_date = sheet['H2'].value
        self.current_sheet_initial_ammount = sheet['H4'].value
        return sheet

    def save_workbook(self, folder_path):
        # insert dates
        self.current_sheet['G1'] = 'Data início:'
        self.current_sheet['G1'].font = self.cell_font
        self.current_sheet['H1'] = self.current_sheet_start_date
        self.current_sheet['G2'] = 'Data fim:'
        self.current_sheet['G2'].font = self.cell_font
        self.current_sheet['H2'] = self.current_sheet_end_date

        # insert ammounts
        self.current_sheet['G4'] = 'Montante inicial:'
        self.current_sheet['G4'].font = self.cell_font
        self.current_sheet['H4'] = self.current_sheet_initial_ammount
        self.current_sheet['H4'].number_format = '#,##0.00 €'
        self.current_sheet['G5'] = 'Montante final:'
        self.current_sheet['G5'].font = self.cell_font
        self.current_sheet['H5'] = '=(H4-H7)'
        self.current_sheet['H5'].number_format = '#,##0.00 €'

        # total spent ammount
        self.current_sheet['G7'] = 'Total gasto:'
        self.current_sheet['G7'].font = self.cell_font
        self.current_sheet['H7'] = '=SUM(B:B)'
        
        self.current_sheet['H7'].number_format = '#,##0.00 €'

        self.workbook.save(self.folder_path + self.year_boxlist.get())

    def clear_data(self):
        # reset data
        self.date_inputCal.selection_set(date.today())
        self.ammount_input.delete('1.0', tk.END)
        self.purpose_input.delete('1.0', tk.END)
        self.description_input.delete('1.0', tk.END)

    def perform_op(self):
        self.performing_operation = True

    def finish_op(self):
        self.performing_operation = False

    def on_closing(self):
        if not self.performing_operation:
            self.workbook.close()
            self.root.destroy()

MyGUI()