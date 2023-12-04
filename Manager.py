import EditData as editFrame
import AddData as addFrame
import AddYear as addYear
import openpyxl as xlrw
import tkinter as tk
import os
from openpyxl.styles import Font
from tkinter import messagebox
from datetime import date
from tkinter import ttk

class Manager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Expense Manager")
        self.root.geometry("1200x700")

        self.folder_path = './expenses/'
        self.performing_operation = False

        self.main_font = ('Arial', 13)
        self.button_font = ('Arial', 9)
        self.cell_font = Font(name='Calibri', size=11, bold=True)
        self.number_format_str = '#,##0.00 €; [Red]-#,##0.00 €'

        self.setup_menubar()
        self.setup_excelExplorerFrame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.root.mainloop();

    def setup_excelExplorerFrame(self):
        self.books_names = os.listdir("./expenses")

        self.excel_explorerFrame = tk.Frame(self.root)

        self.year_boxlist = ttk.Combobox(self.excel_explorerFrame, value=self.books_names)
        self.year_boxlist.current(0)
        self.update_sheets(None)
        self.year_boxlist.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
        self.year_boxlist.bind("<<ComboboxSelected>>", self.update_sheets)

        self.month_boxlist = ttk.Combobox(self.excel_explorerFrame, value=self.sheet_names)
        self.month_boxlist.current(0)
        self.month_boxlist.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        self.month_boxlist.bind("<<ComboboxSelected>>", self.update_expenses)

        self.sheet_

        self.update_expenses(None)
        self.expense_treeList = ttk.Treeview(self.excel_explorerFrame, columns=self.expenses_data[0], show='headings')
        for col in self.expenses_data[0]: self.expense_treeList.heading(col, text=col, anchor=tk.CENTER)
        self.populate_expensesTree()
        self.expense_treeList.grid(row=1, column=1, rowspan=4, columnspan=2, padx=5, pady=5, sticky=tk.NSEW)

        # add expense button
        self.add_button = tk.Button(self.excel_explorerFrame, text="Add Expense", bg='firebrick2', font=('Arial', 9), width=13, command=self.add_expenseToExcel)
        self.add_button.grid(row=1, column=3, padx=5, pady=5, sticky=tk.EW)

        # add revenue button
        self.add_button = tk.Button(self.excel_explorerFrame, text="Add Revenue", bg='green2', font=('Arial', 9), width=13, command=self.add_revenueToExcel)
        self.add_button.grid(row=2, column=3, padx=5, pady=5, sticky=tk.EW)

        # edit button
        self.edit_button = tk.Button(self.excel_explorerFrame, text="Edit", bg='light blue', font=('Arial', 9), command=self.edit_data)
        self.edit_button.grid(row=3, column=3, padx=5, pady=5, sticky=tk.EW)

        # delete button
        self.delete_button = tk.Button(self.excel_explorerFrame, text="Delete Selection", bg='red2', font=('Arial', 9), command=self.delete_selection)
        self.delete_button.grid(row=4, column=3, padx=5, pady=5, sticky=tk.EW)

        self.excel_explorerFrame.pack()

    def setup_menubar(self):
        self.menubar = tk.Menu(self.root)
        
        self.file_menu = tk.Menu(self.menubar, tearoff=0)
        self.file_menu.add_command(label="Create Year", command=self.add_year)
        self.file_menu.add_command(label="Create Month", command=self.on_closing)
        self.file_menu.add_command(label="Save", command=self.on_closing)
        self.file_menu.add_command(label="Close", command=self.on_closing)
        self.menubar.add_cascade(menu=self.file_menu, label="File")

        self.action_menu = tk.Menu(self.menubar, tearoff=0)
        self.action_menu.add_command(label="Add Expense", command=self.add_expenseToExcel)
        self.action_menu.add_command(label="Edit Expense", command=self.edit_data)
        self.action_menu.add_command(label="Delete Expense(s)", command=self.delete_selection)
        self.menubar.add_cascade(menu=self.action_menu, label="Action")

        self.root.config(menu=self.menubar)

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
        for row in self.current_sheet.iter_rows(max_row=self.current_sheet.max_row, max_col=4, values_only=True):
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
            for row in self.expenses_data[1:]:
                self.expense_treeList.insert("", 'end', values=row)

    def create_workbook(self, info):
        new_file_name = info['year_name'] + '.xlsx'

        # create a new workbook, a new sheet and save it
        self.workbook = xlrw.Workbook()
        self.current_sheet = self.workbook.active
        self.create_sheet(info)
        self.workbook.save(self.folder_path + new_file_name)

        # update workbook list
        self.books_names = os.listdir('./expenses')
        self.year_boxlist.config(values=self.books_names)
        self.year_boxlist.current(self.books_names.index(new_file_name))

        info.pop('year_name')
        self.update_sheets(None)

    def add_year(self):
        self.perform_op()
        addYear.AddYear(self)

    def create_sheet(self, info):
        self.current_sheet.title = info['month_name']

        # insert col tags
        self.current_sheet['A1'] = 'DATE'
        self.current_sheet['A1'].font = self.cell_font
        self.current_sheet['B1'] = 'AMOUNT'
        self.current_sheet['B1'].font = self.cell_font
        self.current_sheet['C1'] = 'PURPOSE'
        self.current_sheet['C1'].font = self.cell_font
        self.current_sheet['D1'] = 'DESCRIPTION'
        self.current_sheet['D1'].font = self.cell_font

        # insert dates
        self.current_sheet['G1'] = 'Data início:'
        self.current_sheet['G1'].font = self.cell_font
        self.current_sheet['H1'] = info['init_date']
        self.current_sheet['G2'] = 'Data fim:'
        self.current_sheet['G2'].font = self.cell_font

        # insert ammounts
        self.current_sheet['G4'] = 'Montante inicial:'
        self.current_sheet['G4'].font = self.cell_font
        self.current_sheet['H4'] = float(info['init_ammount'])
        self.current_sheet['H4'].number_format = self.number_format_str
        self.current_sheet['G5'] = 'Montante final:'
        self.current_sheet['G5'].font = self.cell_font
        self.current_sheet['H5'] = '=(H4-H7)'
        self.current_sheet['H5'].number_format = self.number_format_str

        # total spent ammount
        self.current_sheet['G7'] = 'Total gasto:'
        self.current_sheet['G7'].font = self.cell_font
        self.current_sheet['H7'] = '=SUM(B:B)'
        self.current_sheet['H7'].number_format = self.number_format_str

    def add_month(self):
        pass

    def add_row_to_current_sheet(self, row):
        # add row to excel
        self.current_sheet.cell(row=self.row_nmb+1, column=1, value=row[0])
        try:
            self.current_sheet.cell(row=self.row_nmb+1, column=2, value=float(row[1])).number_format = self.number_format_str
        except:
            messagebox.showwarning("Invalid number format", "The value must be seperated by \".\"!")
            return
        self.current_sheet.cell(row=self.row_nmb+1, column=3, value=str(row[2]))
        self.current_sheet.cell(row=self.row_nmb+1, column=4, value=str(row[3]))

        self.expenses_data.append(row)
        self.save_workbook(self.folder_path + self.year_boxlist.get())
        self.row_nmb += 1
        self.populate_expensesTree()

    def add_expenseToExcel(self):
        self.perform_op()
        addFrame.AddData(self, 'expense')

    def add_revenueToExcel(self):
        self.perform_op()
        addFrame.AddData(self, 'revenue')

    def edit_row_on_current_sheet(self, row):
        date = row[0]
        ammount = row[1]
        purpose = row[2]
        description = row[3]
        # update expense tree list
        self.expense_treeList.set(item=self.selection[0], column=0, value=date)
        self.expense_treeList.set(item=self.selection[0], column=1, value=ammount)
        self.expense_treeList.set(item=self.selection[0], column=2, value=purpose)
        self.expense_treeList.set(item=self.selection[0], column=3, value=description)

        # update excel
        target_index = self.expense_treeList.index(self.selection[0])+2
        self.current_sheet.cell(row=target_index, column=1, value=date)
        try:
            self.current_sheet.cell(row=target_index, column=2, value=float(ammount)).number_format = self.number_format_str
        except:
            messagebox.showwarning("Invalid number format", "The value must be seperated by \".\"!")
            return
        self.current_sheet.cell(row=target_index, column=3, value=purpose)
        self.current_sheet.cell(row=target_index, column=4, value=description)

        self.save_workbook(self.folder_path + self.year_boxlist.get())

    def edit_data(self):
        self.selection = self.expense_treeList.selection()
        selection_size = len(self.selection)
        if selection_size == 0:
            messagebox.showwarning("Empty selection", "There are no items selected, cannot perform operation!")
            return
        if selection_size > 1:
            messagebox.showwarning("Too many items selected", "There are too many items selected, cannot perform operation! Edit can only be performed on one item.")
            return
        self.perform_op()
        editFrame.EditExpense(self)

    def delete_selection(self):
        selection = self.expense_treeList.selection()
        selection_size = len(selection)
        if selection_size <= 0:
            messagebox.showwarning("Empty selection", "There are no items selected, cannot perform operation!")
            return
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
        # clear old data
        self.current_sheet.delete_cols(7, 2)

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
        self.current_sheet['H4'].number_format = self.number_format_str
        self.current_sheet['G5'] = 'Montante final:'
        self.current_sheet['G5'].font = self.cell_font
        self.current_sheet['H5'] = '=(H4-H7)'
        self.current_sheet['H5'].number_format = self.number_format_str

        # total spent ammount
        self.current_sheet['G7'] = 'Total gasto:'
        self.current_sheet['G7'].font = self.cell_font
        self.current_sheet['H7'] = '=SUM(B:B)'
        self.current_sheet['H7'].number_format = self.number_format_str

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

Manager()