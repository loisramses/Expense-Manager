import tkcalendar as tkCal
import tkinter as tk
from tkinter import messagebox
from datetime import date

class AddExpense:
    def __init__(self, manager):
        self.manager = manager
        self.root = tk.Toplevel(manager.root)
        self.root.title("Add expense")
        self.root.geometry("400x370")
        
        self.setup_inputFrame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.root.mainloop()

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

    def add_dataToExcel(self):
        date = self.date_inputCal.selection_get().strftime("%d/%m/%Y")
        ammount = self.ammount_input.get('1.0', tk.END).strip()
        purpose = self.purpose_input.get('1.0', tk.END).strip()
        description = self.description_input.get('1.0', tk.END).strip()

        if date == '' or ammount == '' or purpose == '' or description == '':
            messagebox.showwarning("Missing Parameters", "There are empty parameters, cannot perform operation!")
        else:
            if not messagebox.askyesno("Confirm expense", "Are you sure you want to add this expense?"): return
            new_row = (date, ammount, purpose, description)
            self.manager.add_row_to_current_sheet(new_row)
            self.on_closing()

    def clear_data(self):
        # reset data
        self.date_inputCal.selection_set(date.today())
        self.ammount_input.delete('1.0', tk.END)
        self.purpose_input.delete('1.0', tk.END)
        self.description_input.delete('1.0', tk.END)

    def on_closing(self):
        self.manager.finish_op()
        self.root.destroy()

# AddExpense()