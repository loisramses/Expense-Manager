import tkcalendar as tkCal
import tkinter as tk
import Manager as mn
from tkinter import messagebox
from datetime import date
from tkinter import ttk

class AddData:
    def __init__(self, manager: mn.Manager, type_of_op):
        self.type_of_op = type_of_op
        self.text = 'Add Revenue' if self.type_of_op == 'revenue' else 'Add Expense'
        self.op_type = 1 if type_of_op == 'revenue' else -1

        self.manager = manager
        self.root = tk.Toplevel(self.manager.root)
        self.root.title(self.text)
        self.root.geometry("600x370")
        
        self.setup_inputFrame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_inputFrame(self):
        self.input_frame = tk.Frame(self.root)

        # date input
        self.date_label = tk.Label(self.input_frame, text="Date: ", font=self.manager.main_font)
        self.date_label.grid(row=0, column=0, sticky=tk.EW)
        self.date_inputCal = tkCal.Calendar(self.input_frame, firstweekday='sunday')
        self.date_inputCal.grid(row=0, column=1, sticky=tk.EW)

        # ammount input
        self.ammount_label = tk.Label(self.input_frame, text="Ammount: ", font=self.manager.main_font)
        self.ammount_label.grid(row=1, column=0, sticky=tk.EW)
        self.ammount_input = tk.Text(self.input_frame, wrap=tk.WORD, height=1, width=2)
        self.ammount_input.grid(row=1, column=1, sticky=tk.EW)

        # purpose input
        self.purpose_label = tk.Label(self.input_frame, text="Purpose: ", font=self.manager.main_font)
        self.purpose_label.grid(row=2, column=0, sticky=tk.EW)
        self.purpose_input = tk.Text(self.input_frame, wrap=tk.WORD, height=1, width=2)
        self.purpose_input.grid(row=2, column=1, sticky=tk.EW)

        # description input
        self.description_label = tk.Label(self.input_frame, text="Description: ", font=self.manager.main_font)
        self.description_label.grid(row=3, column=0, sticky=tk.EW)
        self.description_input = tk.Text(self.input_frame, wrap=tk.WORD, height=3, width=2)
        self.description_input.grid(row=3, column=1, sticky=tk.EW)

        # category input
        self.category_label = tk.Label(self.input_frame, text="Category: ", font=self.manager.main_font)
        self.category_label.grid(row=4, column=0, sticky=tk.EW)
        self.category_input = ttk.Combobox(self.input_frame, value=self.manager.categories_list if self.type_of_op == 'expense' else self.manager.categories_list[3], state='readonly')
        self.category_input.current(0)
        self.category_input.grid(row=4, column=1, sticky=tk.EW)

        # add button
        self.add_button = tk.Button(self.input_frame, text=self.text, bg='green2', font=self.manager.button_font, width=13, command=self.add_data)
        self.add_button.grid(row=5, column=0)
        
        # clear button
        self.clear_button = tk.Button(self.input_frame, text="Clear", bg='light blue', font=self.manager.button_font, width=13, command=self.clear_data)
        self.clear_button.grid(row=5, column=1)

        # cancel button
        self.cancel_button = tk.Button(self.input_frame, text="Cancel", bg='red', font=self.manager.button_font, width=13, command=self.on_closing)
        self.cancel_button.grid(row=5, column=2)

        self.input_frame.pack(pady=15)

    def add_data(self):
        date = self.date_inputCal.selection_get().strftime("%d/%m/%Y")
        ammount = self.ammount_input.get('1.0', tk.END).strip()
        purpose = self.purpose_input.get('1.0', tk.END).strip()
        description = self.description_input.get('1.0', tk.END).strip()
        category = self.category_input.get()

        if ammount == '' or purpose == '' or description == '':
            messagebox.showwarning("Missing Parameters", "There are empty parameters, cannot perform operation!")
        else:
            if not messagebox.askyesno("Confirm expense", "Are you sure you want to add this expense?"): return
            self.new_row = (date, float(ammount)*self.op_type, purpose, description, category)
            self.manager.add_row_to_current_sheet(self.new_row)
            self.on_closing()

    def clear_data(self):
        # reset data
        self.date_inputCal.selection_set(date.today())
        self.ammount_input.delete('1.0', tk.END)
        self.purpose_input.delete('1.0', tk.END)
        self.description_input.delete('1.0', tk.END)
        self.category_input.current(0)

    def run(self):
        self.root.mainloop()

    def stop(self):
        self.root.quit()

    def on_closing(self):
        self.manager.finish_op()
        self.stop()
        self.root.destroy()