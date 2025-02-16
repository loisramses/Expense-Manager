import tkcalendar as tkCal
import tkinter as tk
import Manager as mn
from tkinter import messagebox
from datetime import date

class AddYear:
    def __init__(self, manager: mn.Manager):
        self.manager = manager

        self.root = tk.Toplevel(self.manager.root)
        self.root.title('Add Year')
        self.root.geometry("600x350")

        self.setup_input_frame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_input_frame(self):
        self.input_frame = tk.Frame(self.root)
        
        # name input
        self.name_label = tk.Label(self.input_frame, text='Year: ', font=self.manager.main_font)
        self.name_label.grid(row=0, column=0, sticky=tk.EW)
        self.name_input = tk.Text(self.input_frame, wrap=tk.WORD, height=1, width=2)
        self.name_input.grid(row=0, column=1, sticky=tk.EW)

        # month input
        self.month_label = tk.Label(self.input_frame, text='Month: ', font=self.manager.main_font)
        self.month_label.grid(row=1, column=0, sticky=tk.EW)
        self.month_input = tk.Text(self.input_frame, wrap=tk.WORD, height=1, width=2)
        self.month_input.grid(row=1, column=1, sticky=tk.EW)

        # initial date input
        self.date_label = tk.Label(self.input_frame, text="Initial Date: ", font=self.manager.main_font)
        self.date_label.grid(row=2, column=0, sticky=tk.EW)
        self.date_inputCal = tkCal.Calendar(self.input_frame, firstweekday='sunday')
        self.date_inputCal.grid(row=2, column=1, sticky=tk.EW)

        # initial ammount input
        self.ammount_label = tk.Label(self.input_frame, text="Initial Ammount: ", font=self.manager.main_font)
        self.ammount_label.grid(row=3, column=0, sticky=tk.EW)
        self.ammount_input = tk.Text(self.input_frame, wrap=tk.WORD, height=1, width=2)
        self.ammount_input.grid(row=3, column=1, sticky=tk.EW)

        # add button
        self.add_button = tk.Button(self.input_frame, text="Create Year", bg='green2', font=self.manager.button_font, width=13, command=self.create_workbook)
        self.add_button.grid(row=4, column=0)
        
        # clear button
        self.clear_button = tk.Button(self.input_frame, text="Clear", bg='light blue', font=self.manager.button_font, width=13, command=self.clear_data)
        self.clear_button.grid(row=4, column=1)

        # cancel button
        self.cancel_button = tk.Button(self.input_frame, text="Cancel", bg='red', font=self.manager.button_font, width=13, command=self.on_closing)
        self.cancel_button.grid(row=4, column=2)

        self.input_frame.pack(pady=15)

    def create_workbook(self):
        year_name = self.name_input.get('1.0', tk.END).strip()
        month_name = self.month_input.get('1.0', tk.END).strip()
        init_date = self.date_inputCal.selection_get().strftime("%d/%m/%Y")
        init_ammount = self.ammount_input.get('1.0', tk.END).strip()
        if year_name == '' or month_name == '' or init_ammount == '':
            messagebox.showwarning("Missing Parameters", "There are empty parameters, cannot perform operation!")
        else:
            if not messagebox.askyesno("Confirm", "Are you sure you want to create this workbook?"): return
            info = {
                'year_name': year_name,
                'month_name': month_name,
                'init_date': init_date,
                'init_ammount': float(init_ammount)
            }
            self.manager.create_workbook(info)
            self.on_closing()

    def clear_data(self):
        # reset data
        self.name_input.delete('1.0', tk.END)
        self.month_input.delete('1.0', tk.END)
        self.date_inputCal.selection_set(date.today())
        self.ammount_input.delete('1.0', tk.END)

    def run(self):
        self.root.mainloop()

    def stop(self):
        self.root.quit()

    def on_closing(self):
        self.manager.finish_op()
        self.stop()
        self.root.destroy()