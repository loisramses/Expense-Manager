import tkcalendar as tkCal
import tkinter as tk
import Manager as mn
from tkinter import messagebox
from datetime import datetime
from datetime import date
from tkinter import ttk

class EditMonth:
    def __init__(self, manager: mn.Manager):
        self.manager = manager
        self.root = tk.Toplevel(self.manager.root)
        self.root.title("Edit Month")
        self.root.geometry("600x470")

        self.setup_input_frame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_input_frame(self):
        self.input_frame = tk.Frame(self.root)
        
        # start date input
        self.start_date_label = tk.Label(self.input_frame, text="Start Date: ", font=self.manager.main_font)
        self.start_date_label.grid(row=0, column=0, sticky=tk.EW)
        self.start_date_inputCal = tkCal.Calendar(self.input_frame, firstweekday="sunday")
        self.start_date_inputCal.grid(row=0, column=1, sticky=tk.EW)
        self.start_date_inputCal.selection_set(self.manager.current_sheet_start_date)

        # end date input
        self.end_date_label = tk.Label(self.input_frame, text="End Date: ", font=self.manager.main_font)
        self.end_date_label.grid(row=1, column=0, sticky=tk.EW)
        self.end_date_inputCal = tkCal.Calendar(self.input_frame, firstweekday="sunday")
        self.end_date_inputCal.grid(row=1, column=1, pady=5, sticky=tk.EW)
        self.end_date_inputCal.selection_set(self.manager.current_sheet_end_date)

        # ammount input
        self.ammount_label = tk.Label(self.input_frame, text="Initial Ammount: ", font=self.manager.main_font)
        self.ammount_label.grid(row=2, column=0, sticky=tk.EW)
        self.ammount_input = tk.Text(self.input_frame, wrap=tk.WORD, height=1, width=2)
        self.ammount_input.grid(row=2, column=1, sticky=tk.EW)
        self.ammount_input.insert('1.0', self.manager.current_sheet_initial_ammount)

        # add button
        self.add_button = tk.Button(self.input_frame, text="Confirm", bg='green2', font=self.manager.button_font, width=13, command=self.update_data)
        self.add_button.grid(row=3, column=0)
        
        # clear button
        self.clear_button = tk.Button(self.input_frame, text="Clear", bg='light blue', font=self.manager.button_font, width=13, command=self.clear_data)
        self.clear_button.grid(row=3, column=1)

        # cancel button
        self.cancel_button = tk.Button(self.input_frame, text="Cancel", bg='red', font=self.manager.button_font, width=13, command=self.on_closing)
        self.cancel_button.grid(row=3, column=2)

        self.input_frame.pack(pady=15)

    def update_data(self):
        start_date = self.start_date_inputCal.selection_get().strftime("%d/%m/%Y")
        end_date = self.end_date_inputCal.selection_get().strftime("%d/%m/%Y")
        ammount = self.ammount_input.get('1.0', tk.END).strip()
        if ammount == '':
            messagebox.showwarning("Missing Parameters", "There are empty parameters, cannot perform operation!")
        else:
            if not messagebox.askyesno("Confirm", "Are you sure you want to edit this month?"): return
            info = {
                'start_date': start_date,
                'end_date': end_date,
                'initial_ammount': float(ammount)
            }
            self.manager.edit_current_sheet(info)
            self.on_closing()

    def clear_data(self):
        # reset data
        self.start_date_inputCal.selection_set(date.today())
        self.end_date_inputCal.selection_set(date.today())
        self.ammount_input.delete('1.0', tk.END)

    def run(self):
        self.root.mainloop()

    def stop(self):
        self.root.quit()

    def on_closing(self):
        self.manager.finish_op()
        self.stop()
        self.root.destroy()