import tkinter as tk

class MainPage:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Add expense")
        self.root.geometry("600x700")
        
        self.setup_inputFrame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.root.mainloop()

    def setup_inputFrame(self):
        pass

    def on_closing(self):
        self.root.destroy()
        pass