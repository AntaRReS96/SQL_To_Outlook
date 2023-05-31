from ProjectClosing_GUI import SQLMailSenderGUI
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    app = SQLMailSenderGUI(master=root)
    app.mainloop()