from tkinter import Tk, messagebox


def show_error_dialog(title, message):
    window = Tk()
    window.withdraw()
    messagebox.showerror(title=title, message=message)
    window.destroy()
