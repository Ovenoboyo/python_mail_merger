from tkinter import messagebox


def show_error(title, message):
    messagebox.showerror(title, message)


def show_info(title, message):
    messagebox.showinfo(title, message)