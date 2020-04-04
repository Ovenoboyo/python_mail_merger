import os
import tkinter as tk
from functools import partial
from tkinter import filedialog
from tkinter import ttk

from main import main_seperate_doc, main_new_page, save_ppt_as_pdf
from tkinter_message import show_error

root = tk.Tk()

style = ttk.Style(root)
style.theme_use("clam")
root.configure(bg=style.lookup('TFrame', 'background'))

path_xlsx = ""
path_docx = ""
path_pptx = ""
out_path = ""


def save_paths():
    with open("last_used_paths", "w+") as f:
        f.seek(0)
        f.write(path_xlsx + "\n")
        f.write(path_docx + "\n")
        f.write(path_pptx + "\n")
        f.write(out_path)
        f.truncate()


def xlsx_file():
    global path_xlsx
    rep = filedialog.askopenfilenames(parent=root, initialdir="/" if not path_xlsx else os.path.dirname(path_xlsx),
                                      initialfile='tmp',
                                      filetypes=[("Excel", "*.xlsx")])
    if rep:
        path_xlsx = rep[0]
        e1.insert(0, path_xlsx)


def docx_file():
    global path_docx
    rep = filedialog.askopenfilenames(parent=root, initialdir="/" if not path_docx else os.path.dirname(path_docx),
                                      initialfile='tmp',
                                      filetypes=[("Word Document", "*.docx")])
    if rep:
        path_docx = rep[0]
        e2.insert(0, path_docx)


def pptx_file():
    global path_pptx
    rep = filedialog.askopenfilenames(parent=root, initialdir="/" if not path_pptx else os.path.dirname(path_pptx),
                                      initialfile='tmp',
                                      filetypes=[("PowerPoint Presentation", "*.pptx")])
    if rep:
        path_pptx = rep[0]
        e3.insert(0, path_pptx)


def out_folder():
    global out_path
    rep = filedialog.askdirectory(parent=root, initialdir="/" if not out_path else out_path)
    if rep:
        out_path = rep
        e4.insert(0, out_path)


def run(seperate_files=True):
    if path_xlsx:
        save_paths()
        if seperate_files:
            main_seperate_doc(xlsx_path=path_xlsx, docx_path=path_docx, pptx_path=path_pptx, out_path=out_path)
        else:
            main_new_page(xlsx_path=path_xlsx, docx_path=path_docx, pptx_path=path_pptx, out_path=out_path)
    else:
        show_error("Path error",  "Excel path can't be empty")


def ppt_as_pdf():
    save_ppt_as_pdf(out_path)


if __name__ == '__main__':
    if os.path.exists("last_used_paths"):
        f = open("last_used_paths")
        lines = f.readlines()
        f.close()

        if len(lines) == 4:
            path_xlsx = lines[0].strip() if os.path.exists(lines[0].strip()) else ""
            path_docx = lines[1].strip() if os.path.exists(lines[1].strip()) else ""
            path_pptx = lines[2].strip() if os.path.exists(lines[2].strip()) else ""
            out_path = lines[3].strip() if os.path.exists(lines[2].strip()) else ""

    ttk.Label(root, text='Paths').grid(row=0, column=0, padx=4, pady=4, sticky='ew')
    e1 = ttk.Entry(root)
    e1.grid(row=1, column=0, padx=4, pady=4)
    e1.insert(0, path_xlsx)
    e2 = ttk.Entry(root)
    e2.grid(row=2, column=0, padx=4, pady=4)
    e2.insert(0, path_docx)
    e3 = ttk.Entry(root)
    e3.grid(row=3, column=0, padx=4, pady=4)
    e3.insert(0, path_pptx)
    e4 = ttk.Entry(root)
    e4.grid(row=4, column=0, padx=4, pady=4)
    e4.insert(0, out_path)

    ttk.Button(root, text="Open xlsx", command=xlsx_file).grid(row=1, column=1, padx=4, pady=4, sticky='ew')
    ttk.Button(root, text="Open docx", command=docx_file).grid(row=2, column=1, padx=4, pady=4, sticky='ew')
    ttk.Button(root, text="Open pptx", command=pptx_file).grid(row=3, column=1, padx=4, pady=4, sticky='ew')
    ttk.Button(root, text="Select output folder", command=out_folder).grid(row=4, column=1, padx=4, pady=4, sticky='ew')

    ttk.Button(root, text="Start (Separate files)", command=partial(run, True)).grid(row=5, column=1, padx=4, pady=4, sticky='ew')
    ttk.Button(root, text="Start (Same file)", command=partial(run, False)).grid(row=5, column=2, padx=4, pady=4, sticky='ew')
    ttk.Button(root, text="pptx to pdf", command=ppt_as_pdf).grid(row=5, column=0, padx=4, pady=4,
                                                                                 sticky='ew')
    root.mainloop()
