import copy
import os
import shutil
from os.path import isfile, join

import win32com.client

from docx_reader import init_document, MergeDocument
from excel_reader import init_sheet, get_headers, get_values, sort_dicts
from pptx_reader import init_ppt, MergePPT
from tkinter_message import show_info, show_error


def wipe_out_folder(out_path):
    for filename in os.listdir(out_path):
        file_path = os.path.join(out_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            show_error("Error", str(e))


def main_new_page(xlsx_path, docx_path, pptx_path, out_path):
    wipe_out_folder(out_path)
    active_sheet = init_sheet(xlsx_path)
    headers = get_headers(active_sheet)
    values = {}
    for i, items in enumerate(headers):
        values.update({items: get_values(active_sheet, (i + 1))})

    max_count = 0
    for value in values.values():
        if len(value) > max_count:
            max_count = len(value)

    ppt_document = init_ppt(pptx_path)
    docx_document = init_document(docx_path)
    doc_ppt = MergePPT(matches=sort_dicts(values),
                       path=(os.path.join(out_path, str(os.path.basename(pptx_path).split(".")[0]) + "-new.pptx")),
                       presentation=copy.deepcopy(ppt_document))
    doc_ppt.replace_matches_newslide(max_count)
    doc_ppt.save()

    doc_word = MergeDocument(matches=sort_dicts(values),
                       path=(os.path.join(out_path, str(os.path.basename(docx_path).split(".")[0]) + "-new.docx")),
                       document=copy.deepcopy(docx_document))
    doc_word.replace_matches_newpage(max_count)
    doc_word.save()

    show_info("Success", "Completed mail merge successfully")


def main_seperate_doc(xlsx_path, docx_path, pptx_path, out_path):
    wipe_out_folder(out_path)
    active_sheet = init_sheet(xlsx_path)
    headers = get_headers(active_sheet)
    values = {}
    for i, items in enumerate(headers):
        values.update({items: get_values(active_sheet, (i + 1))})

    max_count = 0
    for value in values.values():
        if len(value) > max_count:
            max_count = len(value)

    docx_document = init_document(docx_path)
    ppt_document = init_ppt(pptx_path)

    for i, items in enumerate(sort_dicts(values)):
        doc_docx = MergeDocument(matches=items, path=(os.path.join(out_path, str(os.path.basename(docx_path).split(".")[0]) + str(i) + ".docx")),
                                 document=copy.deepcopy(docx_document))
        doc_ppt = MergePPT(matches=items, path=(os.path.join(out_path, str(os.path.basename(pptx_path).split(".")[0]) + str(i) + ".pptx")),
                           presentation=copy.deepcopy(ppt_document))

        doc_docx.replace_matches()
        doc_ppt.replace_matches()
        doc_docx.save()
        doc_ppt.save()

    show_info("Success", "Completed mail merge successfully")


def save_ppt_as_pdf(out_path):
    shutil.copy(os.path.join(os.path.dirname((os.path.abspath(__file__))), "ppt2pdf.ps1"), out_path)
    import subprocess
    print(os.path.join(out_path, "ppt2pdf.ps1"))
    cmd = [r"powershell", os.path.join(out_path, "ppt2pdf.ps1").replace(" ", "' '")]
    subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE).communicate()
    os.remove(os.path.join(out_path, "ppt2pdf.ps1"))
    show_info("Success", "Completed conversion successfully")
