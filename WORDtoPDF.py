
import tkinter as tk
from tkinter import messagebox
import queue
import threading
from tkinter import ttk  # Normal Tkinter.* widgets are not themed!
from ttkthemes import ThemedTk
import os
from enum import Enum
from tkinter import filedialog as fd
from win32com import client
import comtypes.client
import time
from pdf2docx import parse
import pandas as pd
import tabula

cfrm = None


class SettingsStatus(Enum):
    # valid settings
    valid_settings = 1

    # invalid settings
    directory_not_selected = 2
    directory_does_not_exist = 3
    no_files_found = 4


def xl2pdf(input_file, output_file):
    app = client.DispatchEx("Excel.Application")
    app.Interactive = True
    app.Visible = False

    try:
        Workbook = app.Workbooks.Open(input_file)
        Workbook.ExportAsFixedFormat(0, output_file)

    except Exception as e:
        messagebox.showerror("Export Error", e)
        return False
#     finally:
#         Workbook.Close()
#         app.Quit()
    return True


def pdf2xl(input_file, output_file):
    try:
        list_of_dfs = tabula.read_pdf(input_file, pages='all')
        result_df = pd.concat(list_of_dfs)
        # save the above data frame in a a excel file
        result_df.to_excel(output_file)
        # tables = camelot.read_pdf(input_file,pages='all',line_scale=40)
        # tables.export(output_file, f='excel')
    except Exception as e:
        messagebox.showerror("Export Error", e)
        return False
    return True


def doc2pdf(input_file, output_file):
    wdFormatPDF = 17
    word = client.DispatchEx('Word.Application')
    word.Visible = False
    try:

        doc = word.Documents.Open(input_file)
        doc.SaveAs(output_file, FileFormat=wdFormatPDF)

    except Exception as e:
        messagebox.showerror("Export Error", e)
        return False
    finally:
        doc.Close()
        word.Quit()
    return True


def pdf2doc(input_file, output_file):
    try:
        parse(input_file, output_file)

    except Exception as e:
        messagebox.showerror("Export Error", e)
        return False
    return True


def get_settings_status(dirx, files):
    selected_directory = dirx

    # Check if a directory is selected
    if selected_directory == "Select Directory":
        return SettingsStatus.directory_not_selected

    # Check if the selected directory exists
    if not os.path.isdir(selected_directory):
        return SettingsStatus.directory_does_not_exist

    if len(files) == 0:
        return SettingsStatus.no_files_found

    return SettingsStatus.valid_settings


def confirm_settings(dirx, fl):
    confirm_msg = "You are about to export with these settings:"
    confirm_msg += "\nDirectory: " + dirx
    # confirm_msg += "\nThumbnail images to: " + export_properties.get()
    # confirm_msg += "x" + export_properties.get() + " px"
    confirm_msg += "\nSave as: " + fl
    confirm_msg += "\n\nAre you sure you want to continue?"

    return messagebox.askyesno("Export confirmation", confirm_msg)


def export(dirx):
    Ffiles = get_files(dirx, objs.fltyp1, objs.fltyp2)
    settings_status = get_settings_status(dirx, Ffiles)
    print(Ffiles)

    if settings_status is not SettingsStatus.valid_settings:
        # Invalid settings, display an error message
        error_messages = {
            SettingsStatus.directory_not_selected: [
                "Invalid directory",
                "You have to select a directory first"
            ],
            SettingsStatus.directory_does_not_exist: [
                "Invalid directory",
                "The directory \"" + dirx + "\" does not exist"
            ],
            SettingsStatus.no_files_found: [
                "Check Directory for Files",
                "No Valid Files Found"
            ],
        }
        messagebox.showerror(*error_messages[settings_status])
    else:
        # Valid settings, confirm settings with the user and export
        if confirm_settings(dirx, objs.exptyp):
            # call `ImgEdit.export_all_in_dir` as the target of a new thread
            # and put the final result in a `Queue`
            # conv.check(Ffiles)
            q = queue.Queue()
            my_thread = threading.Thread(target=conv.check, args=(Ffiles, q))
            my_thread.start()


class conv():
    def check(inp_files, q):
        if (objs.convtype == "EXCEL TO PDF"):
            x = xl2pdf
        elif (objs.convtype == "PDF TO EXCEL"):
            x = pdf2xl
        elif (objs.convtype == "WORD TO PDF"):
            x = doc2pdf
        elif (objs.convtype == "PDF TO WORD"):
            x = pdf2doc

        conv.num_of_exported = 0

        conv.num_to_export = len(inp_files)
        topmenu.progress_bar['maximum'] = conv.num_to_export

        # loop through the images list to open, resize and save them
        conv.all_exported_successfully = True
        for fl in inp_files:
            # export and check if everything went okay

            input_file = r"{}/{}".format(fl["path"], fl["name"])
            input_file = os.path.abspath(input_file)
            # input_file=os.path.join(fl["path"],fl["name"])
            output_file = fl["path"] + "/" + fl["name"].split('.')[0] + objs.exptyp
            output_file = os.path.abspath(output_file)
            exported_successfully = x(input_file, output_file)
            if not exported_successfully:
                conv.all_exported_successfully = False

            conv.num_of_exported += 1
            conv.updt_pb(conv.num_of_exported, q)

        conv.exported(conv.all_exported_successfully, conv.num_of_exported, conv.num_to_export)

    def updt_pb(num_of_exported, q):
        topmenu.progress_bar['value'] = conv.num_of_exported
        # topmenu.progress_bar.grid(row=3, column=0,columnspan=3,sticky="we", padx=10,pady=10)
        time.sleep(.1)
        return

    def exported(result, num_of_exported, num_to_export):

        if result:
            messagebox.showinfo("Exports completed",
                                ("{} Files were exported successfully").format(num_of_exported))
        else:
            messagebox.showwarning("Exports failed",
                                   ("One or more Files failed to export").format(num_to_export - num_of_exported))


def is_typ(file, typ1, typ2):
    filex = str(file)
    return (filex.endswith(typ1) or filex.endswith(typ2))


def get_files(dirx, typ1, typ2):
    # loop through all the files in the given directory and store the
    # path and the filename of image files in a list
    filesx = []
    for path, subdirs, files in os.walk(dirx):
        for name in files:
            if is_typ(name, typ1, typ2):
                filesx.append({"path": path, "name": name})

    return filesx


class topmenu():

    def create_widgets(cfrm):
        s = ttk.Style()
        s.configure('medp.TButton', font=("Segoe UI", 12))
        cfrm.rowconfigure([0, 1, 2, 3, 4], weight=1)
        cfrm.columnconfigure([0, 1, 2, 3], weight=1)
        cfrm.grid(row=0, column=0, sticky="nswe")
        ttk.Label(cfrm, text=objs.convtype, font=font_big).grid(row=0, column=0, columnspan=4)

        # Browse
        ttk.Label(cfrm, text="Folder :", font=font_medium).grid(row=2, column=0, padx=10, pady=10)

        selected_directory = tk.StringVar(cfrm, value="Select Directory")
        ttk.Entry(cfrm, textvariable=selected_directory).grid(row=2, column=1, columnspan=2, padx=10, pady=10,
                                                              sticky="nswe")
        ttk.Button(cfrm, text="Browse", style='medp.TButton', command=lambda: selected_directory.set(
            fd.askdirectory(initialdir=r"C:\Users\356285\Desktop", title="Select a File", ))).grid(row=2, column=3,
                                                                                                  padx=10, pady=10,
                                                                                                  sticky="nswe")
        # Progress
        ttk.Label(cfrm, text="Progress :", font=font_medium).grid(row=3, column=0, padx=10, pady=10)
        topmenu.progress_bar = ttk.Progressbar(cfrm, orient='horizontal', mode='determinate', length=400)
        topmenu.progress_bar['value'] = 0
        topmenu.progress_bar.grid(column=1, columnspan=2, pady=10, sticky='ew', row=3, padx=10)
        # Export
        ttk.Button(cfrm, text="Export", style='medp.TButton', command=lambda: export(selected_directory.get())).grid(
            row=3, column=3, padx=10, pady=10, sticky="nswe")

        # Copyright

        ttk.Button(cfrm, text="Back", style='medp.TButton', command=lambda: topmenu.back(cfrm), width=20).grid(row=4,
                                                                                                               column=0,
                                                                                                               columnspan=2,
                                                                                                               padx=10,
                                                                                                               pady=10,
                                                                                                               sticky='ns')
        ttk.Button(cfrm, text="Quit", style='medp.TButton', command=root.destroy).grid(row=4, column=2, columnspan=2,
                                                                                       sticky='nswe', padx=10, pady=10)
        ttk.Label(cfrm, text="Developed by praviii", font=font_small).grid(row=5, column=0, columnspan=4, padx=10,
                                                                                pady=10, sticky="nswe")

    def menu():
        s = ttk.Style()
        s.configure('medx.TButton', font=("Segoe UI", 12))

        sel_frame = ttk.Frame(root)
        sel_frame.grid(row=0, column=0, sticky="nswe")
        sel_frame.rowconfigure([0, 1, 2, 3], weight=1)
        sel_frame.columnconfigure([0, 1, 2, 3], weight=1)
        ttk.Label(sel_frame, text="Select File Convertor", font=font_big).grid(row=0, column=0, columnspan=4)
        ttk.Button(sel_frame, text="EXCEL to PDF", command=objs.x2p, style='medx.TButton').grid(row=1, column=0,
                                                                                                padx=10, pady=10,
                                                                                                sticky="nswe")
        ttk.Button(sel_frame, text="PDF to EXCEL", command=objs.p2x, style='medx.TButton').grid(row=1, column=1,
                                                                                                padx=10, pady=10,
                                                                                                sticky="nswe")
        ttk.Button(sel_frame, text="PDF to WORD", command=objs.p2w, style='medx.TButton').grid(row=1, column=3, padx=10,
                                                                                               pady=10, sticky="nswe")
        ttk.Button(sel_frame, text="WORD to PDF", command=objs.w2p, style='medx.TButton').grid(row=1, column=2, padx=10,
                                                                                               pady=10, sticky="nswe")
        ttk.Button(sel_frame, text="Quit", command=root.destroy, style='medx.TButton').grid(row=2, column=0,
                                                                                            columnspan=4, padx=10,
                                                                                            pady=20, sticky="we")
        ttk.Label(sel_frame, text="Developed for Anjali", font=font_small).grid(row=5, column=0, sticky="nswe",
                                                                                     pady=10, padx=10)

    def back(x):
        x.destroy()
        topmenu.menu()


class objs():
    fltyp = None
    exptyp = None
    convtype = None

    def x2p():
        x2pf = ttk.Frame(root)
        x2pf.grid(row=0, column=0, sticky="nswe")
        objs.fltyp1 = '.xlsx'
        objs.fltyp2 = '.xls'
        objs.exptyp = '.pdf'
        objs.convtype = "EXCEL TO PDF"
        topmenu.create_widgets(x2pf)

    def p2x():
        x2pf = ttk.Frame(root)
        x2pf.grid(row=0, column=0, sticky="nswe")
        objs.fltyp1 = '.pdf'
        objs.fltyp2 = '.pdf'
        objs.exptyp = '.xlsx'
        objs.convtype = "PDF TO EXCEL"
        topmenu.create_widgets(x2pf)

    def w2p():
        x2pf = ttk.Frame(root)
        x2pf.grid(row=0, column=0, sticky="nswe")
        objs.fltyp1 = '.docx'
        objs.fltyp2 = '.doc'
        objs.exptyp = '.pdf'
        objs.convtype = "WORD TO PDF"
        topmenu.create_widgets(x2pf)

    def p2w():
        x2pf = ttk.Frame(root)
        x2pf.grid(row=0, column=0, sticky="nswe")
        objs.fltyp1 = '.pdf'
        objs.fltyp2 = '.pdf'
        objs.exptyp = '.docx'
        objs.convtype = "PDF TO WORD"
        topmenu.create_widgets(x2pf)


font_big = ("Segoe UI", 24)
font_medium = ("Segoe UI", 12)
font_small = ("Segoe UI", 8)

root = ThemedTk(theme="arc")
root.title("PDF in idhill change cheyyitaaa....")
root.configure(bg="White")
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)
root.geometry("680x300")

topmenu.menu()

root.mainloop()


