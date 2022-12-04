from tkinter import Entry
from tkinter import filedialog as fd
from tkinter import Frame
from tkinter import Label
from tkinter import N
from tkinter import S
from tkinter import StringVar
from tkinter import Text
from tkinter import ttk
from tkinter import W
import tkinter as tk
from typing import Optional

from document_distributor.document_distributor import DocumentEmailName
from document_distributor.document_distributor import SUPPORTED_FILE_TYPES


class DocumentDetailEntry(Frame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._document_file_types = StringVar()
        self.label_document_file_type = Label(self, text="Document Types")
        self.label_document_file_type.grid(row=0, column=0, sticky=W)
        self.entry_document_file_type = Entry(self, textvariable=self._document_file_types)
        self.entry_document_file_type.grid(row=0, column=1, sticky=W)

    @property
    def document_file_types(self) -> list[str]:
        return self._document_file_types.get().strip(" ").split(" ")


class NameEmailDetailEntries(Frame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._first_name_column = StringVar()
        self.label_first_name_column = Label(self, text="First Name Column*")
        self.label_first_name_column.grid(row=0, column=0, sticky=W)
        self.entry_first_name_column = Entry(self, textvariable=self._first_name_column)
        self.entry_first_name_column.grid(row=0, column=1, sticky=W)
        self._last_name_column = StringVar()
        self.label_last_name_column = Label(self, text="Last Name Column*")
        self.label_last_name_column.grid(row=1, column=0, sticky=W)
        self.entry_last_name_column = Entry(self, textvariable=self._last_name_column)
        self.entry_last_name_column.grid(row=1, column=1, sticky=W)
        self._email_column = StringVar()
        self.label_email_column = Label(self, text="Email Column*")
        self.label_email_column.grid(row=2, column=0, sticky=W)
        self.entry_email_column = Entry(self, textvariable=self._email_column)
        self.entry_email_column.grid(row=2, column=1, sticky=W)
        self._sheet_name = StringVar()
        self.label_sheet_name = Label(self, text="Sheet Name")
        self.label_sheet_name.grid(row=3, column=0, sticky=W)
        self.entry_sheet_name = Entry(self, textvariable=self._sheet_name)
        self.entry_sheet_name.grid(row=3, column=1, sticky=W)
        self._start_row_for_names = StringVar()
        self.label_start_row_for_names = Label(self, text="Start Row Names")
        self.label_start_row_for_names.grid(row=4, column=0, sticky=W)
        self.entry_start_row_for_names = Entry(self, textvariable=self._start_row_for_names)
        self.entry_start_row_for_names.grid(row=4, column=1, sticky=W)
        self._stop_row_for_names = StringVar()
        self.label_stop_row_for_names = Label(self, text="End Row Names")
        self.label_stop_row_for_names.grid(row=5, column=0, sticky=W)
        self.entry_stop_row_for_names = Entry(self, textvariable=self._stop_row_for_names)
        self.entry_stop_row_for_names.grid(row=5, column=1, sticky=W)

    @property
    def first_name_column(self) -> str:
        return self._first_name_column.get().upper().strip(" ")

    @property
    def email_column(self) -> str:
        return self._email_column.get().upper().strip(" ")

    @property
    def last_name_column(self) -> str:
        return self._last_name_column.get().upper().strip(" ")

    @property
    def sheet_name(self) -> str:
        return self._sheet_name.get().strip(" ")

    @property
    def start_row_for_names(self) -> str:
        res = self._start_row_for_names.get().strip(" ")
        return res if res else "1"

    @property
    def stop_row_for_names(self) -> Optional[int]:
        res = self._stop_row_for_names.get().strip(" ")
        return res if res else None


class EmailMessage(Frame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._email_subject = StringVar()
        self.label_email_subject = Label(self, text="Email Subject*")
        self.label_email_subject.grid(row=0, column=0, sticky=W)
        self.entry_email_subject = Entry(self, textvariable=self._email_subject)
        self.entry_email_subject.grid(row=1, column=0, sticky=W)
        self._email_message = StringVar()
        self.label_email_message = Label(self, text="Email Text*")
        self.label_email_message.grid(row=2, column=0, sticky=W)
        self.entry_email_message = Text(self)
        self.entry_email_message.grid(row=3, column=0, sticky=W)


class FolderSelect(Frame):

    def __init__(self, parent=None, folder_description="", **kwwargs):
        super().__init__(master=parent, **kwwargs)
        self._folder_path = StringVar()
        self.label_name = Label(self, text=folder_description)
        self.label_name.grid(row=0, column=0, sticky=W)
        self.entry_path = Entry(self, textvariable=self._folder_path)
        self.entry_path.grid(row=1, column=0, sticky=W)
        self.button_find = ttk.Button(self, text="Browse Folder", command=self.set_folder_path)
        self.button_find.grid(row=1, column=1, sticky=W)

    def set_folder_path(self):
        folder_selected = fd.askdirectory()
        self._folder_path.set(folder_selected)

    @property
    def folder_path(self):
        return self._folder_path.get()


class FileSelect(Frame):

    def __init__(self, parent=None, file_description="", filetypes=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self.filetypes = filetypes
        self._file_path = StringVar()
        self.label_name = Label(self, text=file_description)
        self.label_name.grid(row=0, column=0, sticky=W)
        self.entry_path = Entry(self, textvariable=self._file_path)
        self.entry_path.grid(row=1, column=0, sticky=W)
        self.button_find = ttk.Button(self, text="Browse File", command=self.set_file_path)
        self.button_find.grid(row=1, column=1, sticky=W)

    def set_file_path(self):
        file_selected = fd.askopenfilename(filetypes=self.filetypes)
        self._file_path.set(file_selected)

    @property
    def file_path(self):
        return self._file_path.get()


def gui_root() -> tk.Tk:
    root = tk.Tk()
    root.title("File Distributor")
    root.resizable(width=True, height=True)
    root.geometry("300x150")
    document_dir_select = FolderSelect(root, "Document Directory*")
    document_dir_select.grid(row=0, sticky=W)
    document_detail_entries = DocumentDetailEntry(root)
    document_detail_entries.grid(row=1, sticky=W)
    name_email_file_select = FileSelect(root,
                                        "Name and Email File*",
                                        filetypes=(*SUPPORTED_FILE_TYPES, ("All files", "*.*")))
    name_email_file_select.grid(row=2, sticky=W)
    name_email_details_entries = NameEmailDetailEntries(root)
    name_email_details_entries.grid(row=3, sticky=W)
    email_message_entries = EmailMessage(root)
    email_message_entries.grid(row=0, column=1, rowspan=4, sticky=N + S)
    start_button = ttk.Button(root,
                              text="Start",
                              command=lambda: process(root, document_dir_select, document_detail_entries,
                                                      name_email_file_select, name_email_details_entries))
    start_button.grid(row=4)

    return root


def process(root: tk.Tk, document_dir_select: FolderSelect, document_detail: DocumentDetailEntry,
            name_email_file_select: FileSelect, name_email_details: NameEmailDetailEntries):
    start_row_for_names = int(name_email_details.start_row_for_names)
    stop_row_for_names = int(
        name_email_details.stop_row_for_names) if name_email_details.stop_row_for_names is not None else None
    DocumentEmailName.from_data_paths(
        document_folder_path=document_dir_select.folder_path,
        excel_file_path=name_email_file_select.file_path,
        first_name_column=name_email_details.first_name_column,
        email_column=name_email_details.email_column,
        document_file_type=document_detail.document_file_types,
        last_name_column=name_email_details.last_name_column,
        sheet_name=name_email_details.sheet_name,
        start_row_for_names=start_row_for_names,
        stop_row_for_names=stop_row_for_names,
    )


def app():
    gui_root().mainloop()
