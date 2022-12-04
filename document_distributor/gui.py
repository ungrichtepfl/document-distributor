from tkinter import Checkbutton
from tkinter import END
from tkinter import Entry
from tkinter import filedialog as fd
from tkinter import Frame
from tkinter import IntVar
from tkinter import Label
from tkinter import N
from tkinter import S
from tkinter import StringVar
from tkinter import Text
from tkinter import Toplevel
from tkinter import ttk
from tkinter import W
import tkinter as tk
from typing import Optional

from document_distributor.document_distributor import DocumentEmailName
from document_distributor.document_distributor import EmailConfig
from document_distributor.document_distributor import MainConfig
from document_distributor.document_distributor import SUPPORTED_FILE_TYPES


class DocumentConfigFrame(Frame):

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


class NameEmailConfigFrame(Frame):

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


class EmailServerConfigFrame(Frame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._sender_email = StringVar()
        self.label_sender_email = Label(self, text="Sender Email*")
        self.label_sender_email.grid(row=0, column=0, sticky=W)
        self.entry_sender_email = Entry(self, textvariable=self._sender_email)
        self.entry_sender_email.grid(row=0, column=1, sticky=W)
        self._smtp_server = StringVar()
        self.label_smtp_server = Label(self, text="SMTP Server*")
        self.label_smtp_server.grid(row=1, column=0, sticky=W)
        self.entry_smtp_server = Entry(self, textvariable=self._smtp_server)
        self.entry_smtp_server.grid(row=1, column=1, sticky=W)
        self._smtp_port = IntVar(value=587)
        self.label_smtp_port = Label(self, text="SMTP Port*")
        self.label_smtp_port.grid(row=2, column=0, sticky=W)
        self.entry_smtp_port = Entry(self, textvariable=self._smtp_port)
        self.entry_smtp_port.grid(row=2, column=1, sticky=W)
        self._use_starttls = IntVar(value=1)
        self.label_use_starttls = Label(self, text="Use Starttls*")
        self.label_use_starttls.grid(row=3, column=0, sticky=W)
        self.entry_use_starttls = Checkbutton(self, variable=self._use_starttls)
        self.entry_use_starttls.grid(row=3, column=1, sticky=W)
        self._username = StringVar()
        self.label_username = Label(self, text="Username*")
        self.label_username.grid(row=4, column=0, sticky=W)
        self.entry_username = Entry(self, textvariable=self._username)
        self.entry_username.grid(row=4, column=1, sticky=W)
        self._password = StringVar()
        self.label_password = Label(self, text="Password*")
        self.label_password.grid(row=5, column=0, sticky=W)
        self.entry_password = Entry(self, textvariable=self._password, show="*")
        self.entry_password.grid(row=5, column=1, sticky=W)

    @property
    def sender_email(self) -> str:
        return self._sender_email.get()

    @property
    def smtp_server(self) -> str:
        return self._smtp_server.get()

    @property
    def smtp_port(self) -> int:
        return self._smtp_port.get()

    @property
    def use_starttls(self) -> bool:
        return bool(self._use_starttls.get())

    @property
    def username(self) -> str:
        return self._username.get()

    @property
    def password(self) -> str:
        return self._password.get()


class EmailMessageConfigFrame(Frame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._email_subject = StringVar()
        self.label_email_subject = Label(self, text="Email Subject*")
        self.label_email_subject.grid(row=0, column=0, sticky=W)
        self.entry_email_subject = Entry(self, textvariable=self._email_subject)
        self.entry_email_subject.grid(row=1, column=0, sticky=W)
        self.label_email_message = Label(self, text="Email Text*")
        self.label_email_message.grid(row=2, column=0, sticky=W)
        self.text_email_message = Text(self)
        self.text_email_message.grid(row=3, column=0, sticky=W)

    @property
    def email_subject(self) -> str:
        return self._email_subject.get()

    @property
    def email_message(self) -> str:
        return self.text_email_message.get("1.0", END)


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


def set_email_config(email_config: EmailConfig):
    pop_up = Toplevel()
    email_server_config_frame = EmailServerConfigFrame(pop_up)
    email_server_config_frame.grid(row=0)

    def set_and_destroy():
        email_config.password = email_server_config_frame.password
        email_config.username = email_server_config_frame.username
        email_config.port = email_server_config_frame.smtp_port
        email_config.smtp_server = email_server_config_frame.smtp_server
        email_config.use_starttls = email_server_config_frame.use_starttls
        pop_up.destroy()

    ok_button = ttk.Button(pop_up, text="Ok", command=set_and_destroy)
    ok_button.grid(row=1)
    print(f"In popup {email_config}")


def gui_root() -> tk.Tk:
    root = tk.Tk()
    root.title("File Distributor")
    root.resizable(width=True, height=True)
    root.geometry("300x150")
    document_dir_select = FolderSelect(root, "Document Directory*")
    document_dir_select.grid(row=0, sticky=W)
    document_config_frame = DocumentConfigFrame(root)
    document_config_frame.grid(row=1, sticky=W)
    name_email_file_select = FileSelect(root,
                                        "Name and Email File*",
                                        filetypes=(*SUPPORTED_FILE_TYPES, ("All files", "*.*")))
    name_email_file_select.grid(row=2, sticky=W)
    name_email_config_frame = NameEmailConfigFrame(root)
    name_email_config_frame.grid(row=3, sticky=W)
    email_message_config_frame = EmailMessageConfigFrame(root)
    email_message_config_frame.grid(row=0, column=1, rowspan=4, sticky=N + S)
    email_config = EmailConfig("test_email", "test_server")
    email_config_button = ttk.Button(root, text="Email Config", command=lambda: set_email_config(email_config))
    email_config_button.grid(row=4, column=1)
    print(f"In main {email_config}")
    start_button = ttk.Button(root,
                              text="Start",
                              command=lambda: process(root, document_dir_select, document_config_frame,
                                                      name_email_file_select, name_email_config_frame))
    start_button.grid(row=4, column=0)

    return root


def process(root: tk.Tk, document_dir_select: FolderSelect, document_config_frame: DocumentConfigFrame,
            name_email_file_select: FileSelect, name_email_config_frame: NameEmailConfigFrame):
    start_row_for_names = int(name_email_config_frame.start_row_for_names)
    stop_row_for_names = int(
        name_email_config_frame.stop_row_for_names) if name_email_config_frame.stop_row_for_names is not None else None
    DocumentEmailName.from_data_paths(
        document_folder_path=document_dir_select.folder_path,
        excel_file_path=name_email_file_select.file_path,
        first_name_column=name_email_config_frame.first_name_column,
        email_column=name_email_config_frame.email_column,
        document_file_type=document_config_frame.document_file_types,
        last_name_column=name_email_config_frame.last_name_column,
        sheet_name=name_email_config_frame.sheet_name,
        start_row_for_names=start_row_for_names,
        stop_row_for_names=stop_row_for_names,
    )


def app():
    gui_root().mainloop()
