import os
import threading
from typing import Optional

from customtkinter import BOTH
from customtkinter import CTkButton
from customtkinter import CTkCheckBox
from customtkinter import CTkEntry
from customtkinter import CTkFont
from customtkinter import CTkFrame
from customtkinter import CTkLabel
from customtkinter import CTkProgressBar
from customtkinter import CTkTextbox
from customtkinter import CTkToplevel
from customtkinter import DISABLED
from customtkinter import E
from customtkinter import END
from customtkinter import filedialog as fd
from customtkinter import IntVar
from customtkinter import LEFT
from customtkinter import N
from customtkinter import NORMAL
from customtkinter import S
from customtkinter import StringVar
from customtkinter import ThemeManager
from customtkinter import W
import customtkinter as ctk

from document_distributor.document_distributor import DocumentEmailName
from document_distributor.document_distributor import dump_config
from document_distributor.document_distributor import EmailAddress
from document_distributor.document_distributor import EmailConfig
from document_distributor.document_distributor import FilePath
from document_distributor.document_distributor import load_config
from document_distributor.document_distributor import send_emails
from document_distributor.document_distributor import SUPPORTED_FILE_TYPES

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("dark-blue")
FG_COLOR_ROOT = ThemeManager.theme["CTk"]["fg_color"]
PADDING = {"padx": 7, "pady": 7}


class DocumentConfigFrame(CTkFrame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._document_file_types = StringVar()
        self.label_document_file_type = CTkLabel(self, text="Document Types")
        self.label_document_file_type.grid(row=0, column=0, sticky=W, padx=5)
        self.entry_document_file_type = CTkEntry(self, textvariable=self._document_file_types)
        self.entry_document_file_type.grid(row=0, column=1, sticky=W)

    @property
    def document_file_types(self) -> list[str]:
        return self._document_file_types.get().strip(" ").split(" ")

    def set_document_file_types(self, val: Optional[list[str]]):
        return self._document_file_types.set(" ".join(val) if val is not None else "")


class NameEmailConfigFrame(CTkFrame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._first_name_column = StringVar()
        self.label_first_name_column = CTkLabel(self, text="First Name Column*")
        self.label_first_name_column.grid(row=0, column=0, sticky=W, padx=5)
        self.entry_first_name_column = CTkEntry(self, textvariable=self._first_name_column)
        self.entry_first_name_column.grid(row=0, column=1, sticky=W)
        self._last_name_column = StringVar()
        self.label_last_name_column = CTkLabel(self, text="Last Name Column*")
        self.label_last_name_column.grid(row=1, column=0, sticky=W, padx=5)
        self.entry_last_name_column = CTkEntry(self, textvariable=self._last_name_column)
        self.entry_last_name_column.grid(row=1, column=1, sticky=W)
        self._email_column = StringVar()
        self.label_email_column = CTkLabel(self, text="Email Column*")
        self.label_email_column.grid(row=2, column=0, sticky=W, padx=5)
        self.entry_email_column = CTkEntry(self, textvariable=self._email_column)
        self.entry_email_column.grid(row=2, column=1, sticky=W)
        self._sheet_name = StringVar()
        self.label_sheet_name = CTkLabel(self, text="Sheet Name")
        self.label_sheet_name.grid(row=3, column=0, sticky=W, padx=5)
        self.entry_sheet_name = CTkEntry(self, textvariable=self._sheet_name)
        self.entry_sheet_name.grid(row=3, column=1, sticky=W)
        self._start_row_for_names = StringVar()
        self.label_start_row_for_names = CTkLabel(self, text="Start Row Names")
        self.label_start_row_for_names.grid(row=4, column=0, sticky=W, padx=5)
        self.entry_start_row_for_names = CTkEntry(self, textvariable=self._start_row_for_names)
        self.entry_start_row_for_names.grid(row=4, column=1, sticky=W)
        self._stop_row_for_names = StringVar()
        self.label_stop_row_for_names = CTkLabel(self, text="End Row Names")
        self.label_stop_row_for_names.grid(row=5, column=0, sticky=W, padx=5)
        self.entry_stop_row_for_names = CTkEntry(self, textvariable=self._stop_row_for_names)
        self.entry_stop_row_for_names.grid(row=5, column=1, sticky=W)

    @property
    def first_name_column(self) -> str:
        return self._first_name_column.get().upper().strip(" ")

    def set_first_name_column(self, val: str):
        self._first_name_column.set(val)

    @property
    def email_column(self) -> str:
        return self._email_column.get().upper().strip(" ")

    def set_email_column(self, val: str):
        self._email_column.set(val)

    @property
    def last_name_column(self) -> Optional[str]:
        res = self._last_name_column.get().upper().strip(" ")
        return res if res else None

    def set_last_name_column(self, val: Optional[str]):
        self._last_name_column.set(val if val is not None else "")

    @property
    def sheet_name(self) -> Optional[str]:
        res = self._sheet_name.get().strip(" ")
        return res if res else None

    def set_sheet_name(self, val: Optional[str]):
        self._sheet_name.set(val if val is not None else "")

    @property
    def start_row_for_names(self) -> int:
        res = self._start_row_for_names.get().strip(" ")
        return int(res) if res else 1

    def set_start_row_for_names(self, val: int):
        self._start_row_for_names.set(str(val))

    @property
    def stop_row_for_names(self) -> Optional[int]:
        res = self._stop_row_for_names.get().strip(" ")
        return int(res) if res else None

    def set_stop_row_for_names(self, val: Optional[int]):
        self._stop_row_for_names.set(str(val) if val is not None else "")


class EmailServerConfigFrame(CTkFrame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._sender_email = StringVar()
        self.label_sender_email = CTkLabel(self, text="Sender Email*")
        self.label_sender_email.grid(row=0, column=0, sticky=W, padx=5)
        self.entry_sender_email = CTkEntry(self, textvariable=self._sender_email, width=240)
        self.entry_sender_email.grid(row=0, column=1, sticky=W)
        self._smtp_server = StringVar()
        self.label_smtp_server = CTkLabel(self, text="SMTP Server*")
        self.label_smtp_server.grid(row=1, column=0, sticky=W, padx=5)
        self.entry_smtp_server = CTkEntry(self, textvariable=self._smtp_server, width=240)
        self.entry_smtp_server.grid(row=1, column=1, sticky=W)
        self._smtp_port = IntVar(value=587)
        self.label_smtp_port = CTkLabel(self, text="SMTP Port*")
        self.label_smtp_port.grid(row=2, column=0, sticky=W, padx=5)
        self.entry_smtp_port = CTkEntry(self, textvariable=self._smtp_port, width=240)
        self.entry_smtp_port.grid(row=2, column=1, sticky=W)
        self._use_starttls = IntVar(value=1)
        self.label_use_starttls = CTkLabel(self, text="Use Starttls*")
        self.label_use_starttls.grid(row=3, column=0, sticky=W, padx=5)
        self.entry_use_starttls = CTkCheckBox(self, variable=self._use_starttls, text="")
        self.entry_use_starttls.grid(row=3, column=1, sticky=W)
        self._username = StringVar()
        self.label_username = CTkLabel(self, text="Username*")
        self.label_username.grid(row=4, column=0, sticky=W, padx=5)
        self.entry_username = CTkEntry(self, textvariable=self._username, width=240)
        self.entry_username.grid(row=4, column=1, sticky=W)
        self._password = StringVar()
        self.label_password = CTkLabel(self, text="Password*")
        self.label_password.grid(row=5, column=0, sticky=W, padx=5)
        self.entry_password = CTkEntry(self, textvariable=self._password, show="*", width=240)
        self.entry_password.grid(row=5, column=1, sticky=W)

    @property
    def sender_email(self) -> str:
        return self._sender_email.get()

    def set_sender_email(self, val: str):
        self._sender_email.set(val)

    @property
    def smtp_server(self) -> str:
        return self._smtp_server.get()

    def set_smtp_server(self, val: str):
        self._smtp_server.set(val)

    @property
    def smtp_port(self) -> int:
        return self._smtp_port.get()

    def set_smtp_port(self, val: int):
        self._smtp_port.set(val)

    @property
    def use_starttls(self) -> bool:
        return bool(self._use_starttls.get())

    def set_use_starttls(self, val: bool):
        self._use_starttls.set(int(val))

    @property
    def username(self) -> str:
        return self._username.get()

    def set_username(self, val: Optional[str]):
        return self._username.set(val if val is not None else "")

    @property
    def password(self) -> str:
        return self._password.get()

    def set_password(self, val: Optional[str]):
        self._password.set(val if val is not None else "")


class EmailMessageConfigFrame(CTkFrame):

    def __init__(self, parent=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self._email_subject = StringVar()
        self.label_email_subject = CTkLabel(self, text="Email Subject*")
        self.label_email_subject.grid(row=0, sticky=W)
        self.entry_email_subject = CTkEntry(self, textvariable=self._email_subject)
        self.entry_email_subject.grid(row=1, column=0, sticky=W + E)
        self.label_email_message = CTkLabel(self, text="Email Text*")
        self.label_email_message.grid(row=2, column=0, sticky=W)
        self.message_frame = CTkFrame(self)
        self.message_frame.grid(row=3, column=0, sticky=W + E)

        font = CTkFont(size=14)
        self.text_email_message = CTkTextbox(self.message_frame, width=800, height=400, font=font, wrap="none")
        self.text_email_message.pack(side=LEFT, fill=BOTH, expand=True)

    @property
    def email_subject(self) -> str:
        return self._email_subject.get()

    def set_email_subject(self, val: str):
        self._email_subject.set(val)

    @property
    def email_message(self) -> str:
        return self.text_email_message.get(1.0, END)

    def set_email_message(self, val: str):
        self.text_email_message.delete("1.0", END)
        self.text_email_message.insert(END, val)


class FolderSelect(CTkFrame):

    def __init__(self, parent=None, folder_description="", **kwwargs):
        super().__init__(master=parent, **kwwargs)
        self._folder_path = StringVar()
        self.label_name = CTkLabel(self, text=folder_description)
        self.label_name.grid(row=0, column=0, sticky=W)
        self.entry_path = CTkEntry(self, textvariable=self._folder_path, width=200)
        self.entry_path.grid(row=1, column=0, sticky=W)
        self.entry_path.xview_moveto(1)
        self.button_find = CTkButton(self, text="Browse Folder", command=self.set_folder_path_button)
        self.button_find.grid(row=1, column=1, sticky=W)

    def set_folder_path_button(self):
        folder_selected = fd.askdirectory()
        self._folder_path.set(folder_selected)
        self.entry_path.xview_moveto(1)

    def set_folder_path(self, val: str):
        self._folder_path.set(val)
        self.entry_path.xview_moveto(1)

    @property
    def folder_path(self):
        return self._folder_path.get()


class FileSelect(CTkFrame):

    def __init__(self, parent=None, file_description="", filetypes=None, **kwargs):
        super().__init__(master=parent, **kwargs)
        self.filetypes = filetypes
        self._file_path = StringVar()
        self.label_name = CTkLabel(self, text=file_description)
        self.label_name.grid(row=0, column=0, sticky=W)
        self.entry_path = CTkEntry(self, textvariable=self._file_path, width=200)
        self.entry_path.grid(row=1, column=0, sticky=W)
        self.entry_path.xview_moveto(1)
        self.button_find = CTkButton(self, text="Browse File", command=self.set_file_path_button)
        self.button_find.grid(row=1, column=1, sticky=W)

    def set_file_path_button(self):
        file_selected = fd.askopenfilename(filetypes=self.filetypes)
        self._file_path.set(file_selected)
        self.entry_path.xview_moveto(1)

    def set_file_path(self, val: str):
        self._file_path.set(val)
        self.entry_path.xview_moveto(1)

    @property
    def file_path(self):
        return self._file_path.get()


def modify_email_config(email_config: EmailConfig):
    pop_up = CTkToplevel()
    pop_up.title("Email Config")
    pop_up.resizable(width=False, height=False)
    email_server_config_frame = EmailServerConfigFrame(pop_up, fg_color=FG_COLOR_ROOT)
    email_server_config_frame.set_sender_email(email_config.sender_email)
    email_server_config_frame.set_password(email_config.password)
    email_server_config_frame.set_username(email_config.username)
    email_server_config_frame.set_smtp_port(email_config.port)
    email_server_config_frame.set_smtp_server(email_config.smtp_server)
    email_server_config_frame.set_use_starttls(email_config.use_starttls)

    email_server_config_frame.grid(row=0, **PADDING)

    def set_and_quit():
        email_config.sender_email = email_server_config_frame.sender_email
        email_config.password = email_server_config_frame.password
        email_config.username = email_server_config_frame.username
        email_config.port = email_server_config_frame.smtp_port
        email_config.smtp_server = email_server_config_frame.smtp_server
        email_config.use_starttls = email_server_config_frame.use_starttls
        pop_up.destroy()

    ok_button = CTkButton(pop_up, text="Ok", command=set_and_quit)
    ok_button.grid(row=1, **PADDING)


def process(email_config: EmailConfig, document_dir_select: FolderSelect, document_config_frame: DocumentConfigFrame,
            name_email_file_select: FileSelect, name_email_config_frame: NameEmailConfigFrame,
            email_message_config_frame: EmailMessageConfigFrame):

    pop_up = CTkToplevel()
    pop_up.title("Processing")
    pop_up.resizable(width=False, height=False)
    loading_screen = CTkLabel(pop_up, text="Processing Names, Email addresses and Documents.")
    loading_screen.pack()

    def get_document_email_name():
        email_config.message = email_message_config_frame.email_message
        email_config.subject = email_message_config_frame.email_subject

        document_email_name = DocumentEmailName.from_data_paths(
            document_folder_path=document_dir_select.folder_path,
            excel_file_path=name_email_file_select.file_path,
            first_name_column=name_email_config_frame.first_name_column,
            email_column=name_email_config_frame.email_column,
            document_file_type=document_config_frame.document_file_types,
            last_name_column=name_email_config_frame.last_name_column,
            sheet_name=name_email_config_frame.sheet_name,
            start_row_for_names=name_email_config_frame.start_row_for_names,
            stop_row_for_names=name_email_config_frame.stop_row_for_names,
        )

        pop_up.destroy()

        confirm = CTkToplevel()
        confirm.title("Confirm")
        confirm.resizable(width=False, height=False)
        confirm_frame = CTkFrame(confirm)
        confirm_frame.grid(column=0, rowspan=4, sticky=N + S + E + W, **PADDING)
        font = CTkFont(family="Courier New", size=12)
        confirm_text = CTkTextbox(confirm_frame, width=800, height=400, font=font, wrap="none")
        confirm_text.pack(side=LEFT, fill=BOTH, expand=True)
        confirm_text.insert(END, document_email_name.info())
        confirm_text.configure(state=DISABLED)

        def send() -> None:
            confirm.destroy()
            pop_up = CTkToplevel()
            pop_up.title("Sending emails")
            pop_up.resizable(width=False, height=False)

            sending_email_frame = CTkFrame(pop_up, fg_color=FG_COLOR_ROOT)
            sending_email_frame.grid(column=0, rowspan=4, sticky=N + S + E + W, **PADDING)

            font = CTkFont(family="Courier New", size=12)
            sending_email_text = CTkTextbox(sending_email_frame, width=800, height=400, font=font, wrap="none")
            sending_email_text.grid(row=0, **PADDING)
            sending_email_text.insert(END, "Sent emails:\n")
            sending_email_text.configure(state=DISABLED)

            def save_text() -> None:
                file_path = fd.asksaveasfilename(defaultextension=".txt", filetypes=(("Text files", "*.txt"),))
                if file_path:
                    with open(file_path, "w", encoding="utf-8") as file:
                        file.write(sending_email_text.get(1.0, END))

            save_button = CTkButton(sending_email_frame, text="Save", command=save_text)
            save_button.grid(row=1, **PADDING)
            document_emails_to_send = document_email_name.document_to_unambiguous_name_and_email()
            max_progress = len(document_emails_to_send)
            progressbar = CTkProgressBar(sending_email_frame, width=600, height=10)
            progressbar.set(0)
            progressbar.grid(row=2, **PADDING)
            cancel_event = threading.Event()
            cancel_button = CTkButton(sending_email_frame, text="Cancel", command=cancel_event.set, fg_color="dark red")
            cancel_button.grid(row=3, column=0, **PADDING)
            document_name_email_sent = []
            failed_send_email = threading.Event()

            def on_send_email(sent_document_name_email: tuple[FilePath, tuple[str, EmailAddress]],
                              error: Optional[str]) -> bool:
                document_name, (name, email) = sent_document_name_email
                document_name = os.path.basename(document_name)
                sending_email_text.configure(state=NORMAL)
                if not error:
                    sending_email_text.insert(END, f"* {document_name:<40} -> {email} ({name})\n")
                    document_name_email_sent.append((document_name, name, email))
                else:
                    failed_send_email.set()
                    sending_email_text.insert(
                        END, f"* Could not send {document_name:<40} -> {email} ({name})\n. Error: {error}\n")
                sending_email_text.configure(state=DISABLED)
                progressbar.set((progressbar.get() * max_progress + 1) / max_progress)
                return False

            send_email_thread = threading.Thread(
                target=send_emails,
                args=(document_email_name, email_config, cancel_event, on_send_email),
                daemon=True,
            )
            send_email_thread.start()

            def check_send_progress() -> None:
                if send_email_thread.is_alive():
                    pop_up.after(200, check_send_progress)
                else:
                    sending_email_text.configure(state=NORMAL)
                    sending_email_text.insert(END, "Finished sending emails\n\n")
                    if failed_send_email.is_set():
                        sending_email_text.insert(END, "Failed to send emails to the following:\n")
                        for document_name, (name, email) in (
                            (k, v) for k, v in document_emails_to_send.items() if k not in document_name_email_sent):
                            sending_email_text.insert(END,
                                                      f"* {os.path.basename(document_name):<40} -> {email} ({name})\n")
                    else:
                        sending_email_text.insert(END, "All E-Mails sent successfully!\n")
                    sending_email_text.configure(state=DISABLED)
                    cancel_button.configure(text="Close", command=pop_up.destroy)
                    progressbar.set(1)

            pop_up.after(200, check_send_progress)

        button_frame = CTkFrame(confirm, fg_color=FG_COLOR_ROOT)
        button_frame.grid(row=4, **PADDING)
        confirm_button = CTkButton(button_frame, text="Send Emails", command=send, fg_color="green")
        confirm_button.grid(row=0, column=0, **PADDING)
        cancel_button = CTkButton(button_frame, text="Cancel", command=confirm.destroy, fg_color="dark red")
        cancel_button.grid(row=0, column=1, **PADDING)

    pop_up.after(200, get_document_email_name)


def main() -> int:
    main_config, email_config = load_config()
    root = ctk.CTk()
    root.title("DDist")
    root.resizable(width=False, height=False)

    document_dir_select = FolderSelect(root, "Document Directory*", fg_color=FG_COLOR_ROOT)
    document_dir_select.grid(row=0, sticky=W, **PADDING)
    document_dir_select.set_folder_path(main_config.document_folder_path)

    name_email_file_select = FileSelect(root,
                                        "Name and Email File*",
                                        filetypes=(*SUPPORTED_FILE_TYPES, ("All files", "*.*")),
                                        fg_color=FG_COLOR_ROOT)
    name_email_file_select.grid(row=1, sticky=W, **PADDING)
    name_email_file_select.set_file_path(main_config.excel_file_path)

    document_config_frame = DocumentConfigFrame(root, fg_color=FG_COLOR_ROOT)
    document_config_frame.grid(row=2, sticky=W, **PADDING)
    document_config_frame.set_document_file_types(main_config.document_file_type)

    name_email_config_frame = NameEmailConfigFrame(root, fg_color=FG_COLOR_ROOT)
    name_email_config_frame.grid(row=3, sticky=W, **PADDING)
    name_email_config_frame.set_email_column(main_config.email_column)
    name_email_config_frame.set_first_name_column(main_config.first_name_column)
    name_email_config_frame.set_last_name_column(main_config.last_name_column)
    name_email_config_frame.set_sheet_name(main_config.sheet_name)
    name_email_config_frame.set_start_row_for_names(main_config.start_row_for_names)
    name_email_config_frame.set_stop_row_for_names(main_config.stop_row_for_names)

    email_message_config_frame = EmailMessageConfigFrame(root, fg_color=FG_COLOR_ROOT)
    email_message_config_frame.grid(row=0, column=1, rowspan=4, sticky=S, **PADDING)
    email_message_config_frame.set_email_message(email_config.message)
    email_message_config_frame.set_email_subject(email_config.subject)

    start_button = CTkButton(
        root,
        text="Start",
        fg_color="green",
        command=lambda: process(email_config, document_dir_select, document_config_frame, name_email_file_select,
                                name_email_config_frame, email_message_config_frame))
    start_button.grid(pady=10, row=4, column=0)

    email_config_button = CTkButton(root, text="Email Config", command=lambda: modify_email_config(email_config))
    email_config_button.grid(pady=10, row=4, column=1)

    def on_closing():
        main_config.document_folder_path = document_dir_select.folder_path
        main_config.document_file_type = document_config_frame.document_file_types
        main_config.excel_file_path = name_email_file_select.file_path
        main_config.email_column = name_email_config_frame.email_column
        main_config.first_name_column = name_email_config_frame.first_name_column
        main_config.last_name_column = name_email_config_frame.last_name_column
        main_config.sheet_name = name_email_config_frame.sheet_name
        main_config.start_row_for_names = name_email_config_frame.start_row_for_names
        main_config.stop_row_for_names = name_email_config_frame.stop_row_for_names
        email_config.subject = email_message_config_frame.email_subject
        email_config.message = email_message_config_frame.email_message

        dump_config(main_config, email_config)
        root.quit()

    save_button = CTkButton(root, text="Quit", command=on_closing, fg_color="dark red")

    save_button.grid(row=5, column=0, columnspan=2, sticky=W + E, **PADDING)

    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()

    return 0


if __name__ == "__main__":
    main()
