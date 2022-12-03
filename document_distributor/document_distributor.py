#!/usr/bin/env python3
"""Send bild to clients"""

from __future__ import annotations

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import re
import smtplib
import ssl
from ssl import SSLContext
import string
from typing import Mapping, Optional, Sequence

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

EMAIL_REGEX = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"


def valid_email(email: str) -> bool:
    return re.fullmatch(EMAIL_REGEX, email) is not None


def col2num(col: str) -> int:
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def name_to_mail_from_excel(excel_file_path: str,
                            first_name_column: str,
                            email_column: str,
                            last_name_column: Optional[str] = None,
                            sheet_name: Optional[str] = None,
                            start: int = 1,
                            stop: Optional[int] = None) -> Mapping[str, str]:
    workbook: Workbook = openpyxl.load_workbook(excel_file_path)
    sheet: Worksheet = workbook[sheet_name]
    stop = stop if stop is not None else sheet.max_row
    first_name_column_idx = col2num(first_name_column)
    last_name_column_idx = col2num(
        last_name_column) if last_name_column is not None else None
    email_column_idx = col2num(email_column)
    names = {}
    for row in sheet.iter_rows(min_row=start, max_row=stop):
        first_name: str = row[first_name_column_idx - 1].value
        last_name: Optional[str] = row[
            last_name_column_idx -
            1].value if last_name_column_idx is not None else None

        name = first_name.strip(" ")
        if last_name is not None:
            name = f"{name} {last_name.strip(' ')}"
        email = row[email_column_idx - 1].value.strip(" ")
        names[name] = email

    return names


def send_mail(sender_email: str,
              smtp_server: str,
              receiver_emails: Sequence[str],
              subject: str,
              message: str,
              port: int = 587,
              use_starttls: bool = True,
              starttls_context: Optional[SSLContext] = ssl.create_default_context(),
              username: Optional[str] = None,
              password: Optional[str] = None,
              attachment_file_paths: Sequence[str] | None = None,
              ) -> None:
    """Send an email."""

    if attachment_file_paths is None:
        attachment_file_paths = []

    message_sent = MIMEMultipart()
    message_sent["Subject"] = subject
    message_sent["From"] = sender_email
    message_sent["To"] = ", ".join(receiver_emails)

    message_sent.attach(MIMEText(message, "plain"))

    attachment: str
    for attachment in attachment_file_paths:
        if attachment:  # Not empty string
            part = MIMEBase("application", "octet-stream")
            with open(attachment, 'rb') as bytestream:
                part.set_payload(bytestream.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(attachment)}")
            message_sent.attach(part)

    with smtplib.SMTP(smtp_server, port) as client:
        if use_starttls:
            client.ehlo()
            client.starttls(context=starttls_context)
            # client.starttls()
        if username is not None and password is not None:
            client.ehlo()
            client.login(username, password)
        client.ehlo()
        client.sendmail(sender_email, receiver_emails,
                        message_sent.as_string())


def list_document_file_paths(document_folder_path: str) -> Sequence[str]:
    document_folder_path = os.path.abspath(
        os.path.normpath(document_folder_path))
    pdf_files = filter(lambda f: ".pdf" in f, os.listdir(document_folder_path))
    return tuple(
        map(lambda f: os.path.join(document_folder_path, f), pdf_files))


def is_name_in_document(name: str, document_path: str):
    min_matches = 2
    document_name = os.path.basename(document_path)
    total_matches = 0
    name = name.lower()
    document_name = document_name.lower()
    for token in name.split(" "):
        if token in document_name:
            total_matches += 1
            if total_matches >= min_matches:
                return True
    return False


def map_document_to_names(
        document_file_paths: Sequence[str],
        name_to_email: Mapping[str, str]) -> Mapping[str, Mapping[str, str]]:

    document_to_name_and_email = {}
    for document in document_file_paths:
        matching_names = filter(
            lambda n_e: valid_email(n_e[1]) and is_name_in_document(
                n_e[0], document), name_to_email.items())
        document_to_name_and_email[document] = dict(matching_names)

    return document_to_name_and_email


def map_name_to_documents(
    document_file_paths: Sequence[str],
    name_to_email: Mapping[str,
                           str]) -> Mapping[tuple[str, str], Sequence[str]]:

    name_to_email_and_document = {}
    for (name, email) in name_to_email.items():
        matching_documents = filter(
            lambda b: valid_email(email) and is_name_in_document(name, b),
            document_file_paths)
        name_to_email_and_document[(name, email)] = tuple(matching_documents)
    return name_to_email_and_document




def app() -> int:
    """Application entry point"""

    test_email()
    test_name_to_email()
    test_document_paths()
    test_name_to_documents_and_vice_versa()

    return 0


if __name__ == "__main__":
    raise SystemExit(app())
