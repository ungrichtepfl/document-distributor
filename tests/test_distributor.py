from __future__ import annotations

import os
import sys
from datetime import datetime
from pprint import pprint

from document_distributor.document_distributor import (
    list_document_file_paths, map_document_to_names, map_name_to_documents,
    name_to_mail_from_excel, send_mail)


def test_document_paths():
    document_file_paths = list_document_file_paths("data/rechnungen/")
    pprint(document_file_paths)


def test_email():
    msg: str = ""
    attachments: list[str] = []

    sender_email: str = "info@kaesers-schloss.ch"
    smtp_server: str = "mail5.peaknetworks.net"
    username: str = "info@kaesers-schloss.ch"
    # smtp_server: str = "smtp.office365.com"
    # sender_email: str = "christoph.ungricht@outlook.com"
    # username: str = "christoph.ungricht@outlook.com"
    subject = f"Rechnung von {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    password: str = os.environ.get("EMAIL_PASSWORD")
    receiver_emails: list[str] = [
        "christoph.ungricht@outlook.com",
    ]
    assert password is not None

    if (msg or attachments):
        try:
            send_mail(sender_email=sender_email,
                      smtp_server=smtp_server,
                      username=username,
                      password=password,
                      receiver_emails=receiver_emails,
                      subject=subject,
                      message=msg,
                      attachment_file_paths=attachments)
        except Exception:
            print(f"Could not send:\n\"\n{msg}\"", file=sys.stderr)
            import traceback
            traceback.print_exc()
            return
        print(f"Sent by mail:\n\n{msg}")
        if attachments:
            attachments_txt: str = "\n".join(attachments)
            print(f"Attached documents:\n{attachments_txt}")
    else:
        print("Nothing to send. No message or attachment given.")


def test_name_to_email():
    name_to_mail = name_to_mail_from_excel(
        "./data/Zahlungen Gmüeschistli-20221019.xlsx",
        sheet_name="Abos",
        first_name_column="A",
        last_name_column="B",
        email_column="I",
        start=4,
        stop=108)
    pprint(name_to_mail)


def test_name_to_documents_and_vice_versa():
    document_file_paths = list_document_file_paths("data/rechnungen/")
    name_to_mail = name_to_mail_from_excel(
        "./data/Zahlungen Gmüeschistli-20221019.xlsx",
        sheet_name="Abos",
        first_name_column="A",
        last_name_column="B",
        email_column="I",
        start=4,
        stop=108)
    name_to_documents = map_name_to_documents(document_file_paths,
                                              name_to_mail)
    pprint(name_to_documents)
    document_to_names = map_document_to_names(document_file_paths,
                                              name_to_mail)
    pprint(document_to_names)
