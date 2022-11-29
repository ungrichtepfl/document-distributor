#!/usr/bin/env python3
"""Send bild to clients"""

from __future__ import annotations

import os
import smtplib
import ssl
import sys
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_mail(message: str, attachment_file_paths: list[str] = None) -> None:
    if attachment_file_paths is None:
        attachment_file_paths = []
    port: int = 587  # For starttls # TODO: Ask nati
    smtp_server: str = "smtp.office365.com"  # TODO: Ask nati
    sender_email: str = "christoph.ungricht@outlook.com"
    receiver_emails: list[str] = [
        "christoph.ungricht@outlook.com",
    ]
    username: str = "christoph.ungricht@outlook.com"
    password: str = os.environ.get("EMAIL_PASSWORD")
    assert password is not None

    message_sent = MIMEMultipart()
    message_sent["Subject"] = f"Rechnung von "\
                              f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    message_sent["From"] = sender_email
    message_sent["To"] = ', '.join(receiver_emails)

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

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.ehlo()  # Can be omitted
        server.starttls(context=context)
        server.ehlo()  # Can be omitted
        server.login(username, password)
        server.sendmail(sender_email, receiver_emails,
                        message_sent.as_string())


def app() -> int:

    msg: str = ""
    attachments: list[str] = []
    if len(sys.argv) > 1:
        msg = sys.argv[1]
    elif len(sys.argv) > 2:
        attachments = sys.argv[2:]

    if (msg or attachments):
        try:
            send_mail(msg, attachments)
        except Exception:
            print(f"Could not send:\n\"\n{msg}\"", file=sys.stderr)
            import traceback
            traceback.print_exc()
            return 1
        print(f"Sent by mail:\n\n{msg}")
        if attachments:
            attachments_txt: str = "\n".join(attachments)
            print(f"Attached documents:\n{attachments_txt}")
    else:
        print("Nothing to send. No message or attachment given.")

    return 0


if __name__ == "__main__":
    raise SystemExit(app())
