from __future__ import annotations

from pprint import pprint

from document_distributor.document_distributor import list_document_file_paths
from document_distributor.document_distributor import \
    map_document_to_names_and_emails
from document_distributor.document_distributor import \
    map_name_and_email_to_document
from document_distributor.document_distributor import name_to_mail_from_excel
from document_distributor.document_distributor import send_email


def test_document_paths():
    document_file_paths = list_document_file_paths("tests/test_data/test_documents", [".pdf"])
    pprint(document_file_paths)


def test_email_starttsl_authenticated(smtpd, msg, sender_email, receiver_emails, subject, attachments):
    smtpd.config.use_ssl = True
    smtpd.config.use_starttls = True
    smtpd.config.enforce_auth = True

    send_email(sender_email=sender_email,
              smtp_server=smtpd.hostname,
              password=smtpd.config.login_password,
              username=smtpd.config.login_username,
              receiver_emails=receiver_emails,
              subject=subject,
              message=msg,
              port=smtpd.port,
              use_starttls=True,
              starttls_context=None,
              attachment_file_paths=attachments)
    pprint(smtpd.messages)
    assert len(smtpd.messages) == 1


def test_name_to_email():
    name_to_mail = name_to_mail_from_excel("tests/test_data/test_info.xlsx",
                                           sheet_name="Test Sheet",
                                           first_name_column="A",
                                           last_name_column="B",
                                           email_column="D",
                                           start=4,
                                           stop=5)
    pprint(name_to_mail)


def test_name_to_documents_and_vice_versa():
    document_file_paths = list_document_file_paths("tests/test_data/test_documents", [".pdf"])
    name_to_mail = name_to_mail_from_excel("tests/test_data/test_info.xlsx",
                                           sheet_name="Test Sheet",
                                           first_name_column="A",
                                           last_name_column="B",
                                           email_column="D",
                                           start=4,
                                           stop=5)
    name_to_documents = map_name_and_email_to_document(document_file_paths, name_to_mail)
    pprint(name_to_documents)
    document_to_names = map_document_to_names_and_emails(document_file_paths, name_to_mail)
    pprint(document_to_names)
