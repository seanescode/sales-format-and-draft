import os
import time
import win32com.client as win32
import subprocess
from modules.dialogs import show_error_dialog


def is_outlook_running():
    running_processes = subprocess.check_output(["tasklist"], text=True)
    return "OUTLOOK.EXE" in running_processes.upper()


def connect_to_outlook():
    last_error = None

    for _ in range(5):
        try:
            return win32.Dispatch("Outlook.Application")
        except Exception as e:
            last_error = e
            time.sleep(2)

    raise ConnectionError(f"Unable to connect to Outlook.\n\n{last_error}")


def set_email_fields(mail, recipients, cc_recipients, subject, body):
    mail.To = "; ".join(recipients)

    if cc_recipients:
        mail.CC = "; ".join(cc_recipients)

    mail.Subject = subject
    mail.Body = body


def add_attachments(mail, attachments):
    if not attachments:
        return

    for attachment in attachments:
        if not os.path.exists(attachment):
            raise FileNotFoundError(
                f"The following attachment could not be found:\n\n{attachment}"
            )
        mail.Attachments.Add(attachment)

def check_outlook_ready(recipients):
    if not is_outlook_running():
        show_error_dialog(
            title="Outlook Required",
            message="Automation Failed:\n\nThe Outlook desktop application must be open and running to generate this "
                    "email.\n\nPlease open Outlook and try again. "
        )
        return False
    if not recipients:
        show_error_dialog("No Recipients", "Please specify at least one recipient before creating the email.")
        return False
    return True

def create_outlook_email(recipients, cc_recipients, subject, body, attachments=None):

    if not is_outlook_running(): # this is a defensive check even though checking outlook is running at start this ensures is still running even if outlook is terminated between program being run and code getting to here
        show_error_dialog(
            title="Outlook Required",
            message="Automation Failed:\n\nThe Outlook desktop application must be open and running to generate this "
                    "email.\n\nPlease open Outlook and try again. "
        )
        return

    try:
        outlook = connect_to_outlook()
        mail = outlook.CreateItem(0)

        set_email_fields(mail, recipients, cc_recipients, subject, body)
        add_attachments(mail, attachments)

        mail.Display()

    except Exception as e:

        show_error_dialog("Execution Error", f"Email generation failed:\n\n{e}")