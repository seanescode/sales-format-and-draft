import ast
import configparser


EMAIL_SETTINGS = "EMAIL_SETTINGS"
EXCEL_SETTINGS = "EXCEL_REPORT_FORMATTING"


def load_config(config_file):
    config = configparser.ConfigParser()
    config.read(config_file)
    return config


def get_email_settings(config):
    return {
        "recipients": [
            email.strip()
            for email in config.get(
                EMAIL_SETTINGS, "recipients"
            ).split(";")
        ],
        "cc_recipients": [
            email.strip()
            for email in config.get(
                EMAIL_SETTINGS, "cc_recipients"
            ).split(";")
        ],
        "subject": config.get(
            EMAIL_SETTINGS, "subject"
        ),
        "body": config.get(
            "EMAIL_CONTENT", "body"
        ),
    }


def get_excel_settings(config):
    return {
        "bold_header": config.getboolean(
            EXCEL_SETTINGS, "bold_heading"
        ),
        "italic_cells": config.getboolean(
            EXCEL_SETTINGS, "italic_font"
        ),
        "zoom_percent": config.getint(
            EXCEL_SETTINGS, "zoom_percentage"
        ),
        "header_colour": ast.literal_eval(
            config.get(EXCEL_SETTINGS, "heading_colour")
        ),
        "even_row_colour": ast.literal_eval(
            config.get(EXCEL_SETTINGS, "even_row_colour")
        ),
        "odd_row_colour": ast.literal_eval(
            config.get(EXCEL_SETTINGS, "odd_row_colour")
        ),
        "sheet_name": config.get(
            EXCEL_SETTINGS, "sheet_name"
        ),
        "attachment_paths": [
            file.strip()
            for file in config.get(
                EXCEL_SETTINGS, "attachment_paths"
            ).split(";")
        ],
    }