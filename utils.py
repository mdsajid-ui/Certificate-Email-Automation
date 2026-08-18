"""
utils.py
--------
Shared helpers: Excel validation, logging setup, and report building.
"""

import logging
import os
from datetime import datetime

import pandas as pd

REQUIRED_COLUMNS = ["Name", "Mobile Number", "Email ID"]

LOG_PATH = "email_log.txt"


def setup_logger() -> logging.Logger:
    logger = logging.getLogger("cert_email_app")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s | %(message)s"))
        logger.addHandler(fh)
    return logger


def log_event(recipient: str, status: str, error: str = ""):
    logger = setup_logger()
    detail = f"Recipient={recipient} | Status={status}"
    if error:
        detail += f" | Error={error}"
    logger.info(detail)


def find_matching_column(columns, keyword: str):
    """Case-insensitive / whitespace-tolerant match for a required column name."""
    keyword_norm = keyword.lower().replace(" ", "")
    for col in columns:
        if str(col).lower().replace(" ", "") == keyword_norm:
            return col
    return None


def validate_excel_columns(df: pd.DataFrame):
    """
    Confirms Name, Mobile Number, and Email ID columns exist (flexibly matched).
    Returns (is_valid, column_map, missing_columns).
    """
    column_map = {}
    missing = []
    for required in REQUIRED_COLUMNS:
        match = find_matching_column(df.columns, required)
        if match:
            column_map[required] = match
        else:
            missing.append(required)
    return (len(missing) == 0), column_map, missing


def normalize_records(df: pd.DataFrame, column_map: dict):
    """Return a list of dicts with clean, standardized keys."""
    records = []
    for _, row in df.iterrows():
        name = str(row[column_map["Name"]]).strip()
        mobile = str(row[column_map["Mobile Number"]]).strip()
        email = str(row[column_map["Email ID"]]).strip()
        if mobile.lower() == "nan":
            mobile = ""
        if name.lower() == "nan" or not name:
            continue
        records.append({"Name": name, "Mobile Number": mobile, "Email ID": email})
    return records


def build_report_dataframe(results: list) -> pd.DataFrame:
    """
    results: list of dicts with keys:
        Name, Mobile Number, Email ID, Certificate Generated, Email Sent,
        Sent Date & Time, Error Message
    """
    columns = [
        "Name",
        "Mobile Number",
        "Email ID",
        "Certificate Generated",
        "Email Sent",
        "Sent Date & Time",
        "Error Message",
    ]
    df = pd.DataFrame(results)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    return df[columns]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
