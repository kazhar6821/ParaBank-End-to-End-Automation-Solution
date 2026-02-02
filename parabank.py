from pathlib import Path
import re
import random
import logging
from datetime import datetime
from dataclasses import dataclass
from typing import List
import pandas as pd
from dateutil import parser as date_parser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

# PATHS
SCRIPT_DIR = Path(__file__).resolve().parent
CWD = Path.cwd()

CSV_NAME = "ParaBank_users.csv"

if (SCRIPT_DIR / CSV_NAME).exists():
    CSV_PATH = SCRIPT_DIR / CSV_NAME
elif (CWD / CSV_NAME).exists():
    CSV_PATH = CWD / CSV_NAME
else:
    raise FileNotFoundError("CSV NOT FOUND")

REPORT_XLSX = SCRIPT_DIR / "Parabank_Automation_Report.xlsx"
SCREENSHOT_DIR = SCRIPT_DIR / "screenshots"
SCREENSHOT_DIR.mkdir(exist_ok=True)

# CONFIG
BASE_URL = "https://parabank.parasoft.com/parabank/index.htm"
USD_TO_EUR = 0.92
LOAN_AMOUNT_USD = 10000
DOWNPAYMENT_FACTOR = 0.20
TIMEOUT = 15

# LOGGING
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
log = logging.getLogger("parabank")

# DATA MODEL
@dataclass
class User:
    first_name: str
    last_name: str
    address: str
    city: str
    state: str
    zip_code: str
    phone: str
    ssn: str
    username: str
    password: str
    dob: str
    initial_deposit: float

# HELPERS
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "_", regex=True)
    )
    return df

def parse_dob(v) -> str:
    try:
        return date_parser.parse(str(v)).date().isoformat()
    except Exception:
        return ""

def generate_card():
    card = "4" + "".join(str(random.randint(0, 9)) for _ in range(15))
    cvv = random.randint(100, 999)
    return card, cvv


def usd_to_eur(x: float) -> float:
    return round(x * USD_TO_EUR, 2)


def validate_user(u: User) -> List[str]:
    errors = []
    for k, v in u.__dict__.items():
        if k != "dob" and not v:
            errors.append(f"missing {k}")
    if not re.fullmatch(r"\d+", u.zip_code):
        errors.append("zip_code invalid")
    return errors

# SELENIUM
def make_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1400,900")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def wait(driver, by, value):
    return WebDriverWait(driver, TIMEOUT).until(
        EC.presence_of_element_located((by, value))
    )

# PARABANK ACTIONS
def register(driver, u: User):
    driver.find_element(By.LINK_TEXT, "Register").click()
    wait(driver, By.ID, "customer.firstName")

    driver.find_element(By.ID, "customer.firstName").send_keys(u.first_name)
    driver.find_element(By.ID, "customer.lastName").send_keys(u.last_name)
    driver.find_element(By.ID, "customer.address.street").send_keys(u.address)
    driver.find_element(By.ID, "customer.address.city").send_keys(u.city)
    driver.find_element(By.ID, "customer.address.state").send_keys(u.state)
    driver.find_element(By.ID, "customer.address.zipCode").send_keys(u.zip_code)
    driver.find_element(By.ID, "customer.phoneNumber").send_keys(u.phone)
    driver.find_element(By.ID, "customer.ssn").send_keys(u.ssn)
    driver.find_element(By.ID, "customer.username").send_keys(u.username)
    driver.find_element(By.ID, "customer.password").send_keys(u.password)
    driver.find_element(By.ID, "repeatedPassword").send_keys(u.password)

    driver.find_element(By.XPATH, "//input[@value='Register']").click()
    wait(driver, By.LINK_TEXT, "Log Out")


def open_account(driver) -> str:
    driver.find_element(By.LINK_TEXT, "Open New Account").click()
    wait(driver, By.ID, "type")
    driver.find_element(By.XPATH, "//input[@value='Open New Account']").click()
    try:
        return wait(driver, By.ID, "newAccountId").text
    except TimeoutException:
        return ""


def request_loan(driver, down_payment):
    driver.find_element(By.LINK_TEXT, "Request Loan").click()
    wait(driver, By.ID, "amount")

    driver.find_element(By.ID, "amount").send_keys(str(LOAN_AMOUNT_USD))
    driver.find_element(By.ID, "downPayment").send_keys(str(down_payment))
    driver.find_element(By.XPATH, "//input[@value='Apply Now']").click()

    try:
        status = wait(driver, By.ID, "loanStatus").text
    except TimeoutException:
        status = "UNKNOWN"

    try:
        acc = driver.find_element(By.ID, "newAccountId").text
    except Exception:
        acc = ""

    return status, acc

# EXCEL FORMATTER
def format_excel_report(path: Path, df: pd.DataFrame):
    wb = load_workbook(path)
    ws = wb.active
    ws.title = "Automation Report"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center")
    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Header
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin

    # Auto width
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 35)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Borders
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin

    # Currency columns
    for col_name in ["Loan USD", "Loan EUR", "Down Payment USD", "Down Payment EUR"]:
        if col_name in df.columns:
            idx = df.columns.get_loc(col_name) + 1
            for col in ws.iter_cols(min_col=idx, max_col=idx, min_row=2):
                for cell in col:
                    cell.number_format = "#,##0.00"

    # Conditional formatting
    status_idx = df.columns.get_loc("Status") + 1
    letter = ws.cell(row=1, column=status_idx).column_letter

    ws.conditional_formatting.add(
        f"{letter}2:{letter}{ws.max_row}",
        FormulaRule(
            formula=[f'{letter}2="COMPLETED"'],
            fill=PatternFill("solid", fgColor="C6EFCE"),
            font=Font(color="006100"),
        ),
    )

    ws.conditional_formatting.add(
        f"{letter}2:{letter}{ws.max_row}",
        FormulaRule(
            formula=[f'{letter}2="FAILED"'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
            font=Font(color="9C0006"),
        ),
    )

    wb.save(path)

# MAIN
def main():
    df = normalize_columns(pd.read_csv(CSV_PATH))
    report = []

    for _, r in df.iterrows():
        user = User(
            first_name=str(r.get("first_name", "")),
            last_name=str(r.get("last_name", "")),
            address=str(r.get("address", "")),
            city=str(r.get("city", "")),
            state=str(r.get("state", "")),
            zip_code=str(r.get("zip_code", "")),
            phone=str(r.get("phone_number", "")),
            ssn=str(r.get("ssn", "")),
            username=str(r.get("username", "")),
            password=str(r.get("password", "")),
            dob=parse_dob(r.get("dob", "")),
            initial_deposit=float(r.get("initial_deposit", 0)),
        )

        errors = validate_user(user)
        card, cvv = generate_card()
        down_payment = round(user.initial_deposit * DOWNPAYMENT_FACTOR, 2)

        row = {
            "Username": user.username,
            "DOB": user.dob,
            "Debit Card": card,
            "CVV": cvv,
            "Loan USD": LOAN_AMOUNT_USD,
            "Loan EUR": usd_to_eur(LOAN_AMOUNT_USD),
            "Down Payment USD": down_payment,
            "Down Payment EUR": usd_to_eur(down_payment),
            "Account ID": "",
            "Loan Account ID": "",
            "Loan Status": "",
            "Status": "FAILED",
            "Reason": "",
            "Timestamp": datetime.now().isoformat(timespec="seconds"),
        }

        if errors:
            row["Reason"] = "; ".join(errors)
            report.append(row)
            continue

        driver = None
        try:
            driver = make_driver()
            driver.get(BASE_URL)

            register(driver, user)
            row["Account ID"] = open_account(driver)

            status, loan_acc = request_loan(driver, down_payment)
            row["Loan Status"] = status
            row["Loan Account ID"] = loan_acc
            row["Status"] = "COMPLETED"

            driver.find_element(By.LINK_TEXT, "Log Out").click()

        except Exception as e:
            row["Reason"] = str(e)
            if driver:
                driver.save_screenshot(SCREENSHOT_DIR / f"error_{user.username}.png")

        finally:
            if driver:
                driver.quit()

        report.append(row)

    df_report = pd.DataFrame(report)
    df_report.to_excel(REPORT_XLSX, index=False)
    format_excel_report(REPORT_XLSX, df_report)

    log.info(f"REPORT CREATED: {REPORT_XLSX}")


if __name__ == "__main__":
    main()
