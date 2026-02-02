# ParaBank End-to-End Automation Framework

Automated end-to-end testing and reporting framework for the **ParaBank** demo banking application using **Python, Selenium, Pandas, and OpenPyXL**.

This project simulates realistic banking workflows at scale:
- User registration
- Account creation
- Loan requests
- Data validation
- Structured Excel reporting with conditional formatting

---

## Why This Project Exists

Manual testing of banking flows is slow, error-prone, and impossible to scale.  
This framework demonstrates how to:

- Drive **data-driven UI automation**
- Validate user data before execution
- Handle failures gracefully
- Produce **auditable, business-ready reports**
- Keep automation code modular and readable

This is not a script — it’s a **mini automation framework**.

---

## Features

- **CSV-driven execution** (multiple users, zero code changes)
- **Strong data validation** before UI interaction
- **Selenium WebDriver automation**
- **Loan logic simulation** (currency conversion, down payment calculation)
- **Automatic screenshot capture on failure**
- **Excel report generation with styling**
  - Conditional formatting
  - Currency formatting
  - Auto-sizing columns
  - Filters and frozen headers
- **Timestamped execution records**
- **Clean logging**

---

## Tech Stack

- Python 3.12.7
- Selenium
- Pandas
- OpenPyXL
- webdriver-manager
- Chrome WebDriver

---

## Project Structure

