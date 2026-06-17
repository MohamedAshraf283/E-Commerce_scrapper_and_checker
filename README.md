# Salla / Zid Desktop Scraper

A professional desktop application for scraping and analyzing **Salla** and **Zid** e-commerce stores using a real browser session, Excel-driven workflows, and advanced data extraction modules.

---

# Overview

This project is designed to process large lists of online stores from an Excel file and perform one or more of the following tasks:

* Verify store availability and operational status.
* Extract store contact information.
* Extract products and categories.
* Generate structured Excel reports.
* Sort and organize Excel files independently.

The application uses a real browser profile (Brave/Chrome) through Playwright, allowing it to behave similarly to a normal user session.

---

# Main Features

## 1. Store Status Verification

Checks whether a store is:

* Working
* Suspended
* Under maintenance
* Temporarily unavailable
* Returning HTTP errors
* Facing DNS issues
* Protected by human verification pages

### Status Categories

| Status             | Description                                                     |
| ------------------ | --------------------------------------------------------------- |
| Working            | Store loaded successfully and valid store signals were detected |
| Not Working        | Store is suspended, unavailable, or broken                      |
| Human Verification | Security challenge or anti-bot page detected                    |
| Error              | Timeout, SSL, DNS, or unexpected errors                         |
| Unknown            | Insufficient signals to determine final state                   |

---

## 2. Contact Information Extraction

Extracts data from:

* Header
* Footer
* Contact pages
* Social media links

### Extracted Data

* Store name
* Emails
* Phone numbers
* WhatsApp links
* Instagram
* Twitter / X
* Snapchat
* TikTok
* Facebook
* Other social platforms

---

## 3. Product Extraction

Supports:

### Automatic Detection

The scraper attempts to automatically detect:

* Categories
* Product cards
* Product names
* Prices
* Images
* Product URLs

### Custom Selectors

Advanced users can provide CSS selectors through the GUI.

Example:

```json
{
  "category_links": ".category-link",
  "product_card": ".product-card",
  "product_name": ".product-title",
  "product_price": ".price",
  "product_url": "a",
  "product_image": "img"
}
```

---

# Excel Driven Workflow

The application is designed around Excel files.

## Required Column

At minimum:

| Column |
| ------ |
| link   |

Example:

| link                                     |
| ---------------------------------------- |
| [https://store1.com](https://store1.com) |
| [https://store2.com](https://store2.com) |

---

## Optional Columns

| Column           | Purpose                   |
| ---------------- | ------------------------- |
| categories       | Category URLs             |
| custom selectors | Project specific scraping |
| notes            | Internal notes            |

---

# User Interface

The desktop GUI contains multiple sections:

## Settings

Configure:

* Input Excel file
* Output Excel file
* Browser executable
* User Data directory
* Browser profile
* Concurrency
* Delays
* Timeouts
* Retry behavior

---

## Sheet Preview

Allows:

* Loading Excel file
* Viewing data
* Detecting columns
* Selecting link column
* Verifying records before execution

---

## Selectors & Categories

Configure:

* Category URL column
* Product extraction limits
* Custom CSS selectors

---

## Execution & Reporting

Displays:

* Live logs
* Progress bar
* Processed records
* Remaining records
* Results table
* Final summary report

---

## Excel Sorting Tool

Standalone utility for:

* Alphabetical sorting
* Numeric sorting
* Ascending order
* Descending order

Works independently from scraping operations.

---

# Browser Session Support

Supports:

* Brave Browser
* Google Chrome
* Chromium-based browsers

Example:

```text
Browser:
C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe

User Data:
C:\Users\User\AppData\Local\BraveSoftware\Brave-Browser\User Data

Profile:
Default
```

---

# Recommended Settings

## Status Only

```text
Concurrency: 2
Min Delay: 1
Max Delay: 3
Timeout: 20000
```

---

## Contacts

```text
Concurrency: 2
Min Delay: 2
Max Delay: 4
Timeout: 30000
```

---

## Products

```text
Concurrency: 1
Min Delay: 2
Max Delay: 5
Timeout: 30000
Max Categories: 5
Max Products: 100
```

---

# Human Verification Handling

When a human verification page is detected:

1. Browser session is closed.
2. Application waits for the configured period.
3. Browser session is reopened.
4. Store processing resumes.

The application does **not** attempt to bypass security systems.

---

# Output Files

The generated Excel file contains:

## Stores Sheet

Store level information:

* Store Name
* URL
* Platform
* Status
* Reason

---

## Products Sheet

Product level information:

* Product Name
* Price
* Product URL
* Image URL
* Category

---

## Summary Sheet

Execution statistics:

* Total Stores
* Successful Stores
* Failed Stores
* Human Verification Cases
* Products Extracted

---

# Installation

## 1. Create Virtual Environment

```bash
python -m venv .venv
```

## 2. Activate Environment

Windows:

```bash
.venv\Scripts\activate
```

---

## 3. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## 4. Install Playwright Browser

```bash
playwright install chromium
```

---

## 5. Run Application

```bash
python main.py
```

Or simply:

```bash
install_and_run.bat
```

---

# Project Structure

```text
project/
│
├── main.py
│
├── src/
│   ├── app.py
│   ├── scraper.py
│   ├── excel_io.py
│   ├── models.py
│   ├── utils.py
│
├── output/
│
├── requirements.txt
│
├── install_and_run.bat
├── run_app.bat
├── build_exe.bat
│
└── README.md
```

---

# Technical Stack

* Python 3.11+
* Playwright
* OpenPyXL
* Tkinter
* AsyncIO
* Threading
