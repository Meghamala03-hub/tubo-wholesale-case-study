# Slow-Stock Action Tool

## Project Overview

This project is a browser-based tool designed to help review inventory, identify slow-moving stock, classify items, and support customer targeting decisions.

## Key Features

* Upload inventory and sales data.
* Validate input data and identify data-quality issues.
* Classify active items as Red List, No Action, or Push.
* Calculate aged stock value and months to clear.
* Identify potential customer targets for Push items.
* Export results to Excel.

## Classification Rules

* **Red List:** Six-month sales are less than one master pack.
* **No Action:** No aged stock exists, or months to clear is less than three.
* **Push:** Items that meet neither of the above conditions.

Aged stock includes the 91–180-day and 180+-day inventory buckets.

## Technology

* Python
* Streamlit
* Pandas
* Excel processing

## Run Locally

Install dependencies:

`pip install -r requirements.txt`

Run the application:

`streamlit run app.py`

## Deployment

The application is intended to be deployed through Streamlit Community Cloud using this GitHub repository.

## AI Assistance Disclosure

AI assistance was used to support application development and documentation. The calculations, classification logic, and generated outputs should be validated against the case study requirements before submission.
