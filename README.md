# Credit Situation, VAT Invoice Tracking & Supplier Follow-up Automation

## Overview

This project automates two related operational workflows for hotel booking invoice management:

1. **Credit Situation & VAT Invoice Tracking** — processes booking and invoice data, reconciles received invoices, combines current and historical records, and generates a structured Excel workbook.
2. **Pending VAT Invoice Follow-up** — identifies suppliers with outstanding invoices and sends each supplier a customized email listing the relevant bookings, with the required company VAT and registration documents attached.

The goal is to reduce manual reconciliation and follow-up, improve visibility into outstanding invoices, and standardize communication with suppliers.

## Project Workflows

### 1. Credit Situation & VAT Invoice Tracking

The notebook combines booking data retrieved from Juniper SOAP APIs with received-invoice records, historical credit situation data, supplier/product-group information, and hotel email references.

**Main capabilities**
- Retrieves booking lists and booking details through Juniper SOAP APIs.
- Uses asynchronous HTTP requests to retrieve booking details efficiently.
- Parses SOAP XML responses and processes the results with Pandas.
- Maps product groups, suppliers, departments, hotel countries, destination managers, and hotel email contacts.
- Applies configured business rules and filters to the booking dataset.
- Matches booking codes against received-invoice records to determine invoice status.
- Combines new booking data with the previous credit situation report.
- Removes duplicate booking records and standardizes fields.
- Preserves available supplier contract information and notes.
- Generates an Excel workbook with `Full`, `Summary`, and `Contracts` worksheets.

#### Excel output

| Worksheet | Purpose |
|---|---|
| `Full` | Consolidated booking-level data, booking details, costs, travel dates, hotel/supplier information, and invoice status |
| `Summary` | Invoice counts and follow-up information grouped by relevant hotel, supplier, country, and destination-manager fields |
| `Contracts` | Supplier contract information, including chain, contract status, contracting party, and notes |

### 2. Pending VAT Invoice Follow-up to Suppliers

The second notebook automates supplier communication for invoices that have not yet been received.

**Main capabilities**
- Loads the pending-invoice dataset from Excel.
- Removes duplicate records using `BookingCode`.
- Standardizes supplier names to improve matching.
- Matches suppliers against a reference file containing supplier email addresses and CC recipients.
- Separates suppliers with available email contacts from suppliers without email contacts.
- Exports separate Excel files for records that can be emailed and records requiring contact-data follow-up.
- Groups pending bookings by supplier.
- Creates a customized HTML email for each supplier, including a table of relevant reservations.
- Includes booking code, passenger name, travel dates, rate, currency, and remarks in the email.
- Attaches the configured company registration and VAT certificate documents.
- Sends emails through Microsoft Outlook using the configured sender account.
- Includes basic Outlook recovery logic and periodically restarts Outlook during larger sending runs.
- Limits sending volume and pauses between batches to help manage email throughput.

#### Supplier email output

Each supplier receives a separate email with a subject similar to:

`Pending Invoices - [Supplier Name] - [Date]`

The email asks the supplier to provide outstanding invoices and requests that the invoices be issued to the appropriate company entity with the required VAT details.

The workflow also produces two Excel files:

- `SendPendingInvoices[Date].xlsx` — pending-invoice records associated with suppliers that have email addresses.
- `NotSentPendingInvoices[Date].xlsx` — records for suppliers without a matched email address, for manual review and contact-data updates.

## End-to-End Workflow

1. Retrieve or load booking and invoice data.
2. Clean and standardize booking codes, supplier names, and relevant fields.
3. Match booking records with invoice status and supplier contact information.
4. Generate the credit situation report and identify outstanding invoices.
5. Prepare supplier-specific pending-invoice datasets.
6. Send customized follow-up emails to suppliers with valid email contacts.
7. Export records that could not be emailed so they can be investigated.
8. Review the resulting reports and follow up on outstanding items.

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Automation and business logic |
| Pandas | Data cleaning, filtering, merging, grouping, and Excel data processing |
| Requests | Synchronous HTTP/API requests |
| Aiohttp | Asynchronous HTTP requests |
| Asyncio | Concurrent booking-detail retrieval |
| XMLtodict | SOAP XML response parsing |
| Nest Asyncio | Supporting nested asynchronous event loops |
| OpenPyXL | Reading and processing Excel input files |
| XlsxWriter | Writing formatted Excel output |
| PyWin32 (`win32com.client`) | Microsoft Outlook integration and email sending |
| TQDM | Progress tracking during data retrieval |
| Regular Expressions | Text cleaning and email-related processing |

## Requirements

- Python 3.x
- Access to the required Juniper SOAP APIs and permissions
- Microsoft Windows
- Microsoft Outlook configured with the authorized sender account
- Access to the required input Excel files, historical reports, and supplier contact references
- Access to the approved company registration and VAT certificate files

## Installation

Clone the repository:

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
cd YOUR_REPOSITORY
```

Install the Python dependencies:

```bash
pip install pandas requests aiohttp xmltodict tqdm nest_asyncio openpyxl xlsxwriter pywin32
```

## Configuration

Before running the notebooks, update the environment-specific settings:

- Juniper API endpoint URLs and authentication details.
- Input file locations and output directories.
- Travel-date ranges and business filtering rules.
- Supplier email and CC reference file.
- Outlook sender account.
- Company registration and VAT certificate attachment paths.
- Outlook sending limits and restart interval.

The notebooks currently rely on environment-specific Windows paths and internal data sources. Update these paths for your environment and confirm that the required files and accounts are available.

## Execution

Open the relevant notebook in Jupyter Notebook or JupyterLab and execute its cells in order.

**Credit Situation & VAT Invoice Tracking**
- Run the notebook to retrieve/process booking information, reconcile invoice status, integrate historical records, and generate the Excel workbook.

**Pending VAT Invoice Follow-up**
- Confirm the pending-invoice input file, supplier email mapping, Outlook sender account, and attachment paths.
- Review the dataset and recipient mapping before running the email-sending cells.
- Run the notebook to generate the sendable/not-sendable Excel files and send supplier-specific follow-up emails.

> **Important:** The supplier follow-up notebook sends real emails through Outlook. Test with a small, approved recipient set before running it against the full supplier list.

## Business Value

- Reduces manual invoice reconciliation and follow-up.
- Improves visibility into received and outstanding invoices.
- Standardizes supplier communications.
- Helps identify suppliers whose email contact information is missing.
- Maintains historical credit situation and contract information.
- Produces structured Excel outputs for operational monitoring and follow-up.
- Supports more consistent invoice collection and VAT documentation processes.

## Data Privacy & Security

These workflows may process confidential booking, passenger, supplier, invoice, email, and contract information.

Before publishing this repository publicly:

- Never commit API credentials, passwords, authentication tokens, or private endpoint details.
- Replace internal Windows/network paths and internal account identifiers with placeholders.
- Do not upload real booking records, passenger names, supplier email lists, invoice data, or company documents.
- Do not commit the company VAT certificate, trade license, or email signature assets.
- Store credentials in environment variables or an approved secrets manager.
- Review all email automation settings and business rules before allowing execution.
- Confirm that you have permission to publish the code and related business logic.

## Suggested Repository Structure

```text
.
├── README.md
├── requirements.txt
├── notebooks/
│   ├── Final_Credit_Situation_VAT.ipynb
│   └── New_Send_Pending_VCC_to_Supplier_VAT.ipynb
└── .gitignore
```

Keep input data, generated reports, credentials, and company documents outside the public repository.

## Author

**Ahmed Saker**  
Data Analyst | Python Automation | SQL | Power BI | Excel | Data Analytics
