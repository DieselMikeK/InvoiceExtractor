# Invoice Extractor

A Windows desktop application that automates invoice processing by downloading PDF invoices from Gmail, extracting structured data, and validating against SkuNexus purchase orders.

## Features

- **Gmail Integration**: Connects to Gmail via OAuth 2.0 to automatically download invoice PDF attachments
- **PDF Parsing**: Extracts text from PDFs using pdfplumber, with OCR fallback (Tesseract) for scanned documents
- **Data Extraction**: Parses invoices to extract:
  - Vendor name and address
  - Invoice number and dates
  - PO number
  - Line items (SKU, description, quantity, price, amount)
  - Shipping costs
- **QuickBooks Export**: Outputs to Excel format compatible with QuickBooks Bill Import
- **SkuNexus Validation**: Validates extracted invoice data against SkuNexus purchase orders via GraphQL API
- **GUI Interface**: User-friendly Tkinter interface with progress tracking and logging

## Screenshot

The application provides a simple interface with:
- **Go**: Start downloading and parsing invoices
- **Stop**: Halt the current operation
- **Open Spreadsheet**: View the generated Excel output
- **Open Invoices Folder**: Access downloaded invoice files
- **Validate POs**: Compare invoice data against SkuNexus

## Installation

### Prerequisites

- Python 3.10+
- Google Cloud project with Gmail API enabled
- Tesseract OCR (optional, for scanned PDFs)

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/DieselMikeK/InvoiceExtractor.git
   cd InvoiceExtractor
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up Google OAuth credentials:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a project and enable the Gmail API
   - Create OAuth 2.0 credentials (Desktop application)
   - Download the credentials and save as `client_secret.json` in the project folder

4. (Optional) Install Tesseract OCR for scanned PDF support:
   - Download from [Tesseract GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
   - Add to system PATH

## Usage

### Running the Application

```bash
python invoice_extractor_gui.py
```

Or use the pre-built executable: `InvoiceExtractor.exe`

### Workflow

1. **First Run**: Click "Go" to authenticate with Gmail (opens browser for OAuth)
2. **Download**: The app fetches emails and downloads invoice attachments to the `invoices/` folder
3. **Parse**: PDFs are parsed and data is extracted to `invoices_output.xlsx`
4. **Validate**: Click "Validate POs" to compare invoice data against SkuNexus purchase orders

### Output Format

The generated Excel file follows QuickBooks Bill Import format with columns:
- Bill No., Vendor, Mailing Address, Terms
- Bill Date, Due Date, Location, Memo (PO Number)
- Type, Category/Account, Product/Service
- Qty, Rate, Description, Amount
- Billable, Customer/Project, Tax Rate, Class
- SkuNexus Validation (Yes/No), SkuNexus Failed Fields

## Project Structure

```
InvoiceExtractor/
├── invoice_extractor_gui.py  # Main GUI application
├── gmail_client.py           # Gmail API integration
├── invoice_parser.py         # PDF parsing and data extraction
├── spreadsheet_writer.py     # Excel output generation
├── skunexus_client.py        # SkuNexus GraphQL API client
├── requirements.txt          # Python dependencies
├── client_secret.json        # Google OAuth credentials (not in repo)
├── token.pickle              # Cached auth token (not in repo)
├── invoices/                 # Downloaded invoice files (not in repo)
└── invoices_output.xlsx      # Generated output (not in repo)
```

## Building the Executable

To create a standalone Windows executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name InvoiceExtractor invoice_extractor_gui.py
```

The executable will be created in the `dist/` folder.

## Configuration

### Gmail Scopes

The application uses `gmail.modify` scope to:
- Read emails and download attachments
- Add labels to mark processed emails

### SkuNexus Permissions

For PO validation, the SkuNexus user needs:
- Purchase orders: Show, Index
- Products: Show, Index

## License

Private/Internal Use

## Acknowledgments

Built with assistance from Claude (Anthropic)
