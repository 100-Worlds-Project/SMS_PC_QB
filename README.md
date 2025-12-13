# SMS Pricing Calculator

A desktop Python application for calculating print pricing and generating
QuickBooks Online invoices via the Intuit API.

## Features
- Tkinter-based desktop UI
- OAuth 2.0 with token refresh
- QuickBooks customer & invoice creation
- PDF invoice generation
- Sandbox / Production environment support

## Setup
1. Create a `.env` file based on `.env.example`
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
3. Download the draft invoice word doc to same dir as everything else. (Note: This program requires current version of MS Word).
4. Download 'SMS Pricing Calculator vs QB_019.py' and run. Input correct/test information as needed (bulk_pricing dictionary, volume discounts starting at line 2235, prices in send_to_draft, current path for draft invoice (line 1269))
