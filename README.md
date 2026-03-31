# Gino's PLUG Option Pricer

A Flask web application for pricing options packages with PLUG strikes. All strikes are assumed to be in % terms. B/O (bid/offer) is configured as a percentage of notional.

## Features

- **PLUG Strike Solving**: Automatically solves for the PLUG strike such that Package Premium + B/O = Target Notional
- **B/O Shifts Table**: Configure bid/offer as % of notional by ticker and expiry month (out to 2 years)
- **% Strike Support**: All strikes are interpreted as percentages of the reference price
- **Excel I/O**: Upload Excel files and download priced results

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python app.py
```

The app runs on `http://localhost:1237` (port 1237).

For network access:
```bash
http://YOUR_IP_ADDRESS:1237
```

## File Format

The Excel file should have these columns:
- **Package #**: Package identifier
- **Symbol**: Underlying symbol (e.g., SPY)
- **Style**: Option style (e.g., "European CASH")
- **Date**: Expiration date
- **Strike Or %**: Strike as percentage (e.g., 0.91 = 91%) or "PLUG"
- **Qty**: Quantity
- **Call or Put**: CALL or PUT
- **Side**: BYO (customer buys) or SLO (customer sells)
- **Ref Price**: Reference price for the underlying
- **Notional**: Target notional (what customer pays) - required for PLUG rows
- **Mult**: Multiplier (optional, defaults to 100)

## Output Columns

- **Strike**: Solved dollar strike
- **Strike %**: Strike as percentage of ref price
- **Cap**: The solved PLUG strike as % of ref price (only for PLUG rows)
- **Notional**: Target notional from input
- **B/O %**: Actual B/O achieved as % of notional
- **Actual B/O**: Actual B/O in dollars

## B/O Shift Configuration

1. Go to the "B/O Shifts" tab
2. Select a ticker from the dropdown and click "Add Ticker"
3. Enter B/O percentages for each expiry month
4. Click "Save All Changes"

The B/O for a package is calculated as:
```
B/O = Notional × (B/O Shift % / 100)
```

The shift is interpolated based on the option's expiry date.

## Integration

Replace the `create_model()` function in `app.py` with your actual pricing model:

```python
from utils import login_connector

def create_model(symbol, date_val, quantity, call_put, style, strike, ref_price, side, settle_timing):
    # Your implementation here
    pass
```
