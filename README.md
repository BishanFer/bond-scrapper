# Sri Lanka Treasury Bond Yield Scraper

Automated scraper that extracts daily Treasury Bond yields from the Sri Lanka Treasury Department website and generates formatted Excel reports.

## Features

- **Automated daily scraping** via GitHub Actions
- **Two bond categories**: TWO WAY QUOTES and DDO/EDR restructured bonds
- **Incremental updates**: Appends new days to existing Excel without rebuilding
- **Data validation**: Alerts for missing data, high yields, negative yields
- **Formatted Excel output**: Color-coded headers, percentage formatting, borders
- **Email delivery**: Sends daily report via Resend
- **Historical archive**: Commits reports to `output/` folder

## Project Structure

```
.
├── .github/
│   └── workflows/
│       └── daily-bond-report.yml    # GitHub Actions workflow
├── output/
│   └── treasury_bond_yields.xlsx   # Generated reports (committed)
├── treasury_reports/               # Raw downloaded Excel files
├── extract_data.py                 # Main scraper script
├── pyproject.toml                  # Dependencies & metadata
├── requirements.txt                # Legacy requirements
└── README.md                       # This file
```

## Setup

### 1. Clone Repository

```bash
git clone https://github.com/BishanFer/bond-scrapper.git
cd bond-scrapper
```

### 2. Install Dependencies

```bash
pip install -e .
# or
pip install -r requirements.txt
```

### 3. Configure GitHub Secrets

Go to **Settings → Secrets and variables → Actions** and add:

| Secret | Value |
|--------|-------|
| `RESEND_API_KEY` | Your Resend API key |
| `TO_EMAIL` | Recipient email address |
| `FROM_EMAIL` | Sender email (use `onboarding@resend.dev` for testing) |

### 4. Enable GitHub Actions

Go to **Actions** tab → Enable workflows

## Usage

### Run Locally

```bash
# Full month export
python extract_data.py --year 2025 --month 12

# Fetch latest report only
python extract_data.py --today

# Append latest day to existing Excel (incremental)
python extract_data.py --incremental

# Custom output path
python extract_data.py --year 2025 --output my_report.xlsx
```

### Automated Daily Run

The GitHub Actions workflow runs daily at **6:00 PM Sri Lanka Time** (12:30 UTC):

1. Scrapes treasury.gov.lk for the latest Daily Summary Report
2. Downloads the Excel file
3. Extracts QuotesTBond data
4. Validates data and checks for anomalies
5. Appends to existing Excel (incremental) or rebuilds
6. Commits updated report to `output/` folder
7. Emails the report via Resend

### Manual Trigger

Go to **Actions → Daily Treasury Bond Report → Run workflow**

## Data Structure

### Output Excel Format

**Tab 1: TWO_WAY_QUOTES**
- Rows: Bond Number + Maturity Date
- Columns: Trading dates (chronological)
- Values: Buy Yield (as percentage)

**Tab 2: DDO_EDR_BONDS**
- Same structure for Domestic Debt Optimisation & External Debt Restructuring bonds

### Bond Data Extracted

| Field | Description |
|-------|-------------|
| Bond Number | Treasury bond series (e.g., `06.752026A`) |
| Maturity Date | Bond maturity date |
| Buy Yield | Average buying yield |
| Date of Trading | Report publication date |

## Validation & Alerts

The scraper automatically checks for:

- **Missing data**: No reports or empty sheets
- **Zero yields**: Potential data corruption
- **High yields**: >50% (suspicious)
- **Negative yields**: Data errors
- **New bonds**: Bonds not seen before

Issues are logged and included in the email summary.

## Cost

- **GitHub Actions**: Free (2,000 minutes/month)
- **Resend**: Free tier (3,000 emails/month)
- **Total**: $0

## Troubleshooting

### Workflow fails

1. Check **Actions → Daily Treasury Bond Report → Latest run**
2. Review logs for error messages
3. Common issues:
   - Treasury website down (auto-retries 3x)
   - No report published today (weekends/holidays)
   - Resend API key invalid

### No email received

1. Check spam/junk folder
2. Verify `TO_EMAIL` secret is correct
3. Check `FROM_EMAIL` is verified in Resend dashboard
4. Review workflow logs for email response

## License

MIT License - see LICENSE file

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes
4. Submit a pull request

## Support

- Open an issue: https://github.com/BishanFer/bond-scrapper/issues
- Email: [your-email]