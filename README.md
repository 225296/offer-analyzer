# Offer Letter Analyzer

A smart tool to compare your previous and current offer letters, extract key salary components, and calculate your salary hike percentage.

## Features

✨ **Smart Extraction** - Automatically detects:
- Total CTC / Annual Package
- Base / Fixed Salary
- Variable Pay / Performance Bonus
- Joining / Sign-on Bonus
- Stock / RSU / ESOP (annual value)

📊 **Smart Formatting** - Handles multiple formats:
- Indian notation (₹, Rs., LPA, Lakhs, Crores)
- USD, EUR, GBP, AED
- Monthly figures (auto-annualized)
- Scanned PDFs, Word documents, plain text

🔐 **Private** - All processing happens locally in your browser/session. Files are never uploaded.

📈 **Detailed Comparison** - Shows:
- Overall hike percentage
- Annual and monthly increases
- Component-wise breakdown
- Benefit changes

## Usage

### Option 1: Streamlit Web App (Recommended)

```bash
streamlit run app.py
```

Then open `http://localhost:8501` in your browser.

### Option 2: Standalone HTML

Open `offer-analyzer.html` directly in any browser (no server needed).

## Files

- `app.py` - Streamlit web application
- `offer-analyzer.html` - Standalone HTML/JavaScript version
- `requirements.txt` - Python dependencies
- `previous-offer.txt` - Sample previous offer letter (for testing)
- `current-offer.txt` - Sample current offer letter (for testing)

## Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Run Streamlit app
streamlit run app.py
```

## Deployment to Streamlit Cloud

1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io/)
3. Click "New app"
4. Connect to this GitHub repo
5. Select `app.py` as the main file
6. Deploy!

### Or Deploy in 3 Steps:

```bash
# 1. Create GitHub repo (if not already created)
git init
git add .
git commit -m "Initial commit: Offer Letter Analyzer"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/offer-analyzer.git
git push -u origin main

# 2. Go to https://share.streamlit.io/
# 3. Paste your GitHub repo URL and deploy!
```

## Currency Support

- 🇮🇳 INR (Indian Rupee)
- 🇺🇸 USD (US Dollar)
- 🇪🇺 EUR (Euro)
- 🇬🇧 GBP (British Pound)
- 🇦🇪 AED (UAE Dirham)

## How It Works

1. **Upload** your previous and current offer letters (PDF, DOCX, or TXT)
2. **Verify** the extracted values and edit if needed
3. **Calculate** to get your hike percentage, component comparison, and detailed breakdown

## Supported Offer Formats

The parser handles:
- Professional offer letters with detailed components
- Indian-style offers with HRA, DA, Special Allowance
- LPA notation (e.g., "32.5 LPA")
- Lakh/Crore notation (e.g., "12,00,000" or "1.2 Cr")
- USD-style annual packages
- Monthly salary with benefits
- Table-formatted offers
- Scanned PDFs with OCR text

## Testing

Use the included sample offer letters to test:

```bash
# These files are included for testing
- previous-offer.txt (₹28,44,000 CTC)
- current-offer.txt (₹42,37,500 CTC)
```

Load these in the app to see a ~49% hike comparison.

## Privacy & Security

- ✅ All processing happens in your browser (HTML version) or session (Streamlit)
- ✅ No files are stored on any server
- ✅ No data is sent to external APIs
- ✅ Open source - you can audit the code

## Limitations

- PDF text extraction works best with selectable text (not scanned images)
- For scanned PDFs, use the text-paste option
- Very unconventional offer formats may require manual field entry
- Currency detection is automatic but can be overridden

## Contributing

Found a bug? Have a feature request? Create an issue!

## License

MIT - Feel free to use and modify as needed.

---

Made with ❤️ for career decisions
