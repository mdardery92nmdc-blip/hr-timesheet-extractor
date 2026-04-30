# 📊 HR Timesheet Extractor Pro

A web-based tool to extract attendance data from PDF timesheets and generate Comp-Off & Leave reports.

## Features

- 📄 **PDF Text Extraction**: Extracts employee info and attendance codes from timesheet PDFs
- 🔍 **OCR Fallback**: Uses EasyOCR when text extraction fails
- 📊 **Comp-Off Calculation**: Automatically calculates comp-off based on contractual days
- 📋 **Leave Tracking**: Tracks leave days with configurable leave codes
- 💾 **Excel Export**: Formatted Excel output with weekday headers
- 📱 **Web Interface**: Clean Streamlit interface accessible from any device

## Deployment

### Option 1: Render (Recommended - Free Tier)

1. Fork/clone this repository to GitHub
2. Go to [render.com](https://render.com) and sign up
3. Click "New +" → "Web Service"
4. Connect your GitHub repository
5. Render will auto-detect `render.yaml` and configure everything
6. Click "Create Web Service"

**Note**: Free tier spins down after 15 minutes of inactivity (30-50s cold start).

### Option 2: Railway (Free $5 Credit)

1. Push code to GitHub
2. Go to [railway.app](https://railway.app)
3. Click "New Project" → "Deploy from GitHub repo"
4. Railway auto-detects Python and deploys

### Option 3: Local Development

```bash
# Clone repository
git clone <your-repo-url>
cd hr-timesheet-extractor

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run app
streamlit run app.py
```

## File Structure

```
.
├── app.py                    # Main Streamlit application
├── attendance_analysis.py    # Analysis engine
├── requirements.txt          # Python dependencies
├── render.yaml              # Render deployment config
├── Procfile                 # Process file for Railway/Heroku
├── runtime.txt              # Python version specification
└── README.md                # This file
```

## Input Format

### Contract File (CSV/Excel)
Must contain columns:
- `Employee #`: Employee ID number
- `Contractual Days Per Week`: 5, 5.5, 6, or 7

### Timesheet PDFs
Should contain:
- Employee Name, Number, Designation, Company
- Attendance code row with day columns (1-31)

**Supported Attendance Codes:**
- `W` = Full day work (1.0)
- `WHF` = Work from home full (1.0)
- `WHH` = Work from home half (0.5)
- `H` = Holiday (1.0)
- `HD` = Half day (0.5)
- `CO` = Comp-off (0.0)
- `S` = Sick leave (0.0)
- `L` = Leave (0.0)
- `U` = Unpaid (0.0)
- `OFF` = Off day (0.0)

## Environment Variables

No environment variables required for basic operation.

## License

MIT License - Free for personal and commercial use.
