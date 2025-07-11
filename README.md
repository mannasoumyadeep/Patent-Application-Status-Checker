# Patent Application Status Checker

A Streamlit web application that automates checking patent application statuses from the Indian Patent Office website.

## Features

- 📤 Upload Excel files containing application numbers
- 🔍 Automated status checking with CAPTCHA bypass
- 📊 Real-time progress tracking
- 💾 Download results in Excel format
- 🔄 Retry failed applications
- ⏹️ Stop processing mid-way and save partial results
- 🚀 Concurrent processing (5 workers)

## Deployment on Streamlit Cloud

### Prerequisites

1. A GitHub account
2. A Streamlit Cloud account (free at [share.streamlit.io](https://share.streamlit.io))

### Step-by-Step Deployment

1. **Fork or Create Repository**
   - Create a new GitHub repository
   - Upload all the files:
     - `app.py` (the main Streamlit application)
     - `requirements.txt`
     - `packages.txt`
     - `.streamlit/config.toml`

2. **Deploy on Streamlit Cloud**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Click "New app"
   - Select your repository
   - Set the main file path to `app.py`
   - Click "Deploy"

3. **Wait for Deployment**
   - The app will take a few minutes to deploy
   - Streamlit Cloud will install all dependencies automatically

## Local Development

### Requirements

- Python 3.8+
- Chrome browser installed

### Installation

```bash
# Clone the repository
git clone <your-repo-url>
cd patent-status-checker

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Running Locally

```bash
streamlit run app.py
```

## Usage

1. **Prepare Input File**
   - Create an Excel file (.xlsx or .xls)
   - Put application numbers in the first column
   - No header required

2. **Upload and Process**
   - Upload the Excel file
   - Click "Start Processing"
   - Monitor the progress bar
   - Stop anytime if needed

3. **Download Results**
   - Download all results (includes successful and failed)
   - Download error report separately
   - Retry failed applications if needed

## File Structure

```
patent-status-checker/
├── app.py                    # Main Streamlit application
├── requirements.txt          # Python dependencies
├── packages.txt             # System dependencies for Streamlit Cloud
├── .streamlit/
│   └── config.toml          # Streamlit configuration
└── README.md                # This file
```

## Important Notes

- The application uses Chrome in headless mode
- CAPTCHA is bypassed using the audio CAPTCHA API
- Processing speed depends on the Patent Office website response time
- Failed applications can be retried individually
- Results maintain the original order from the input file

## Troubleshooting

### Common Issues on Streamlit Cloud

1. **Chrome/Selenium Issues**
   - The app is configured to use system Chrome on Streamlit Cloud
   - No additional configuration needed

2. **Memory Issues**
   - If processing large files, consider splitting them
   - The app uses concurrent processing which may consume more memory

3. **Timeout Issues**
   - The Patent Office website might be slow
   - Failed applications will be automatically retried up to 3 times

### Support

For issues or questions:
- Check the Streamlit Cloud logs
- Ensure all files are properly uploaded to GitHub
- Verify the packages.txt file includes chromium dependencies

## License

This project is for educational and personal use. Please respect the Patent Office website's terms of service.

## Disclaimer

This tool automates the process of checking patent application statuses. Users are responsible for:
- Ensuring compliance with the Patent Office website's terms of service
- Verifying the accuracy of retrieved data
- Using the tool responsibly and not overloading the server