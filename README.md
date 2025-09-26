# BOM Tree Generator - Web Application

A web-based Bill of Materials (BOM) analysis tool that creates hierarchical BOM trees from multiple data sources. This application allows users to analyze BOM structures without requiring Python installation.

## Features

- 🌳 **Hierarchical BOM Analysis**: Build complete BOM trees with all levels
- 📁 **Multiple File Support**: Upload Excel, CSV, or TXT BOM files
- 🔍 **Part Search**: Analyze specific parts across multiple BOM sources
- 📊 **Excel Report Generation**: Download comprehensive Excel reports with hyperlinks
- 🌐 **Web-based Interface**: No Python installation required for end users

## Quick Start

### Option 1: Run Locally (Development)

1. **Install Python** (if not already installed)
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the application**:
   ```bash
   streamlit run app.py
   ```
4. **Open your browser** to `http://localhost:8501`

### Option 2: Deploy to Cloud (Production)

Choose one of these cloud platforms for easy deployment:

#### Streamlit Cloud (Recommended - Free)
1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repository
4. Deploy with one click!

#### Heroku
1. Create a `Procfile` with: `web: streamlit run app.py --server.port=$PORT --server.address=0.0.0.0`
2. Deploy using Heroku CLI or GitHub integration

#### Railway
1. Connect your GitHub repository
2. Railway will automatically detect Streamlit and deploy

## Usage

1. **Upload BOM Files**: Use the sidebar to upload your BOM files (Excel, CSV, or TXT)
2. **Enter Part Numbers**: Add part numbers manually or upload a file containing them
3. **Select Sources**: Choose which BOM sources to search
4. **Generate Report**: Click "Analyze Parts" to generate and download the Excel report

## File Format Requirements

Your BOM files must contain these columns:
- `Product no`: Parent part number
- `Component no`: Component part number
- Additional columns for names, descriptions, etc. are supported

## Supported File Formats

- **Excel**: `.xlsx`, `.xls`
- **CSV**: `.csv` (with automatic encoding detection)
- **Text**: `.txt` (with automatic delimiter detection)

## Deployment Options

### Free Hosting Options

1. **Streamlit Cloud** (Recommended)
   - Free tier available
   - Automatic deployments from GitHub
   - No server management required

2. **Railway**
   - Free tier with usage limits
   - Easy GitHub integration
   - Automatic deployments

3. **Heroku**
   - Free tier discontinued, but paid plans available
   - Reliable and well-documented

### Self-Hosted Options

1. **Docker** (for advanced users)
2. **VPS/Cloud Server** (AWS, Google Cloud, Azure)
3. **Local Network** (for internal company use)

## Technical Details

- **Framework**: Streamlit
- **Dependencies**: pandas, openpyxl, chardet
- **Output**: Excel files with hyperlinks and formatting
- **Encoding**: Automatic detection for text files

## Troubleshooting

### Common Issues

1. **File Upload Errors**: Ensure your BOM files have the required columns (`Product no`, `Component no`)
2. **Encoding Issues**: The app automatically detects file encoding, but you can try saving files as UTF-8
3. **Large Files**: For very large BOM files, consider splitting them or using a more powerful server

### Getting Help

- Check the file format requirements
- Ensure your BOM files have the correct column names
- Try with smaller files first to test the functionality

## Original Script

This web application is based on the original Python script `bom.py` which provides the same functionality through a command-line interface.
