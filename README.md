# RVtools Excel Processor

This application processes RVtools Excel files and creates a new ServerList tab with specific information extracted from the vInfo tab.

## Features

- Upload RVtools Excel files
- Extract specific columns from the vInfo tab
- Convert memory and disk measurements from MiB to GiB
- Add new columns for scope and notes
- Download the processed Excel file

## Installation

1. Clone this repository
2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
streamlit run app.py
```

2. Open your web browser and navigate to the URL shown in the terminal (typically http://localhost:8501)
3. Upload your RVtools Excel file using the file uploader
4. View the processed data in the web interface
5. Download the processed Excel file using the download button

## Output Columns

The processed Excel file will contain the following columns:
- VM Name
- Powerstate
- CPUs
- Memory (GiB)
- Provisioned Disk (GiB)
- In Use Disk (GiB)
- Cluster
- OS According to the configuration file
- In Scope for Prod?
- In Scope for DR?
- Notes 