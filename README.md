# RVtools Excel Processor

A Streamlit application that processes RVtools Excel files to create a standardized ServerList tab with additional scope tracking and summary statistics.

## Features

- Processes RVtools Excel files (vInfo tab)
- Creates a standardized ServerList tab with:
  - VM Name
  - Powerstate
  - CPUs
  - Memory (GB)
  - Provisioned Disk (GB)
  - In Use Disk (GB)
  - Cluster
  - OS
  - Production Scope tracking
  - DR Scope tracking
  - Notes
- Generates a Summary tab with:
  - Powerstate statistics
  - Operating System statistics
  - Production Scope summary
  - DR Scope summary
  - Subtotals for each section

## Requirements

- Python 3.8+
- pandas
- openpyxl
- streamlit

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/rvtools-processor.git
cd rvtools-processor
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the Streamlit app:
```bash
streamlit run app.py
```

2. Upload your RVtools Excel file through the web interface
3. Download the processed Excel file with the new ServerList and Summary tabs

## Scope Values

The application recognizes the following values as "In Scope":
- yes
- true
- 1
- X
- y

Any other value is considered "Not In Scope". 