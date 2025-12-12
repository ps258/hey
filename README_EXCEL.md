# Excel Report Generator for Hey Load Testing Results

This Python script processes CSV output from load testing and generates an Excel spreadsheet with charts showing response times and requests per second.

## Features

- **Data Sheet**: Contains all raw data from the CSV file
- **Statistics Sheet**: Summary statistics including:
  - Mean, min, and max for average response times
  - Mean, min, and max for P95 response times
  - Mean, min, and max for requests per second
- **Charts Sheet**: Visual representations including:
  - Line chart showing average and P95 response times over time
  - Line chart showing requests per second over time

## Installation

Install the required Python package:

```bash
pip install -r requirements.txt
```

Or install directly:

```bash
pip install openpyxl
```

## Usage

### Basic Usage

```bash
python process_results.py input.csv output.xlsx
```

### Example with Sample Data

```bash
python process_results.py sample_data.csv results.xlsx
```

## CSV Format

The script expects a CSV file with the following columns:

```
timestamp,avg_response_time,p95_response_time,requests_per_second
```

- **timestamp**: ISO 8601 timestamp or any string identifier
- **avg_response_time**: Average response time in seconds (float)
- **p95_response_time**: 95th percentile response time in seconds (float)
- **requests_per_second**: Number of requests per second (float)

## Example CSV

```csv
timestamp,avg_response_time,p95_response_time,requests_per_second
2025-12-12T13:00:00,0.0234,0.0456,425.5
2025-12-12T13:00:01,0.0245,0.0478,418.2
2025-12-12T13:00:02,0.0229,0.0445,432.1
```

## Output

The generated Excel file contains three sheets:

1. **Data**: Raw data with formatted headers
2. **Statistics**: Calculated summary statistics
3. **Charts**: Two line charts visualizing the performance metrics

## Script Features

- Automatic column width adjustment
- Styled headers with colors
- Error handling for missing files and invalid data
- Summary statistics printed to console
- Professional chart formatting

## Requirements

- Python 3.6+
- openpyxl 3.1.0+

## Error Handling

The script handles:
- Missing input files
- Invalid CSV format
- Missing required columns
- Invalid numeric values
- File permission errors when writing output

## Example Output

```
Reading data from: sample_data.csv
Processing 10 samples...
Generating Excel file: results.xlsx
Successfully created Excel file: results.xlsx
  - Data sheet with 10 samples
  - Statistics sheet with summary metrics
  - Charts sheet with visualizations

Summary Statistics:
  Average Response Time: 0.0240s
  P95 Response Time: 0.0467s
  Average RPS: 421.30