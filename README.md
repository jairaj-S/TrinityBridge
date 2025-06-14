# Raw Material Supplier Analysis Tool

A comprehensive Python tool for evaluating suppliers based on price, tariffs, and maturity ratings to support procurement decisions. The tool provides reliability scoring with configurable weighting between price and quality factors.

## Features

- **Excel Integration**: Complete Excel-to-Excel workflow for business users
- **Tariff Calculations**: Automatic tariff application based on country of origin and material type
- **Price Scoring**: Inverse scoring system where lower prices receive higher scores
- **Maturity Assessment**: Support for both qualitative (High/Medium/Low) and quantitative (0.0-1.0) ratings
- **Reliability Scoring**: Weighted combination of price and maturity scores with configurable weights
- **Comprehensive Reports**: Multi-sheet Excel output with detailed analysis

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Create Sample Excel Template
```bash
python main.py --create-sample-excel template.xlsx
```

### 3. Edit the Excel file with your supplier data

### 4. Run Analysis
```bash
python main.py --input-excel template.xlsx --output-excel results.xlsx
```

## Usage Examples

**Basic Excel Analysis:**
```bash
python main.py --input-excel suppliers.xlsx --output-excel results.xlsx
```

**Custom Weight Analysis:**
```bash
python main.py --input-excel suppliers.xlsx --output-excel results.xlsx --price-weight 0.7 --maturity-weight 0.3
```

**Run Demonstration:**
```bash
python main.py --demo
```

**Create Sample Template:**
```bash
python main.py --create-sample-excel my_template.xlsx
```

## Excel File Format

### Input Requirements

**Config Sheet:**
- `raw_material`: Name of the material being analyzed
- `quantity_needed`: Total quantity required

**Suppliers Sheet:**
- `name`: Supplier company name (required)
- `price_per_rm`: Price per unit of raw material (required)
- `country_of_origin`: 2-letter country code (required)
- `maturity_rating_qualitative`: High/Medium/Low (optional)
- `maturity_rating_quantitative`: 0.0-1.0 decimal value (optional)

### Output Sheets

The analysis generates an Excel file with multiple sheets:

1. **Summary**: Key results ranked by reliability score
2. **Detailed Analysis**: Complete breakdown of all metrics
3. **Cost Comparison**: Best vs worst options analysis
4. **Analysis Info**: Metadata and parameters used

## Tariff System

Current tariff rates configured for:
- Chinese Steel: 25%
- US Steel: 5%
- German Steel: 7%
- Indian Steel: 20%
- Japanese Steel: 8%
- Chinese Aluminum: 15%
- US Aluminum: 3%
- German Aluminum: 5%
- Default fallback: 0%

## Scoring Methodology

### Price Scoring
- Uses inverse scoring methodology (lower price = higher score)
- Normalized to 0-100 scale where 100 is the best (lowest) price

### Maturity Scoring
- Qualitative: High=100, Medium=50, Low=20
- Quantitative: Supports 0.0-1.0 decimal scale or 0-100 direct scale
- Default: 50 (neutral) if not specified

### Reliability Scoring
- Weighted combination: (price_score × price_weight) + (maturity_score × maturity_weight)
- Default weights: 60% price, 40% maturity
- Fully configurable via command-line arguments

## Command-Line Options

```
--input-excel INPUT_EXCEL         Path to input Excel file
--output-excel OUTPUT_EXCEL       Path for output Excel file
--price-weight PRICE_WEIGHT       Weight for price score (default: 0.6)
--maturity-weight MATURITY_WEIGHT Weight for maturity score (default: 0.4)
--create-sample-excel FILE        Create sample Excel template
--demo                            Run demonstration with sample data
--help                           Show help message
```

## Dependencies

- **pandas**: Excel file processing and data manipulation
- **openpyxl**: Excel file writing engine

## Error Handling

The tool provides comprehensive error handling for:
- Missing or invalid Excel files
- Incorrect file formats
- Invalid weight configurations
- Missing required columns
- Data validation errors

## License

This tool is provided as-is for procurement analysis purposes.