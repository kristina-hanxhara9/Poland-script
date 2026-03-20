# Poland Business PKD Code Analysis

Analyzes Polish businesses to determine which PKD (Polska Klasyfikacja Działalności) codes correspond to specific channel categories: **DIY shops**, **Paint specialists**, and **Builders merchants**.

## How It Works

1. Reads an Excel file with business data (`name`, `zip`, `channel` columns)
2. Looks up each business via Poland government APIs (CEIDG / GUS REGON) to get their official PKD codes
3. Analyzes which PKD codes correspond to each channel category
4. Performs keyword matching on business names (Polish + English terms)
5. Produces an enriched Excel output + analysis report

## Setup

```bash
pip install -r requirements.txt
```

## Usage

### Basic (keyword matching only, no API key needed)
```bash
python main.py your_data.xlsx --output results.xlsx
```

### With API keys
```bash
# Using CEIDG API
export CEIDG_API_KEY="your-key-here"
python main.py your_data.xlsx

# Using GUS REGON API
export GUS_API_KEY="your-key-here"
python main.py your_data.xlsx

# Using GUS sandbox (test data)
python main.py your_data.xlsx --sandbox
```

### Try with sample data
```bash
python create_sample.py
python main.py sample_input.xlsx
```

## Input Format

Excel file (.xlsx) with these columns:

| Column  | Description                                              |
|---------|----------------------------------------------------------|
| name    | Business name                                            |
| zip     | Polish postal code (XX-XXX format)                       |
| channel | Category: "DIY shops", "Paint specialists", "Builders merchants" |

## Getting API Keys

### CEIDG API (dane.biznes.gov.pl)
1. Go to https://dane.biznes.gov.pl
2. Register and verify your identity
3. Generate an API key in the developer portal

### GUS REGON/BIR API
1. Go to https://api.stat.gov.pl/Home/RegonApi
2. Register for a user account
3. Request an API key (free of charge)

## Key PKD Codes

| Category            | PKD Code | Description                                        |
|---------------------|----------|----------------------------------------------------|
| DIY shops           | 47.52.Z  | Retail sale of hardware, paints and glass           |
| DIY shops           | 47.53.Z  | Retail sale of carpets, rugs, wall/floor coverings  |
| Paint specialists   | 47.52.Z  | Retail sale of hardware, paints and glass           |
| Paint specialists   | 20.30.Z  | Manufacture of paints, varnishes, coatings          |
| Builders merchants  | 46.73.Z  | Wholesale of wood, construction materials           |
| Builders merchants  | 46.74.Z  | Wholesale of hardware, plumbing, heating equipment  |

## Output

- **results.xlsx** — enriched data with PKD codes, descriptions, keyword matches, confidence scores
- **report.txt** — full analysis report including PKD distribution by channel, cross-channel analysis, and keyword matching results
