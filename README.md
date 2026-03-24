# Third Sector PDF Indicator Extractor
## Project Overview
This project automates the extraction of 40 specific indicators from 44 Italian Third Sector project reports. It uses Docling for layout-aware PDF parsing and GPT-4o for structured data extraction.

## Pipeline
1. Ingestion: Raw PDFs are placed in data/raw/.
2. Parsing: PDFs are converted to Markdown to preserve document hierarchy and table structures.
3. Extraction: Markdown text is sent to OpenAI with a structured prompt to identify indicators.
4. Export: Results are validated and compiled into a final extracted_data.csv.

## Setup
1. Clone the repo.
2. Create a virtual environment: python -m venv venv.
3. Install dependencies: pip install -r requirements.txt.
4. Add your OPENAI_API_KEY to a .env file.

## Usage
1. Run python src/utils/pdf_parser.py to generate Markdown files.
2. Run python src/extractor.py to perform the LLM extraction.