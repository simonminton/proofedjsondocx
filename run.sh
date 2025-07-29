#!/bin/bash

# JSON to DOCX Converter Runner Script

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install dependencies if needed
if [ ! -f "venv/lib/python*/site-packages/docx" ]; then
    echo "Installing dependencies..."
    pip install -r requirements.txt
fi

# Run the converter with all arguments passed through
if [ $# -eq 0 ]; then
    echo "Usage: ./run.sh <input_json_file> [output_docx_file] [--trackchanges|--comments]"
    echo "Examples:"
    echo "  ./run.sh sample_data.json"
    echo "  ./run.sh sample_data.json output.docx --comments"
    echo "  ./run.sh sample_data.json output.docx --trackchanges"
    exit 1
fi

python json_to_docx.py "$@" 