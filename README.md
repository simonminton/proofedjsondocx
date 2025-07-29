# PROOFED Example: JSON to DOCX Converter with Track Changes or Comments

This Python script converts JSON data containing content and comments into a Microsoft Word document (.docx) with either track changes or comments inserted at specified positions.

## Features

- Converts HTML content to properly formatted Word document paragraphs
- Supports headings (H1-H6), paragraphs, and lists (ordered and unordered)
- Inserts either track changes or comments at specified start and end positions within text
- **Track Changes Mode**: Shows original text as deletions and comment content as insertions
- **Comments Mode**: Adds traditional Word comments with comment bubbles
- Handles both HTML and plain text content types
- Preserves document structure and formatting

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Command Line Options

The script supports two modes via command-line flags:

- `--trackchanges` (default): Creates track changes with redlined revisions
- `--comments`: Creates traditional Word comments

### Method 1: Using a JSON file

**Track Changes (default):**
```bash
python json_to_docx.py input.json output.docx
# or explicitly
python json_to_docx.py input.json output.docx --trackchanges
```

**Comments:**
```bash
python json_to_docx.py input.json output.docx --comments
```

### Method 2: Piping JSON content
```bash
echo 'JSON_CONTENT' | python json_to_docx.py - output.docx --comments
```

### Method 3: Using the shell script
```bash
./run.sh input.json output.docx --comments
```

## JSON Format

The script expects JSON data in the following format:

```json
[
  {
    "field_name": "string",
    "content": "string (HTML or plain text)",
    "content_type": "html|text",
    "comments": [
      {
        "start_index": 0,
        "end_index": 10,
        "comment_content": "Comment or replacement text",
        "commented_text": "Text being commented on or replaced",
        "id": "unique_id"
      }
    ]
  }
]
```

## Track Changes Mode

When using `--trackchanges` (default), the script creates track changes that show:
- **Deletions**: Original text marked as deleted (strikethrough)
- **Insertions**: Comment content marked as inserted (underlined)
- **Author**: All changes are attributed to "Reviewer"
- **Timestamp**: Changes include the current date/time

## Comments Mode

When using `--comments`, the script creates traditional Word comments that show:
- **Comment Bubbles**: Comments appear in the margin or as popups
- **Highlighted Text**: The commented text is highlighted
- **Comment Content**: The `comment_content` appears as the comment text
- **Author**: Comments are attributed to "Reviewer"

## Output

The script generates a Word document (.docx) with:
- Properly formatted content based on HTML structure
- Either track changes or comments inserted at the specified positions
- Field names as section headings
- Preserved formatting and structure

## Notes

- Comments/track changes are inserted at the character level within paragraphs
- HTML content is parsed and converted to appropriate Word document elements
- The script handles basic HTML elements: p, h1-h6, ul, ol, li
- Comments/track changes are only applied to the first matching paragraph for each comment range
- **For Track Changes**: Enable "Track Changes" in the Review tab of Word to see the revisions
- **For Comments**: Comments are visible by default in Word's Review pane 