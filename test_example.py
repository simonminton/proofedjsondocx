#!/usr/bin/env python3
"""
Test script demonstrating how to use the JSON to DOCX converter with track changes or comments programmatically.
"""

import json
from json_to_docx import process_json_to_docx

# Example JSON data
sample_data = [
    {
        "field_name": "test_section",
        "content": "<h1>Test Document</h1><p>This is a test paragraph with some <strong>bold text</strong>.</p><ul><li>Item 1</li><li>Item 2</li></ul>",
        "content_type": "html",
        "comments": [
            {
                "start_index": 10,
                "end_index": 25,
                "comment_content": "This is a replacement text",
                "commented_text": "test paragraph",
                "id": "1"
            }
        ]
    },
    {
        "field_name": "plain_text_section",
        "content": "This is plain text content that will be added as a paragraph.",
        "content_type": "text",
        "comments": [
            {
                "start_index": 5,
                "end_index": 15,
                "comment_content": "replacement text",
                "commented_text": "is plain text",
                "id": "2"
            }
        ]
    }
]

def main():
    """Demonstrate the JSON to DOCX conversion with both track changes and comments."""
    print("Creating test documents...")
    
    # Convert JSON to DOCX with track changes (default)
    process_json_to_docx(sample_data, "test_track_changes_output.docx", use_track_changes=True)
    print("✓ Track changes document created: test_track_changes_output.docx")
    
    # Convert JSON to DOCX with comments
    process_json_to_docx(sample_data, "test_comments_output.docx", use_track_changes=False)
    print("✓ Comments document created: test_comments_output.docx")
    
    print("\nTest documents created successfully!")
    print("• test_track_changes_output.docx - Use 'Track Changes' in Word Review tab to see revisions")
    print("• test_comments_output.docx - Comments are visible in Word's Review pane")

if __name__ == "__main__":
    main() 