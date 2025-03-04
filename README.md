# Google Sheets OpenAI Translator

A Google Apps Script tool that enables translation of words in Google Sheets using OpenAI's GPT-4 API. This tool provides a simple interface for translating words with optional context support.

## Features

- Easy-to-use menu interface in Google Sheets
- Support for source and target language specification
- Optional context column for more accurate translations
- Progress tracking during translation
- Error handling and logging
- Rate limiting to avoid API throttling

## Prerequisites

- A Google Sheets account
- An OpenAI API key
- Basic understanding of Google Apps Script

## Setup

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Create a new script file and paste the contents of `translate.js`
4. Replace the empty `OPENAI_API_KEY` constant with your OpenAI API key
5. Save the script
6. Refresh your Google Sheet

## Usage

1. In your Google Sheet:

   - Enter the source language in cell A1 (e.g., "English")
   - Enter the target language in cell B1 (e.g., "Spanish")

2. Select the column containing the words you want to translate

3. Click on "Translation Tools" in the menu bar and select "Translate Words with OpenAI"

4. When prompted, indicate whether you have a context column (optional)

5. The script will:
   - Translate each word in the selected column
   - Place translations in the next available column
   - Show progress and completion status

## Context Support

The tool supports an optional context column to provide more accurate translations. If enabled:

- Words to translate should be in the selected column
- Context should be in the column immediately to the right
- Translations will appear in the third column

## Error Handling

- The script includes error handling for API issues
- Failed translations will be marked with "ERROR:" followed by the error message
- Errors are logged in the Apps Script execution log

## Rate Limiting

The script includes a 200ms delay between translations to avoid hitting OpenAI's rate limits.

## Security Note

⚠️ Never share your OpenAI API key or commit it to version control. The key should be kept private and secure.

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
