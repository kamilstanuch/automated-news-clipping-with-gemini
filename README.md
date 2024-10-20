# Gmail Newsletters Parser for Google Sheets via Gemini-1.5-Flash API 

This Google Apps Script automatically extracts news articles from your Gmail newsletters and organizes them into a Google Sheet with Gemini-1.5-Flash API. 

## Features

- Fetches emails from a specific Gmail label.
- Uses the Gemini API to intelligently extract news articles.
- Categorizes news articles based on their content.
- Saves extracted news and email metadata to separate Google Sheets.
- Includes error handling and retry logic for reliable API interaction.
- Shortens long URLs to avoid Google Sheet cell limits and expands them before saving.

## How it Works

1. **Email Fetching:** The script searches your Gmail inbox for emails with a specified label.
2. **Content Extraction:** It extracts the content of each email, prioritizing HTML content and converting it to Markdown for better readability.
3. **News Analysis (Gemini API):** The script sends the email content to the Gemini API, which uses a powerful prompt to identify and extract news articles.
4. **News Validation:**  Extracted news items are validated to ensure they have required fields like titles, descriptions, and links.
5. **Data Storage:** The script saves the extracted news articles and relevant email metadata (sender, date, etc.) to designated Google Sheets. 

## How to Use

1. **Copy the Code:** Copy the entire code from this repository.
2. **Open Google Apps Script:** 
   - Go to your Google Sheet.
   - Click "Tools" > "Script editor".
3. **Paste and Configure:** 
   - Paste the copied code into the script editor.
   - **Important:** Replace the following:
      - `'YOUR_LABEL_NAME'` (in the `CONFIG` object) with the Gmail label you use for newsletters.
      - `'YOUR_SPREADSHEET_ID'` (in the `CONFIG` object) with the ID of your Google Sheet.
      - `'YOUR_GEMINI_API_KEY'` (inside the `generateContent` function) with your Gemini API key.
4. **Set up API Key (Recommended):**
   - Go to "File" > "Project properties" > "Script properties".
   - Add a property named `GEMINI_API_KEY` and set its value to your API key. This is more secure than storing the key directly in the code.
5. **Run the Script:** 
   - Save the script.
   - Run the `processNewEmails()` function. 
   - You may need to authorize the script to access your Gmail and Google Sheets.

## Important Notes

- **API Key:** Protect your Gemini API key! 
- **Rate Limits:** Be aware of Gemini API rate limits. The script has retry logic, but adjust the settings (`MAX_RETRIES`, `RETRY_DELAY`) if needed. 
- **Customization:** The prompt, categories, and data formatting can be customized to your needs.

Let me know if you have any other questions.
