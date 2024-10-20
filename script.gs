// Configuration for the script
const CONFIG = {
  LABEL_NAME: 'YOUR_LABEL_NAME',  // Replace with the label name used for news emails
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID', // Replace with the ID of your Google Sheet
  BUFFER_TIME_MINUTES: 10 // Buffer time to avoid processing duplicates
};

// Prefix for shortened links 
const LINK_PREFIX = 'https://short.link/'; 

/**
 * Main function to process new emails, extract news, and save to sheets.
 */
function processNewEmails() {
  console.log('Starting to process new emails...');

  // Get the spreadsheet and sheets
  var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var emailSheet = spreadsheet.getSheetByName("Emails");
  var newsSheet = spreadsheet.getSheetByName("News");

  // Check if the sheets exist
  if (!emailSheet || !newsSheet) {
    console.error("Error: Required sheets ('Emails' and 'News') not found.");
    return;
  }

  // Ensure correct columns in sheets
  ensureCorrectColumns(emailSheet);
  ensureCorrectColumns(newsSheet);

  // Get the timestamp of the last processed email
  var lastProcessedTime = getLastProcessedTime(emailSheet);

  // Build the Gmail search query
  var searchQuery = buildSearchQuery(lastProcessedTime);
  if (!searchQuery) {
    return; // No need to proceed if the query is null
  }

  console.log("Search query: " + searchQuery);

  // Get threads (emails) matching the search query
  var threads = GmailApp.search(searchQuery);
  console.log(`Found ${threads.length} email threads`);

  // Process each email thread
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    console.log(`Processing thread ${i + 1} of ${threads.length}, containing ${messages.length} messages`);

    // Process each message in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      console.log(`Processing message ${j + 1} (ID: ${message.getId()})`);
      processEmail(message, emailSheet, newsSheet, i);

      // Save progress after each email
      SpreadsheetApp.flush();
    }
  }

  console.log('Finished processing new emails.');
}

/**
 * Processes a single email: extracts content, analyzes for news, saves data.
 * @param {GmailMessage} message The Gmail message object.
 * @param {Sheet} emailSheet The sheet to store email metadata.
 * @param {Sheet} newsSheet The sheet to store extracted news.
 * @param {number} processedCount The count of processed emails. 
 */
function processEmail(message, emailSheet, newsSheet, processedCount) {
  // Extract email data
  var subject = message.getSubject();
  var plainContent = message.getPlainBody();
  var htmlContent = message.getBody();

  // Prefer HTML content, convert to Markdown if available
  var content = htmlContent ? convertHtmlToMarkdown(htmlContent) : plainContent;

  // Truncate content for API limits
  var contentShort = truncateContent(content);

  // Prepare email metadata
  var emailDate = message.getDate();
  var sender = message.getFrom();
  var fetchedTime = new Date();
  var fetchedTimeUnix = Math.floor(fetchedTime.getTime() / 1000);
  var fetchStatus = 'success'; 
  var emailId = message.getId();
  var { shortenedContent, linkMap } = shortenLinks(contentShort); 

  var emailMetadata = {
    email_date: emailDate,
    sender: sender,
    fetched_time: fetchedTime,
    linkMap: linkMap,
    emailId: emailId
  };

  // Log email processing details
  console.log(`Processing email:`);
  console.log(`- Subject: ${subject}`);
  console.log(`- Date: ${emailDate.toISOString()}`);
  console.log(`- Sender: ${sender}`);
  console.log(`- Email ID: ${emailId}`);
  console.log(`- Fetched Time: ${fetchedTime.toISOString()} (Unix: ${fetchedTimeUnix})`);
  console.log(`- Content Type: ${htmlContent ? 'HTML (converted to Markdown)' : 'Plain Text'}`);

  // Analyze email content for news
  console.log(`Analyzing email ${processedCount + 1}`);
  var geminiStatus, newsCount, rawResponse;

  try {
    var result = GeminiAPI.analyzeEmailContent(content, shortenedContent, emailMetadata);
    geminiStatus = result.status;
    newsCount = result.newsCount;
    rawResponse = result.rawResponse;

    // If news extraction was successful, validate and add to sheet
    if (result.status === 'success' && result.newsItems) {
      var validatedNews = GeminiAPI.validateNewsItems(result.newsItems);
      var expandedNews = GeminiAPI.expandShortLinks(validatedNews, linkMap);
      newsCount = GeminiAPI.addNewsToSheet(expandedNews, emailMetadata, content, shortenedContent);
    }
  } catch (error) {
    // Handle errors during news analysis
    console.error(`Error analyzing email ${processedCount + 1}:`, error);
    geminiStatus = 'failure';
    newsCount = 0;
    rawResponse = GeminiAPI.truncateContent(error.toString());
  }

  try {
    // Add processed email data to the "Emails" sheet
    var newRow = [
      subject,
      GeminiAPI.truncateContent(content),
      GeminiAPI.truncateContent(shortenedContent),
      emailDate,
      sender,
      fetchedTime,
      fetchedTimeUnix,
      fetchStatus,
      geminiStatus,
      newsCount,
      GeminiAPI.truncateContent(rawResponse),
      JSON.stringify(linkMap),
      emailId
    ];

    emailSheet.insertRowAfter(1); 
    emailSheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);

    console.log(`Added email to Emails sheet below the header. Status: ${geminiStatus}, News Count: ${newsCount}`);
  } catch (error) {
    console.error(`Error adding email ${processedCount + 1} to sheet:`, error);
  }
}

/**
 * Ensures the sheet has the correct columns, adding them if necessary.
 * @param {Sheet} sheet The Google Sheet to check and update.
 */
function ensureCorrectColumns(sheet) {
  var sheetName = sheet.getName();
  var expectedHeaders = [];

  if (sheetName === "Emails") {
    expectedHeaders = ["subject", "content", "content_short", "email_date", "sender", "fetched_time", "fetched_time_unix", "fetch_status", "gemini_status", "news_count", "raw_gemini_response", "link_map", "email_id"]; 
  } else if (sheetName === "News") {
    expectedHeaders = ["Title", "Description", "Link", "Category", "Email Date", "Sender", "Fetched Time", "Email ID"];
  } else {
    console.error(`Unexpected sheet name: ${sheetName}`);
    return;
  }

  // If the sheet is empty, add all headers
  if (sheet.getLastColumn() === 0) { 
    console.log(`${sheetName} sheet is empty. Adding all expected headers.`);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return;
  }

  var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var missingHeaders = expectedHeaders.filter(header => !currentHeaders.includes(header));

  // Add missing columns
  missingHeaders.forEach(header => {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()).setValue(header);
    console.log(`Added missing column to ${sheetName} sheet: ${header}`);
  });
}

/**
 * Gets the timestamp of the last processed email from the Emails sheet.
 * @param {Sheet} sheet The "Emails" sheet.
 * @return {number} The Unix timestamp (in seconds) of the last processed email.
 */
function getLastProcessedTime(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    return sheet.getRange(lastRow, 7).getValue(); // Assuming column 7 is fetched_time_unix
  }
  // If no previous emails, use a default start time (e.g., 24 hours ago)
  return Math.floor(Date.now() / 1000) - 86400; // 86400 seconds in a day
}


/**
 * Builds the Gmail search query based on the last processed time.
 * @param {number} lastProcessedTime The Unix timestamp of the last processed email.
 * @return {string} The Gmail search query string.
 */
function buildSearchQuery(lastProcessedTime) {
  if (lastProcessedTime === 0) { 
    console.log("No last processed time found. Fetching the most recent email with the label.");
    var lastEmailDate = getLastEmailDate();
    if (lastEmailDate) {
      return 'after:' + Math.floor(lastEmailDate.getTime() / 1000) + ' label:' + CONFIG.LABEL_NAME;
    } else {
      console.log("No emails found with the specified label. Exiting.");
      return null;
    }
  } else {
    return 'after:' + Math.floor(lastProcessedTime) + ' label:' + CONFIG.LABEL_NAME;
  }
}

/**
 * Gets the date of the most recent email with the specified label.
 * @return {Date} The date of the most recent email or null if none found.
 */
function getLastEmailDate() {
  var threads = GmailApp.search('label:' + CONFIG.LABEL_NAME, 0, 1);
  if (threads.length > 0) {
    var messages = threads[0].getMessages();
    if (messages.length > 0) {
      return messages[0].getDate();
    }
  }
  return null;
}

/**
 * Converts HTML content to Markdown format.
 * @param {string} html The HTML content to convert.
 * @return {string} The converted Markdown content.
 */
function convertHtmlToMarkdown(html) {
  // Remove script and style elements
  html = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  html = html.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');

  // Convert common HTML elements to Markdown
  var markdown = html
    .replace(/<h1[^>]*>(.*?)<\/h1>/gi, '# $1\n\n')
    .replace(/<h2[^>]*>(.*?)<\/h2>/gi, '## $1\n\n')
    .replace(/<h3[^>]*>(.*?)<\/h3>/gi, '### $1\n\n')
    .replace(/<h4[^>]*>(.*?)<\/h4>/gi, '#### $1\n\n')
    .replace(/<h5[^>]*>(.*?)<\/h5>/gi, '##### $1\n\n')
    .replace(/<h6[^>]*>(.*?)<\/h6>/gi, '###### $1\n\n')
    .replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<hr\s*\/?>/gi, '---\n\n')
    .replace(/<b>(.*?)<\/b>/gi, '**$1**')
    .replace(/<strong>(.*?)<\/strong>/gi, '**$1**')
    .replace(/<i>(.*?)<\/i>/gi, '*$1*')
    .replace(/<em>(.*?)<\/em>/gi, '*$1*')
    .replace(/<a\s+(?:[^>]*?\s+)?href="([^"]*)"[^>]*>(.*?)<\/a>/gi, '[$2]($1)')
    .replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, function(match, list) {
      return list.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, '- $1\n');
    })
    .replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, function(match, list) {
      var index = 1;
      return list.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, function(match, item) {
        return (index++) + '. ' + item + '\n';
      });
    });

  // Remove any remaining HTML tags
  markdown = markdown.replace(/<[^>]+>/g, '');

  // Decode HTML entities
  markdown = markdown.replace(/ /g, ' ')
                     .replace(/&/g, '&')
                     .replace(/</g, '<')
                     .replace(/>/g, '>')
                     .replace(/"/g, '"')
                     .replace(/'/g, "'");

  // Trim extra whitespace and newlines
  markdown = markdown.replace(/^\s+|\s+$/gm, '').replace(/\n{3,}/g, '\n\n');

  return markdown;
}

/**
 * Truncates text content to a maximum length.
 * @param {string} content The content to truncate.
 * @param {number} maxLength The maximum allowed length of the content.
 * @return {string} The truncated content.
 */
function truncateContent(content, maxLength = 49999) { 
  if (content && content.length > maxLength) {
    console.log(`Truncating content from ${content.length} to ${maxLength} characters`);
    return content.substring(0, maxLength);
  }
  return content;
}

/**
 * Shortens links in the content and creates a map for expansion.
 * @param {string} content The text content containing links.
 * @return {object} An object containing the shortened content and the link map.
 */
function shortenLinks(content) {
  let linkMap = {};
  let linkCounter = 1;
  let shortenedContent = content.replace(/https?:\/\/[^\s"'<>]+/g, function(match) {
    let cleanLink = match.replace(/[.,;:!?)"'\]]+$/, ''); 
    let shortLink = `${LINK_PREFIX}${linkCounter.toString().padStart(3, '0')}`;
    linkMap[shortLink] = cleanLink;
    linkCounter++;
    return shortLink;
  });
  return { shortenedContent, linkMap };
}

/**
 * Expands shortened links in news items using the provided link map.
 * @param {Array<object>} newsItems An array of news items, potentially with shortened links.
 * @param {object} linkMap A map of shortened links to their original URLs.
 * @return {Array<object>} The array of news items with expanded links.
 */
function expandShortLinks(newsItems, linkMap) {
  if (!Array.isArray(newsItems)) {
    console.error("newsItems is not an array:", newsItems);
    return []; 
  }
  return newsItems.map(item => {
    if (item && item.link && linkMap[item.link]) {
      item.link = linkMap[item.link];
    }
    return item; 
  });
}

/**
 * Contains functions for interacting with the Gemini API.
 */
const GeminiAPI = { 
  // API configuration (retries and delay)
  MAX_RETRIES: 3, 
  RETRY_DELAY: 2000,

  /**
   * Analyzes email content to extract news items using the Gemini API.
   * @param {string} content The email content to analyze.
   * @param {string} contentShort The shortened email content (for API limits).
   * @param {object} emailMetadata Metadata about the email.
   * @return {object} An object containing the status, news count, raw response, and extracted news items (if successful).
   */
  analyzeEmailContent: function(content, contentShort, emailMetadata) {
    console.log(`Processing email from ${emailMetadata.sender} sent on ${emailMetadata.email_date}`);

    // Construct the prompt for the Gemini API
    const prompt = `You are expert in fetching news from newsletters like Morning Brew, New York Times etc. Please extract ALL news or articles in JSON format from the newsletter below. Each news item should have the following fields: 'title', 'description', 'link', and 'category'. The category should be a single word or short phrase describing the main topic of the news item (e.g., 'Technology', 'Politics', 'Economy', 'Sports', etc.).
    
    If it's a longer article or section in newsletter then fetch the whole content of this particular passage. The output should be a JSON array of objects, each representing a single news item. When extracting the link, intelligently fetch the most relevant link associated with the news item - there ALWAYS must be a link, if you can't find, don't fetch this news. 
    
    Newsletter: ${contentShort}`; 

    // Retry logic for API requests
    for (let retries = 0; retries < this.MAX_RETRIES; retries++) {
      console.log(`Sending request to Gemini API (Attempt ${retries + 1})...`);
      try {
        const rawResponse = this.generateContent(prompt); 
        console.log("Received response from Gemini API");
        console.log("API Response Preview:", rawResponse.substring(0, 500) + "...");

        // Extract JSON array from the response
        const jsonMatch = rawResponse.match(/\[[\s\S]*\]/);
        if (jsonMatch) {
          const newsItems = JSON.parse(jsonMatch[0]);
          console.log(`Extracted ${newsItems.length} news items from the email`);
          return { status: 'success', newsCount: newsItems.length, rawResponse: rawResponse, newsItems: newsItems };
        } else {
          throw new Error("No JSON array found in the response");
        }
      } catch (error) {
        console.error(`Error on attempt ${retries + 1}:`, error);
        if (retries < this.MAX_RETRIES - 1) {
          console.log(`Retrying in ${this.RETRY_DELAY} milliseconds...`);
          Utilities.sleep(this.RETRY_DELAY);
        } else {
          return { status: 'failure', newsCount: 0, rawResponse: null }; 
        }
      }
    }

    return { status: 'failure', newsCount: 0, rawResponse: null }; 
  },

  /**
   * Validates extracted news items to ensure they have required fields.
   * @param {Array<object>} newsItems An array of news items.
   * @return {Array<object>} The array of validated news items.
   */
  validateNewsItems: function(newsItems) {
    return newsItems.filter(item => {
      if (!item.link) { 
        console.warn("Skipping news item without link:", item);
        return false;
      }
      // Ensure all fields exist, set to empty string if missing
      item.title = item.title || ""; 
      item.description = item.description || ""; 
      item.category = item.category || ""; 
      return true;
    });
  },

  /**
   * Sends a request to the Gemini API to generate content based on a prompt.
   * @param {string} prompt The prompt to send to the Gemini API.
   * @return {string} The generated text response from the Gemini API.
   */
  generateContent: function(prompt) {
    // **IMPORTANT!** Replace 'YOUR_GEMINI_API_KEY' with your actual Gemini API key
    const apiKey = 'YOUR_GEMINI_API_KEY'; 

    const apiEndpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent";
    const url = `${apiEndpoint}?key=${apiKey}`;

    const payload = {
      contents: [
        {
          parts: [
            {
              text: prompt
            }
          ]
        }
      ],
      generationConfig: { 
        temperature: 1,
        topP: 0.95,
        topK: 40,
        maxOutputTokens: 10192, // Adjust as needed
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true 
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      console.log(`API Response Code: ${responseCode}`);

      if (responseCode !== 200) { 
        throw new Error(`API request failed with status ${responseCode}: ${responseText}`);
      }

      const parsedResponse = JSON.parse(responseText);
      return parsedResponse.candidates[0].content.parts[0].text; 
    } catch (error) {
      console.error('Error in generateContent:', error);
      throw error; 
    }
  },

  /**
   * Adds extracted news items to the "News" sheet in the spreadsheet.
   * @param {Array<object>} newsItems An array of validated news items.
   * @param {object} emailMetadata Metadata about the email containing the news.
   * @param {string} content The original email content.
   * @param {string} contentShort The shortened email content.
   * @return {number} The number of news items successfully added to the sheet.
   */
  addNewsToSheet: function(newsItems, emailMetadata, content, contentShort) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("News");
    if (!sheet) {
      console.error("News sheet not found");
      return 0; 
    }

    // Add headers if the sheet is empty
    if (sheet.getLastRow() === 0) { 
      const headers = [
        "Title", 
        "Description", 
        "Link", 
        "Category", 
        "Email Date", 
        "Sender", 
        "Fetched Time", 
        "Email ID"
      ];
      sheet.appendRow(headers);
      console.log("Added headers to News sheet");
    }

    const newRows = newsItems.map(item => [
      this.truncateContent(item.title),
      this.truncateContent(item.description),
      item.link, 
      item.category,
      emailMetadata.email_date,
      emailMetadata.sender,
      emailMetadata.fetched_time,
      emailMetadata.emailId
    ]);

    // Add new rows to the sheet
    if (newRows.length > 0) {
      try {
        sheet.insertRowsAfter(1, newRows.length); 
        sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows); 
        console.log(`Added ${newRows.length} news items to the News sheet below the header`);
      } catch (error) {
        console.error("Error adding news to sheet:", error);
        return 0; 
      }
    }
    return newRows.length; 
  },

  // Truncate content for Google Sheet cell limit
  truncateContent: function(content) {
    return truncateContent(content); 
  }
}; 
