/**
 * Google Apps Script to translate words in a spreadsheet using OpenAI API
 * With support for optional context column
 */

// Your OpenAI API key
const OPENAI_API_KEY = "";

/**
 * Creates a menu item in Google Sheets to run the translation
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Translation Tools")
    .addItem("Translate Words with OpenAI", "translateWordsInColumn")
    .addToUi();
}

/**
 * Main function to translate words in the selected column
 */
function translateWordsInColumn() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get languages from A1 and B1
  const sourceLanguage = sheet.getRange("A1").getValue();
  const targetLanguage = sheet.getRange("B1").getValue();

  // Validate that languages are specified
  if (!sourceLanguage || !targetLanguage) {
    ui.alert("Please specify source language in A1 and target language in B1");
    return;
  }

  // Ask if context column is present
  const contextResponse = ui.alert(
    "Context Column",
    "Is there a context column next to your words? (Column to the right of selected column)",
    ui.ButtonSet.YES_NO
  );

  const hasContextColumn = contextResponse === ui.Button.YES;

  // Get active spreadsheet and selected range
  const selectedRange = sheet.getActiveRange();

  if (selectedRange.getNumColumns() !== 1) {
    ui.alert("Please select a single column containing words to translate.");
    return;
  }

  const numRows = selectedRange.getNumRows();
  const startRow = selectedRange.getRow();
  const wordColumn = selectedRange.getColumn();
  const contextColumn = wordColumn + 1;
  const translationColumn = hasContextColumn ? wordColumn + 2 : wordColumn + 1;

  // Get all words at once
  const words = selectedRange.getValues().flat();

  // If using context, get all context values
  let contexts = [];
  if (hasContextColumn) {
    contexts = sheet
      .getRange(startRow, contextColumn, numRows, 1)
      .getValues()
      .flat();
  }

  // Create a progress indicator
  let translatedCount = 0;
  const totalToTranslate = words.filter((word) => word.trim() !== "").length;

  // Process translations sequentially with promises
  const processTranslations = async () => {
    for (let i = 0; i < numRows; i++) {
      const word = words[i];
      if (word.trim() === "") continue;

      const context = hasContextColumn ? contexts[i] : "";

      try {
        const translation = await translateWithOpenAI(
          word,
          sourceLanguage,
          targetLanguage,
          context
        );

        // Write each translation as it completes
        sheet.getRange(startRow + i, translationColumn).setValue(translation);
        translatedCount++;

        // To avoid hitting OpenAI rate limits
        Utilities.sleep(200);
      } catch (error) {
        Logger.log(`Error translating word "${word}": ${error.toString()}`);
        sheet
          .getRange(startRow + i, translationColumn)
          .setValue("ERROR: " + error.toString());
      }
    }

    // Show completion message only after all translations are done
    ui.alert(
      `Translation complete! Translated ${translatedCount} of ${totalToTranslate} words.`
    );
  };

  // Start the translation process
  processTranslations();
}

/**
 * Translates a word using OpenAI API, with optional context
 *
 * @param {string} word - The word to translate
 * @param {string} sourceLanguage - The source language
 * @param {string} targetLanguage - The target language
 * @param {string} context - Optional context for translation (can be empty)
 * @return {Promise<string>} The translated word
 */
async function translateWithOpenAI(
  word,
  sourceLanguage,
  targetLanguage,
  context = ""
) {
  const url = "https://api.openai.com/v1/chat/completions";

  // Adjust the prompt based on whether context is provided
  let systemPrompt = `You are a professional translator. Translate the given word from ${sourceLanguage} to ${targetLanguage}.`;

  if (context && context.trim() !== "") {
    systemPrompt += ` Important: Consider this specific context: "${context}". The translation should be appropriate for this context.`;
  }

  systemPrompt += ` Respond with just the translated word, nothing else.`;

  const payload = {
    model: "gpt-4-turbo",
    messages: [
      {
        role: "system",
        content: systemPrompt,
      },
      {
        role: "user",
        content: word,
      },
    ],
    temperature: 0.5,
    max_tokens: 50,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = await UrlFetchApp.fetch(url, options);
  const responseData = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error(`API Error: ${responseData.error.message}`);
  }

  return responseData.choices[0].message.content.trim();
}
