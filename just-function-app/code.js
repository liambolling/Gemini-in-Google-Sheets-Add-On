/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e); // Ensures that the menu is created when the add-on is installed.
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createAddonMenu();
  menu.addItem('🔑 Add your API key', 'showApiKeyModal')
  menu.addToUi(); // Apply the configured menu to the UI
}

/**
 * Send this text to Gemini for a response.
 * @param {string} inputText - The text for the API to process.
 * @customfunction
 */
function gemini(inputText) {
  const storedModelName = PropertiesService.getUserProperties().getProperty('geminiModelName');
  if (!storedModelName) {
    PropertiesService.getUserProperties().setProperty('geminiModelName', 'gemini-pro-1.0');
  }

  try {
    // Retrieves the API key and throws an error if it is not set.
    const apiKey = PropertiesService.getUserProperties().getProperty('geminiApiKey');
    if (!apiKey) throw new Error('API key not set.');

    // Constructs the request URL, setting a default if no specific URL is stored.
    const baseModelUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent`;
    const url = `${baseModelUrl}?key=${apiKey}`;

    // Sets the request options, including payload and HTTP headers.
    const options = {
      method: "post",
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{role: "user", parts: [{text: inputText}]}],
        safetySettings: [
          { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
          { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
          { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
          { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
        ]
      }),
      muteHttpExceptions: true
    };

    // Sends the API request and handles the response.
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`API request failed with status ${response.getResponseCode()}`);
    }

    // Parses the response and returns the generated content.
    const responseData = JSON.parse(response.getContentText());
    return responseData.candidates[0].content.parts[0].text;
  } catch (error) {
    throw new Error("Failed to reach Gemini API: " + error.message);
  }
}

/**
 * Displays a modal for API key entry.
 */
function showApiKeyModal() {
  const ui = HtmlService.createHtmlOutputFromFile('key_modal')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Add your API key');
}

/**
 * Sets the API key in the user's properties.
 * @param {string} apiKey - The API key to set.
 */
function setApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('geminiApiKey', apiKey);
  SpreadsheetApp.getUi().alert('API key was saved.');
}

function removeApiKey() {
  // Deletes the stored API key for the Gemini API from the user's property store.
  PropertiesService.getUserProperties().deleteProperty('geminiApiKey');
  // Alerts the user that the API key has been successfully removed.
  SpreadsheetApp.getUi().alert('API key was removed.');
}