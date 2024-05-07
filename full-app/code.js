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

  // Configure menu items for all users with authentication mode other than NONE
  if (e.authMode !== ScriptApp.AuthMode.NONE) {
    const userProperties = PropertiesService.getUserProperties();
    const hasSeenSidebar = userProperties.getProperty('hasSeenSidebar');

    // Show the sidebar for first-time users and mark it as seen
    if (!hasSeenSidebar) {
      showSidebar();
      userProperties.setProperty('hasSeenSidebar', 'true');
    }

    // Add items to the Gemini™ Functions menu
    menu.addItem('📓 Get started', 'showSidebar')
        .addItem('🔑 Add your API key', 'showApiKeyModal')
        .addItem('👯‍♀️ Select a model', 'showModelSelectionModal')
        .addItem('🚀 Run Gemini™ on selection', 'showDelayPrompt')
        .addItem('🧪 Who made this add-in', 'attributeGoogleLabs');
  } else {
    // Add items to the menu without user-specific options
    menu.addItem('📓 Get started', 'showSidebar')
        .addItem('🔑 Add your API key', 'showApiKeyModal')
        .addItem('👯‍♀️ Select a model', 'showModelSelectionModal')
        .addItem('🚀 Run Gemini™ on selection', 'showDelayPrompt')
        .addItem('🧪 Who made this add-in', 'attributeGoogleLabs');
  }

  menu.addToUi(); // Apply the configured menu to the UI
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Getting started')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Send this text to Gemini for a response.
 * @param {string} inputText - The text for the API to process.
 * @customfunction
 */
function gemini(inputText) {
  return callGeminiAPI(inputText);
}

/**
 * Displays a modal for API key entry.
 */
function showApiKeyModal() {
  const storedModelName = PropertiesService.getUserProperties().getProperty('geminiModelName');
  if (!storedModelName) {
    selectModel('gemini-pro-1.0');
  }
  const ui = HtmlService.createHtmlOutputFromFile('apiKeyModal')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Add your API key');
}

/**
 * Opens a Google Docs guide in a new tab.
 */
function openGetStartedDoc() {
  const url = "https://docs.google.com/document/d/171rGYBQcVDCJrKJrB4Wbl1rAyfaw68cdPlX7voodG1k/edit";
  const html = HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank');google.script.host.close();</script>`);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening get started doc...');
}

/**
 * Displays a modal to select a model for the API.
 */
function showModelSelectionModal() {
  const storedModelName = PropertiesService.getUserProperties().getProperty('geminiModelName');
  if (!storedModelName) {
    selectModel('gemini-pro-1.0');
  }
  const html = HtmlService.createHtmlOutputFromFile('modelSelectionModal')
    .setWidth(500)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Gemini model');
}

/**
 * Sets the API key in the user's properties.
 * @param {string} apiKey - The API key to set.
 */
function setApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('geminiApiKey', apiKey);
  SpreadsheetApp.getUi().alert('API key was saved.');
}

/**
 * Selects the API model and sets relevant properties.
 * @param {string} modelName - The model name to select.
 */
function selectModel(modelName) {
  const modelData = {
    'gemini-pro-1.0': { url: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent', description: 'The ok model' },
    'gemini-pro-1.5': { url: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent', description: 'The mostly best model' },
    'gemini-ultra': { url: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-ultra:generateContent', description: 'The best but slow model' },
  }[modelName];

  if (modelData) {
    PropertiesService.getUserProperties().setProperty('geminiModelName', modelName);
    PropertiesService.getUserProperties().setProperty('geminiModelUrl', modelData.url);
    google.script.host.close();
    SpreadsheetApp.getUi().alert(`Model selected: ${modelName}`);
  } else {
    SpreadsheetApp.getUi().alert('Invalid model selection.');
  }
}


function removeApiKey() {
  // Deletes the stored API key for the Gemini API from the user's property store.
  PropertiesService.getUserProperties().deleteProperty('geminiApiKey');
  // Alerts the user that the API key has been successfully removed.
  SpreadsheetApp.getUi().alert('API key was removed.');
}

function attributeGoogleLabs() {
  // Displays an alert with credits and copyright information.
  SpreadsheetApp.getUi().alert('Created by Liam on the Google Labs team. All mentions of Gemini, Google Sheets or Google are the rights and copyright of Google Inc.');
}

function showApiKeyPrompt() {
  // Creates a user interface prompt to input the Gemini API key.
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Gemini API Key', 'Enter your Gemini API key:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() === ui.Button.OK) {
    // Saves the entered API key into the user's property store if the OK button is clicked.
    const apiKey = prompt.getResponseText();
    PropertiesService.getUserProperties().setProperty('geminiApiKey', apiKey);
    // Alerts the user that the API key has been saved successfully.
    ui.alert('API key saved successfully!');
  }
}

function callGeminiAPI(inputText) {
  // Checks and sets a default model if not already set.
  const storedModelName = PropertiesService.getUserProperties().getProperty('geminiModelName');
  if (!storedModelName) {
    PropertiesService.getUserProperties().setProperty('geminiModelName', 'gemini-pro-1.0');
  }

  try {
    // Retrieves the API key and throws an error if it is not set.
    const apiKey = PropertiesService.getUserProperties().getProperty('geminiApiKey');
    if (!apiKey) throw new Error('API key not set.');

    // Constructs the request URL, setting a default if no specific URL is stored.
    const modelName = PropertiesService.getUserProperties().getProperty('geminiModelName') || 'gemini-pro-1.0';
    const baseModelUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;
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
 * Displays a modal dialog box using an HTML file from the script's resources.
 * The dialog box is used to prompt the user before running the Gemini API on selected cells.
 */
function showDelayPrompt() {
  const ui = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(550)
    .setHeight(525);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Run Gemini on Selection');
}

/**
 * Executes batch processing on a selection of cells in the active spreadsheet,
 * applying a specified delay between processing each cell.
 * 
 * @param {number} delay - The delay in seconds between processing each cell.
 */
function runBatchFunctions(delay) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();

  // First, mark all cells as 'pending' for processing
  values.forEach((row, i) => {
    row.forEach((_, j) => {
      const cell = range.getCell(i + 1, j + 1);
      cell.setComment('To be processed soon...');
      cell.setBackground('#ffcc00');  // Temporary color indicating processing
    });
  });

  // Process each cell with a delay
  values.forEach((row, i) => {
    row.forEach((value, j) => {
      const cell = range.getCell(i + 1, j + 1);
      cell.clearNote();
      cell.setComment('Processing now...');

      try {
        const result = callGeminiAPI(value);
        cell.setValue(result);
        cell.setBackground(null);  // Reset background after processing
      } catch (e) {
        cell.setValue(e.message);
        cell.setBackground('#ff0000');  // Red background on error
      }

      cell.clearNote();
      Utilities.sleep(delay * 1000);  // Delay in milliseconds
    });
  });
}





/**
 * Adds input text to a processing queue and returns the state of the queue.
 * If the queue reaches a predefined batch size, it triggers a batch processing.
 *
 * @param {string} inputText - The text to be added to the processing queue.
 * @returns {string} - A message indicating the processing state: either waiting for more inputs ("Loading...")
 *                     or starting the batch processing ("Processing..."). If no input text is provided, 
 *                     it prompts the user to provide text.
 */
function gemini_api(inputText) {
  const batchSize = 1; // Number of entries in the queue required to trigger processing
  const queue = CacheService.getScriptCache(); // Access to the script's cache service
  const cacheKey = "geminiQueue"; // Key under which the queue is stored in the cache

  if (!inputText) return "Please provide input text."; // Check for empty input and prompt for text

  // Retrieve the current queue from cache, or initialize it if not present
  let currentQueue = queue.get(cacheKey) || JSON.stringify([]);
  currentQueue = JSON.parse(currentQueue);
  currentQueue.push(inputText); // Add new input text to the queue
  queue.put(cacheKey, JSON.stringify(currentQueue), 21600); // Update the queue in the cache with a 6-hour expiration

  // Check if the queue length is equal to the batch size to trigger processing
  if (currentQueue.length % batchSize === 0) {
    const batch = currentQueue.splice(0, batchSize); // Remove processed items from the queue
    queue.put(cacheKey, JSON.stringify(currentQueue), 21600); // Update the cache after removing processed items
    return "Processing..."; // Return processing message when batch size is met
  }

  return "Loading..."; // Return loading message if batch size is not yet met
}



/**
 * Processes a batch of input texts using the Gemini model API, updating a given result range in a spreadsheet.
 * Each input text is sent sequentially to the API with a specified delay between requests.
 * 
 * @param {Array} batch - An array of strings, where each string is a batch item to be processed.
 * @param {number} waitTime - Time in milliseconds to wait before processing the next item in the batch.
 * @param {Range} resultRange - The spreadsheet range where results will be written.
 */
async function processBatch(batch, waitTime, resultRange) {
  // Retrieve configuration only once to avoid redundant API calls
  const apiKey = PropertiesService.getUserProperties().getProperty('geminiApiKey');
  const modelName = PropertiesService.getUserProperties().getProperty('geminiModelName') || 'gemini-ultra';
  const baseUrl = PropertiesService.getUserProperties().getProperty('geminiModelUrl') || 'https://generativelanguage.googleapis.com/v1beta/models';
  const url = `${baseUrl}/${modelName}:generateContent?key=${apiKey}`;
  const headers = {
    "Content-Type": "application/json",
    "Accept": "application/json"
  };
  const safetySettings = [
    { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
    { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
  ];

  // Process each item in the batch
  for (let i = 0; i < batch.length; i++) {
    Utilities.sleep(waitTime); // Introduce delay as specified

    const postData = JSON.stringify({
      contents: [{
        role: "user",
        parts: [
          { "text": "Whatever I say next, always return something. If I violate any guidelines, just tell me instead of causing an error." },
          { "text": batch[i] }
        ]
      }],
      generationConfig: {
        temperature: 0.9,
        topK: 1,
        topP: 1,
        maxOutputTokens: 2048,
        stopSequences: []
      },
      safetySettings: safetySettings
    });

    try {
      const response = UrlFetchApp.fetch(url, { method: "POST", headers: headers, payload: postData, muteHttpExceptions: true });
      handleResponse(response, i, resultRange);
    } catch (error) {
      Logger.log(error);
      resultRange.getCell(i + 1, 1).setValue(`Error: ${error.message}`);
    }
  }
}

/**
 * Handles the API response by parsing it and updating the spreadsheet with results or errors.
 * 
 * @param {HTTPResponse} response - The response object from the URL fetch.
 * @param {number} index - The index of the current batch item.
 * @param {Range} resultRange - The spreadsheet range to update with results.
 */
function handleResponse(response, index, resultRange) {
  const statusCode = response.getResponseCode();
  const content = response.getContentText();

  if (statusCode === 200) {
    const json = JSON.parse(content);
    const resultText = json.candidates && json.candidates[0].content.parts[0].text
                       ? json.candidates[0].content.parts[0].text
                       : 'Error: This prompt might go against the Gemini terms of use.';
    resultRange.getCell(index + 1, 1).setValue(resultText);
  } else if (statusCode === 429) {
    resultRange.getCell(index + 1, 1).setValue('Error: Too many requests to the Gemini model. Please wait then try again.');
  } else {
    const errorJson = JSON.parse(content);
    resultRange.getCell(index + 1, 1).setValue(`Error: ${errorJson.message}`);
  }
}
