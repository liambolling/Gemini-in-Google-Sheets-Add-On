/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e); // Ensures that the menu is created when the add-on is installed.
}

function onOpen(e) {

  const ui = SpreadsheetApp.getUi();

  if (e.authMode !== ScriptApp.AuthMode.NONE) {
    const userProperties = PropertiesService.getUserProperties();
    const hasSeenSidebar = userProperties.getProperty('hasSeenSidebar');

    // Conditionally showing the sidebar based on whether they've seen it before
    // and based on authMode allowing for user properties access
    if (!hasSeenSidebar) {
      showSidebar();
      userProperties.setProperty('hasSeenSidebar', 'true');
    }

    ui.createMenu("Gemini™ Functions") // Changed from createMenu to createAddonMenu
      .addItem('📓 Get started', 'showSidebar')
      .addItem('🔑 Add your API key', 'showApiKeyModal')
      .addItem('👯‍♀️ Select a model', 'showModelSelectionModal')
      .addItem('🚀 Run Gemini™ on selection', 'showDelayPrompt')
      .addItem('🧪 Who made this add-in', 'attributeGoogleLabs')
      .addToUi();
  } else {
    ui.createAddonMenu() // Changed from createMenu to createAddonMenu
      .addItem('📓 Get started', 'showSidebar')
      .addItem('🔑 Add your API key', 'showApiKeyModal')
      .addItem('👯‍♀️ Select a model', 'showModelSelectionModal')
      .addItem('🚀 Run Gemini™ on selection', 'showDelayPrompt')
      .addItem('🧪 Who made this add-in', 'attributeGoogleLabs')
      .addToUi();
  }
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Getting started')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * One off call to Gemini without batching.
 * @param {string} inputText - The text for the API to process.
 * @customfunction
 */
function gemini(inputText) {
  return callGeminiAPI(inputText);
}

function showApiKeyModal() {
  const storedModelName = PropertiesService.getUserProperties().getProperty('geminiModelName');
  if (!storedModelName) {
    selectModel('gemini-pro-1.0');
  }

  var ui = HtmlService.createHtmlOutputFromFile('apiKeyModal')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Add you API key');
}

function openGetStartedDoc() {
  const url = "https://docs.google.com/document/d/171rGYBQcVDCJrKJrB4Wbl1rAyfaw68cdPlX7voodG1k";
  const html = HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank');google.script.host.close();</script>`);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening get started doc...');
}

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

function setApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('geminiApiKey', apiKey);
  SpreadsheetApp.getUi().alert('API key was saved.');
}

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
  PropertiesService.getUserProperties().deleteProperty('geminiApiKey');
  SpreadsheetApp.getUi().alert('API key was removed.');
}


function attributeGoogleLabs() {
  SpreadsheetApp.getUi().alert('Created by Liam on the Google Labs team. All mentions of Gemini, Google Sheets or Google are the rights and copyright of Google Inc.');
}

function showApiKeyPrompt() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Gemini API Key', 'Enter your Gemini API key:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() === ui.Button.OK) {
    const apiKey = prompt.getResponseText();
    PropertiesService.getUserProperties().setProperty('geminiApiKey', apiKey);
    ui.alert('API key saved successfully!');
  }
}



function callGeminiAPI(inputText) {

  const storedModelName = PropertiesService.getUserProperties().getProperty('geminiModelName');
  if (!storedModelName) {
    selectModel('gemini-pro-1.0');
  }

  try {
    const apiKey = PropertiesService.getUserProperties().getProperty('geminiApiKey');
    if (!apiKey) throw new Error('API key not set.');

    const modelName = PropertiesService.getUserProperties().getProperty('geminiModelName') || 'gemini-pro-1.0';
    const url = PropertiesService.getUserProperties().getProperty('geminiModelUrl') + `?key=${apiKey}` ||
      `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

    const options = {
      method: "post",
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{
          role: "user",
          parts: [{
            "text": "Whatever I say next, always return something. If I violate any guidelines, just tell me instead of causing an error."
          }, {
            "text": inputText
          }]
        }],
        safetySettings: [
          { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_ONLY_HIGH" },
          { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_ONLY_HIGH" },
          { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_ONLY_HIGH" },
          { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_ONLY_HIGH" }
        ]
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(`API request failed with status ${response.getResponseCode()}`);
    }

    const responseData = JSON.parse(response.getContentText());

    return responseData.candidates[0].content.parts[0].text;
  } catch (error) {
    throw new Error("Failed to reach Gemini API: " + error.message);
  }
}


function showDelayPrompt() {
  var ui = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(550)
    .setHeight(525);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Run Gemini on Selection');
}

function runBatchFunctions(delay) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getValues();

  // Mark all 'pending' cells first
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cell = range.getCell(i + 1, j + 1);
      var background = cell.getBackground();

      // Set note or change background color to indicate processing
      cell.setComment('To be processed soon...');
      cell.setBackground('#ffcc00');  // Temporary color
    }
  }

  // Then gradually process each cell
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cell = range.getCell(i + 1, j + 1);
      var background = cell.getBackground();

      cell.clearNote();
      cell.setComment('Processing now...');

      try {
        var result = callGeminiAPI(values[i][j]);
        cell.setValue(result);
        cell.setBackground(null);
      } catch (e) {
        cell.setValue(e.message);
        cell.setBackground('#ff0000');
      }

      cell.clearNote();
      Utilities.sleep(delay * 1000);

    }
  }
}





function gemini_api(inputText) {
  const batchSize = 1;
  const waitTime = 1000; // Adjust wait time as needed
  const queue = CacheService.getScriptCache();
  const cacheKey = "geminiQueue";

  if (!inputText) return "Please provide input text.";

  // Add input to the queue
  let currentQueue = queue.get(cacheKey) || JSON.stringify([]);
  currentQueue = JSON.parse(currentQueue);
  currentQueue.push(inputText);
  queue.put(cacheKey, JSON.stringify(currentQueue), 21600); // Cache queue for 6 hours

  // Trigger batch processing if batch size reached
  if (currentQueue.length % batchSize === 0) {
    const batch = currentQueue.splice(0, batchSize);
    // Return a placeholder instead of directly calling processBatch
    return "Processing...";
  }

  return "Loading...";
}


// Make sure to mark the processBatch function as async
async function processBatch(batch, waitTime, resultRange) {

  // Use a for loop instead of forEach to ensure serial execution
  for (let i = 0; i < batch.length; i++) {
    const inputText = batch[i];
    try {
      Utilities.sleep(waitTime);

      const apiKey = PropertiesService.getUserProperties().getProperty('geminiApiKey');
      const modelName = PropertiesService.getUserProperties().getProperty('geminiModelName') || 'gemini-ultra';
      const url = PropertiesService.getUserProperties().getProperty('geminiModelUrl') || `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

      const options = {
        method: "post",
        contentType: 'application/json',
        muteHttpExceptions: true
      };

      const data = {
        contents: [{
          role: "user",
          parts: [{
            "text": "Whatever I say next, always return something. If I violate any guidelines, just tell me instead of causing an error."
          }, {
            "text": inputText
          }]
        }],
        generationConfig: {
          temperature: 0.9,
          topK: 1,
          topP: 1,
          maxOutputTokens: 2048,
          stopSequences: []
        },
        safetySettings: [
          { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_ONLY_HIGH" },
          { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_ONLY_HIGH" },
          { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_ONLY_HIGH" },
          { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_ONLY_HIGH" }
        ]
      };

      options.payload = JSON.stringify(data);

      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() === 200) {
        const json = JSON.parse(response.getContentText());
        if (json.candidates && json.candidates[0].content.parts[0].text) {
          // Update the resultRange with the API results
          resultRange.getCell(i + 1, 1).setValue(json.candidates[0].content.parts[0].text);
        } else {
          Logger.log(JSON.stringify(data));
          resultRange.getCell(i + 1, 1).setValue('Error: This prompt might go against the Gemini terms of use.');
        }
      } else if (response.getResponseCode() === 429) {
        resultRange.getCell(i + 1, 1).setValue('Error: Too many requests to the Gemini model. Please wait then try again.');
      } else {
        const errorJson = JSON.parse(response.getContentText());
        resultRange.getCell(i + 1, 1).setValue(`Error: ${errorJson.message}`);
      }
    } catch (error) {
      Logger.log(error);
      resultRange.getCell(i + 1, 1).setValue(`Error: ${error.message}`);
    }
  }
}
