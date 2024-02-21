// Define global variables
var apiKey;

// Create custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Extension')
    .addItem('Configure API Key', 'configureApiKey')
    .addToUi();
}

// Function to configure API key
function configureApiKey() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter your OpenAI API Key:', ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    apiKey = result.getResponseText();
    PropertiesService.getDocumentProperties().setProperty('apiKey', apiKey);
    ui.alert('API Key set successfully!');
  }
}



// CAN USE FOR FREE
function CHATGPT(text, prompt, systemPrompt='', maxTokens=200, temperature=0.0, model='gpt-3.5-turbo') {
  console.log("Executing function: ChatGPT")
  return OpenAIGPT(text, prompt, systemPrompt=systemPrompt, maxTokens=maxTokens, temperature=temperature, model=model)
}
// NEED TO UPGRADE YOUR ACCOUNT BEFORE USE
function GPT4(text, prompt, systemPrompt='', maxTokens=200, temperature=0.0, model='gpt-4') {
  console.log("Executing function: GPT4")
  return OpenAIGPT(text, prompt, systemPrompt=systemPrompt, maxTokens=maxTokens, temperature=temperature, model=model)
}

// base function for OpenAI GPT models with chat format
function OpenAIGPT(text, prompt, systemPrompt='', maxTokens=200, temperature=0.0, model='gpt-3.5-turbo') {
  var apiKey = PropertiesService.getDocumentProperties().getProperty('apiKey');
  Logger.log(apiKey)

  
  // set default system prompt
  if (systemPrompt === '') {
    systemPrompt = "You are a helpful assistant."
  }
  Logger.log("System prompt: " + systemPrompt)
  // reset default values in case user puts empty string/cell as argument
  if (maxTokens === '') {
    maxTokens = 200
  }
  Logger.log("maxTokens: " + maxTokens)
    if (temperature === '') {
    temperature = 0.0
  }
  Logger.log("temperature: " + temperature)

  // configure the API request to OpenAI
  const data = {
    "model": model, // "gpt-3.5-turbo",
    "messages": [
      {"role": "system", "content": systemPrompt},
      {"role": "user", "content": text + "\n" + prompt}
    ],
    "temperature": temperature,
    "max_tokens": maxTokens,
    "n": 1,
    "top_p": 1.0,
    "frequency_penalty": 0.0,
    "presence_penalty": 0.0,
  };
  const options = {
    'method' : 'post',
    'contentType': 'application/json',
    'payload' : JSON.stringify(data),
    'headers': {
      Authorization: 'Bearer ' + apiKey,
    },
  };
  const response = UrlFetchApp.fetch(
    'https://api.openai.com/v1/chat/completions',
    options
  );
  Logger.log("Model API response: " + response.getContentText());

  // Send the API request 
  const result = JSON.parse(response.getContentText())['choices'][0]['message']['content'];
  Logger.log("Content message response: " + result)


  // use OpenAI moderation API to block problematic outputs.
  // see: https://platform.openai.com/docs/guides/moderation/overview
  var optionsModeration = {
    'payload' : JSON.stringify({"input": result}),
    'method' : 'post',
    'contentType': 'application/json',
    'headers': {
      Authorization: 'Bearer ' + apiKey,
    },
  }
  const responseModeration = UrlFetchApp.fetch(
    'https://api.openai.com/v1/moderations',
    optionsModeration
  );
  Logger.log("Moderation API response: " + responseModeration)
  // Send the API request  for moderation API
  const resultFlagged = JSON.parse(responseModeration.getContentText())['results'][0]['flagged'];
  Logger.log('Output text flagged? ' + resultFlagged)
  // do not return result, if moderation API determined that the result is problematic
  if (resultFlagged === true) {
    return "The OpenAI moderation API blocked the response."
  }

  // try parsing the output as JSON or markdown table. 
  // If JSON, then JSON values are returned across the cells to the right of the selected cell.
  // If markdown table, then the table is spread across the cells horizontally and vertically
  // If not JSON or markdown table, then the full result is directly returned in a single cell
  
  // JSON attempt
  if (result.includes('{') && result.includes('}')) {
    // remove boilerplate language before and after JSON/{} in case model outputs characters outside of {} delimiters
    // assumes that only one JSON object is returned and it only has one level and is not nested with multiple levels (i.e. not multiple {}{} or {{}})
    let cleanedJsonString = '{' + result.split("{")[1].trim()
    cleanedJsonString = cleanedJsonString.split("}")[0].trim() + '}'
    Logger.log(cleanedJsonString)
    // try parsing
    try {
      // parse JSON
      const resultJson = JSON.parse(cleanedJsonString);
      Logger.log(resultJson);
      // Initialize an empty array to hold the values
      const resultJsonValues = [];
      // Loop through the object using for...in loop
      for (const key in resultJson) {
        // Push the value of each key into the array
        resultJsonValues.push(resultJson[key]);
        Logger.log(key);
        Logger.log(resultJson[key])
      }
      console.log("Final JSON output: " + resultJsonValues)
      return [resultJsonValues]
    // Just return full response if result not parsable as JSON
    } catch (e) {
      Logger.log("JSON parsing did not work. Error: " + e);
    }
  }

  // .md table attempt
  // determine if table by counting "|" characters. 6 or more
  if (result.split("|").length - 1 >= 6) {
    let arrayNested = result.split('\n').map(row => row.split('|').map(cell => cell.trim()));
    // remove the first and last element of each row, they are always empty strings due to splitting
    // this also removes boilerplate language before and after the table
    arrayNested = arrayNested.map((array) => array.slice(1, -1));
    // remove nested subArray if it is entirely empty. This happens when boilerplate text is separated by a "\n" from the rest of the text
    arrayNested = arrayNested.filter((subArr) => subArr.length > 0);
    // remove the "---" header indicator
    arrayNested = arrayNested.filter((subArr) => true != (subArr[0].includes("---") && subArr[1].includes("---")))
    console.log("Final output as array: ", arrayNested)
    return arrayNested
  }

  // if neither JSON nor .md table
  console.log("Final output without JSON or .md table: ", result)
  return result; 

}




