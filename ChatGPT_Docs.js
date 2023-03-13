// Constants
const API_KEY = "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";
const MODEL_TYPE = "gpt-3.5-turbo"; //chatGPT model

// Creates a custom menu in Google Docs
function onOpen() {
  DocumentApp.getUi().createMenu("ChatGPT")
      .addItem("寫文章", "generatePrompt")
      .addItem("產生貼文", "generateIdeas")
      .addItem("翻譯成英文", "TranslateToEn")
      .addItem("翻譯成中文", "TranslateToCh")
      // .addItem("回覆Email", "ReplyEmail")
      // .addItem("翻譯英文", "generateIdeas")
      .addItem("詢問機器人", "menuItem")      
      .addToUi();
}

// Generates prompt based on the selected text and adds it to the document
function generatePrompt() {
  const doc = DocumentApp.getActiveDocument();
  const selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText();
  const body = doc.getBody();
  const prompt = "Generate an essay on " + selectedText;
  const temperature = 0.8;
  const maxTokens = 2060;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{role: "user", content: prompt}],
    temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };


  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  Logger.log(generatedText);
  body.appendParagraph(generatedText.toString());
}


// Generates prompt based on the selected text and adds it to the document
function generateIdeas() {
  const doc = DocumentApp.getActiveDocument();
  const selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText();
  const body = doc.getBody();
  const prompt = "Help me write 5 social media post on " + selectedText;
  const temperature = 0.8;
  const maxTokens = 2060;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{role: "user", content: prompt}],
    temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };


  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  Logger.log(generatedText);
  body.appendParagraph(generatedText.toString());
}


// Translate to ENglish
function TranslateToEn() {
  const doc = DocumentApp.getActiveDocument();
  const selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText(); 
  const body = doc.getBody();
  const prompt = "Help me translate" + selectedText + "in English";
  const temperature = 0.8;
  const maxTokens = 2060;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{role: "user", content: prompt}],
    temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };


  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  Logger.log(generatedText);
  body.appendParagraph(generatedText.toString());
}


// Translate to Chinese
function TranslateToCh() {
  const doc = DocumentApp.getActiveDocument();
  const selectedText = doc.getSelection().getRangeElements()[0].getElement().asText().getText();
  const body = doc.getBody();
  const prompt = "Help me translate" + selectedText + "in 台灣繁體中文";
  const temperature = 0.8;
  const maxTokens = 2060;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{role: "user", content: prompt}],
    temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };


  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  Logger.log(generatedText);
  body.appendParagraph(generatedText.toString());
}




// Generates input prompt 
function menuItem(){
  const doc = DocumentApp.getActiveDocument();
  const ui = DocumentApp.getUi();
  const userInput = ui.prompt (
    
    "ChatGPT 3.5 TurBo",
    "請輸入文字",
    ui.ButtonSet.OK_CANCEL 

  );
  const button = userInput.getSelectedButton(); 
  const prompt = userInput.getResponseText();
  if (!prompt & button == ui.Button.OK) {
    return DocumentApp.getUi().alert('不能輸入空白');
  }
  else if (button == ui.Button.CANCEL) {
    return;
  } 
  else if (button == ui.Button.CLOSE) {
    return;
  }
  const body = doc.getBody();
  const temperature = 0.8;
  const maxTokens = 2060;

  const requestBody = {
    model: MODEL_TYPE,
    messages: [{role: "user", content: prompt}],
    temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const generatedText = json['choices'][0]['message']['content'];
  Logger.log(generatedText);
  body.appendParagraph(generatedText.toString());
}
