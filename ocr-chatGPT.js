
// Google Cloud Vision APIを使ってOCRを実行する関数
function extractTextFromImage(data) {
  const CLOUD_VISION_API_KEY = PropertiesService.getScriptProperties().getProperty('Cloud_Vision');
  const visionApiUrl = 'https://vision.googleapis.com/v1/images:annotate?key=' + CLOUD_VISION_API_KEY;

  const base64Image = data.slice(data.indexOf(',') + 1);
  const payload = {
    requests: [{
      image: { content: base64Image },
      features: [{ type: 'TEXT_DETECTION' }]
    }]
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(visionApiUrl, options);
  const json = JSON.parse(response.getContentText());

  const annotation = json.responses?.[0]?.fullTextAnnotation?.text;

  if (!annotation) {
    throw new Error("OCRに失敗しました。文字が検出されませんでした。");
  }

  return annotation;
}


// ChatGPT APIを使ってテキストを解析する関数
function analyzePdfText(pdfText) {
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OpenAI_Key');
  const openaiUrl = 'https://api.openai.com/v1/chat/completions';
  //金額,着工,完了,発注元,発行日,住所,担当,物件名
  const payload = {
    model: "gpt-3.5-turbo",
    messages: [
      { role: "system", content: "You are a helpful assistant." },
      {
        role: "user",
        content: `以下の契約書のテキストから、次の項目を可能な限り抽出してください。
  - 発注金額（契約金額・注文金額・請負金額などでも可）
  - 着工日（工事開始日とも言う）
  - 完了日（工事完了日・引き渡し日などでも可）
  - 発注元（発行者・注文者・発注者 など）
  - 発行日（契約日・書類作成日などでも可）
  - 現場住所（工事場所・所在地・受渡場所などでも可）
  - 担当者（工事担当者・工事担当なども可）
  - 物件名（工事物件名・名称・工事名称など可）

  以下が契約書の全文です：
  ${pdfText}`
      }
    ],
    max_tokens: 300,
    temperature: 0
  };


  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(openaiUrl, options);
  Logger.log(response.getContentText());  // レスポンスの内容を確認
  const json = JSON.parse(response.getContentText());



  if(json && json.choices && json.choices.length>0){
    const content = json.choices[0].message.content;
    const result = json解析(content);
    Logger.log(result);
    return result;
  }
  // // ChatGPTの結果を返す
  // return json.choices[0].message.content.trim();
}

//受け取ったJsonの内容を分析する
function json解析(content){
  const lines = content.split('\n');

  let kin = "";
  let start = "";
  let finish = "";
  let client = "";//顧客は選択させるのでここでは必要なし
  let day = "";
  let add = "";
  let tan = "";
  let gen = "";

  lines.forEach(line => {
    const parts = line.split(':');
    if(parts.length > 1){
      const key = parts[0].trim();
      const value = parts[1].trim();
      if(key.includes('発注金額')){
        kin = value;
      }else if(key.includes('着工日')){
        start = value;
      }else if(key.includes('完了日')){
        finish = value;
      }else if(key.includes('発行元')){
        client = value;
      }else if(key.includes('発行日')){
        day = value;
      }else if(key.includes('現場住所')){
        add = value;
      }else if(key.includes('担当者')){
        tan = value;
      }else if(key.includes('物件名')){
        gen = value;
      }
    }
  });
  rn = [kin,start,finish,client,day,add,tan,gen];
  Logger.log(rn);
  return rn;
}
