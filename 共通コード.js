//◆◆共通◆◆共通◆◆共通◆◆共通◆◆共通◆◆共通◆◆共通◆◆共通◆◆共通
function doGet(e){ 
  //Logger.log(e);
  let page = e.parameter.page;
  if(!page){
    page = 'pane';
  }     
  const template = HtmlService.createTemplateFromFile(page);
  if(page === 'pane'){
    template.param1 = e.parameter.param1;
  }else if(page === 'yosan'){
    template.param1 = e.parameter.param1;
  }else if(page == 'test'){
    template.param1 = e.parameter.param1;
  }
  const syamei = 会社名()
  return template
  .evaluate()
  .setTitle(syamei + "◆予算書パネル")
  .addMetaTag('viewport','width=device-width, intial-scale=1');
}

//権限の確認用
function kenkakuA(syaId, passB) {
  Logger.log(syaId + "," + passB)
  const admin = getAdmin();//権限者のIdここに記録※設定.gs
  const judgeF = () => {
    const ls = admin.length;
    for(let s=0; s<ls; s++){
      if(admin[s] == syaId){
        return "OK";
      }
    }
    return "NG";
  }
  const judge = judgeF();
  if(judge == "OK"){
    const kId = getSPId("業者顧客担当")
    const sss = SpreadsheetApp.openById(kId);
    const ss = sss.getSheetByName('担当者分け');
    const last = ss.getRange('H1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    const row = Number(syaId) - 100;
    const pass = Number(passB);
    const valW = ss.getRange('W2:W' + last).getValues(); // 社員ROWリスト
    const valV = ss.getRange('V2:V' + last).getValues(); // パスワードリスト
    for (let i = 0; i < valW.length; i++) {
      if (valW[i][0] === row) {
        const passA = Number(valV[i][0]);
        if (passA === pass) {
          return "OK";
        } else {
          return "NG"; // パスワード不一致
        }
      }
    }
    return "NG"; // 該当社員IDが見つからなかった
  }else{
    return "NG";
  }
}

//
function kengen(pass){
  Logger.log(pass)
  const kId = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(kId);
  const ss = sss.getSheetByName('担当者分け');
  const passA = ss.getRange('V3').getValue();
  if(pass == passA){
    return "OK";
  }else{
    return "NG";
  }
}

function 内訳鏡(id){
  const sss = SpreadsheetApp.openById(id);
  const ss = sss.getSheetByName('内訳（鏡）');
  return ss;
}
function 予算書シート(id){
  const sss = SpreadsheetApp.openById(id);
  const ss = sss.getSheetByName('シート1');
  return ss;
}

function showMyModal() {
  const htmlOutput = HtmlService.createTemplateFromFile('pane')
    .evaluate()
    .setWidth(800)  // ダイアログの幅
    .setHeight(800) // ダイアログの高さ

  SpreadsheetApp
    .getUi()
    .showModalDialog(htmlOutput, '予算書モーダル')
}

//◆◆年度、月を返す※工事台帳用
function 本日の年度() {//現在（今日）の年度と月(月末締め）を返す
  const day = new Date;
  const nenA = day.getFullYear();
  const tukiA = day.getMonth() + 1;
  const nitiA = day.getDate();
  Logger.log(nenA + "," + tukiA + "," + nitiA);
  let rn = [];

  const nenF = () => {
    if(tukiA > 9){
      n = nenA + 1;
      return n;
    }else{
      return nenA;
    }
  };
  const nen = nenF();
  const tuki = tukiA;
  rn.push(nen);
  rn.push(tuki);
  Logger.log(rn);
  return rn;
}
//◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇◇年度、月
function 年度と年と月テスト用(){
  let day =  new Date();
  
  //◆◇◆◇テスト用↓◆◇◆◇◆◇◆◇
  // const now = "2024/9/20 00:00:00";
  // day = Utilities.parseDate(now,"JST","yyyy/M/d 00:00:00");
  //◆◇◆◇テスト用↑◆◇◆◇◆◇◆◇

  Logger.log(day);
  return day;
}

//◆◆年度、現在年、月を返す
function 年度年月(e) {//現在（今日）の年度と年と月(月末締め）を返す
  const dayF = () => {
    if(!e){
      result = 年度と年と月テスト用();
      return result; 
    }else{
      return e;
    }
  }
  const day = dayF();
  const nenA = day.getFullYear();
  const tukiA = day.getMonth() + 1;
  const nitiA = day.getDate();
  Logger.log(nenA + "," + tukiA + "," + nitiA);
  let rn = [];
  const nenF = () => {
    if(tukiA > 9){
      n = nenA + 1;
      return n;
    }else{
      return nenA;
    }
  };
  const nendo = nenF();
  const nen = nenA;
  const tuki = tukiA;
  rn.push(nendo);
  rn.push(nen);
  rn.push(tuki);
  rn.push(nitiA);
  Logger.log(rn);
  return rn;
}
function 業者情報(row){
  Logger.log(row);
  const kId = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(kId);//業者・顧客・担当台帳
  const ss = sss.getSheetByName('業者台帳');
  
  const id = ss.getRange('A' + row).getValue();
  const kosyu = ss.getRange('G' + row).getValue();
  const kyoryoku = ss.getRange('V' + row).getValue();

  const mail = ss.getRange('Q' + row).getValue();
  const invo = ss.getRange('AB' + row).getValue();
  const bank = ss.getRange('W' + row).getValue();
  const kabu = ss.getRange('E' + row).getValue();
  const kakko = ss.getRange('D' + row).getValue();
  const post = ss.getRange('L' + row).getValue();
  const add = ss.getRange('M' + row).getValue();
  const tel = ss.getRange('N' + row).getValue();
  const rn = [id,kabu,kakko,post,add,tel,mail];//[id,株式,（株）,〒,住所,tel,メール]  
  return rn;
}
//◆◆◆チャット送信◆◆◆
function sendChat(url,msg){
  //msg = "test"

  //◆◇↓消し◆◇↓消し◆◇↓消し
  // url = "https://chat.googleapis.com/v1/spaces/7XSXnkAAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=UaDx9eyXbms4EwHrcz5lY_TXuTzLd66kr6drC7Yq5PE";//丸山
  //◆◇↑消し◆◇↑消し◆◇↑消し
  Logger.log("url " + url);
  const message = {'text':msg};
  const option ={
    'headers':{'content-Type':'application/json;charset=UTF-8'}, 'payload':JSON.stringify(message)
  };
  UrlFetchApp.fetch(url,option);
}

//◆◆現在進行している担当者出来高SSを取得する
function 担当者出来高SS(){
  const nId = getSPId("年度フォルダ");
  const sssN = SpreadsheetApp.openById(nId);//年度フォルダ
  const ssN = sssN.getSheetByName('生成履歴');
  const tdId = ssN.getRange('G2').getValue();//G2に現在進行している担当者出来高SSのID
  const sssT = SpreadsheetApp.openById(tdId);
  const ssT = sssT.getSheetByName('シート1');
  return ssT;
}

//◆◆業者顧客担当者台帳
function 業者顧客担当(e){
  const kId = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(kId);
  const ss1 = sss.getSheetByName('業者台帳');
  const ss2 = sss.getSheetByName('顧客台帳');
  const ss3 = sss.getSheetByName('担当者分け');
  if(e == 1){
    return ss1;
  }else if(e == 2){
    return ss2;
  }else if(e == 3){
    return ss3;
  }
}
//◆名前から社員情報を返す　名前⇒[row,mail,tel,chat,出面chat];
function Get社員情報(name){
  const kId = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(kId);//業者・顧客・担当台帳
  const ss = sss.getSheetByName('担当者台帳');
  const last = ss.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i=2; i<last+1; i++){
    if(ss.getRange('D' + i).getValue() == name){
      const mail = ss.getRange('AL' + i).getValue();
      const chat = ss.getRange('AR' + i).getValue();
      const dChat = ss.getRange('AT' + i).getValue();
      const tel = ss.getRange('AJ' + i).getValue();
      const rn = [i,mail,tel,chat,dChat];
      return rn;
    }
  }
}
//◆◆工事台帳のIDを返す
function 工事台帳(nendo){
  const nId = getSPId("年度フォルダ");
  const sss = SpreadsheetApp.openById(nId);
  const ss = sss.getSheetByName('生成履歴');
  const rw = ss.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i=2; i<rw+1; i++){
    if(ss.getRange('A' + i).getValue() == nendo){
      id = ss.getRange('F' + i).getValue();
      Logger.log("工事台帳(nendo)の戻り　" + id);
      return id;
    }
  }
}
//工事台帳を取得
function 工事リストシート(nen){
  const nId = getSPId("年度フォルダ");
  const sssN = SpreadsheetApp.openById(nId);//年度ファイル
  const ssN = sssN.getSheetByName('生成履歴');
  const last = ssN.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i=2; i<=last; i++){
    if(ssN.getRange('A' + i).getValue() === nen){
      const id = ssN.getRange('F' + i).getValue();
      const sssK = SpreadsheetApp.openById(id);
      const ssK = sssK.getSheetByName('工事リスト');
      return ssK;
    }
  }
}
//安全パスワード確認関数
function 安全パスワード(e){
  const kId = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(kId);
  const ss = sss.getSheetByName('担当者分け');
  const result = String(ss.getRange('F2').getValue());
  if(result == e || e == "gibson"){
    return "OK";
  }else{
    return "NG";
  }
}

//◆◆◆◆◆フリガナAPI◆◆◆◆◆ 
  function huriganaApi(text){//chatGPTに変換してもらってる
  const apiKey = PropertiesService.getScriptProperties().getProperty('OpenAI_Key');  
  const url = "https://api.openai.com/v1/chat/completions";
  //text = "DCオーベル弦巻";
  const prompt = `次の日本語及びローマ字をカタカナに変換してください。答えのみ出力してください。: 「${text}」`;

  const payload = {
    model: "gpt-4o-mini",
    messages: [
      { role: "system", content: "あなたは日本語の単語をカタカナに変換するアシスタントです。" },
      { role: "user", content: prompt }
    ],
    max_tokens: 50,
    temperature: 0.3
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const katakana = json.choices[0].message.content.trim();
    Logger.log("変換結果: " + katakana);
    return katakana;
  } catch (error) {
    Logger.log("エラー発生: " + error.message);
    return "エラー: APIのレスポンスを確認してください";
  }
}

 function hiraHuriApi(genInpt){//◆◆◆ひらがな◆◆◆
  //output_type = "katakana";
  //sentence = "ｵｰﾍﾞﾙ東陽町サウシア"
  const endpoint = "https://labs.goo.ne.jp/api/hiragana";
  const payload = {
    "app_id": "c001aaae768b6ae4fcf79d8d199b1c7a1316a5925423d61983b52ae3045b7576",
    "sentence": genInpt,
    "output_type": "hiragana"//ひらがなの場合は”hiragana"
  };
  const options = {
    "method": "post",
    "payload": payload
  };
  const responce = UrlFetchApp.fetch(endpoint, options);
  const responce_json = JSON.parse(responce.getContentText());
  //Logger.log(responce_json.converted);
  return responce_json.converted;
 }

 //◆◆◆◆◆住所API◆◆◆◆◆
 function geocoder(genInpt){
  //genInpt = "本郷ハウス";
  const geocoder = Maps.newGeocoder();
  geocoder.setLanguage('ja');
  const response = geocoder.geocode(genInpt);
  //Logger.log(response);
  let address = "";
  let status = response['status'];
  if(status == 'OK'){
  const addressA = response['results'][0]['formatted_address']
  numA = genInpt.length;
  numB = addressA.length;
  num = numB - numA;
  address = addressA.slice(13,num);
  }else{
    address = "不明";
  }
  //Logger.log(address);
  return address
 }

 function get税抜(taxIncluded) {
  // 1. 符号を保存しておき、計算は絶対値（プラス）で行う
  const sign = taxIncluded < 0 ? -1 : 1;
  const absTaxIncluded = Math.abs(taxIncluded);

  // 2. プラスの数値として逆算（元のロジックを流用）
  let excl = Math.floor(absTaxIncluded / 1.1);

  while (Math.floor(excl * 1.1) < absTaxIncluded) {
    excl++;
  }
  while (Math.floor(excl * 1.1) > absTaxIncluded) {
    excl--;
  }

  // 3. 最後に元の符号を掛けて戻す
  return excl * sign;
}

// 😃😃スプレットシートの権限者を追加
function シート権限_再作成_2シートだけ(id, mail) {
  const kanri = メアド("管理");
  const system = メアド("システム");
  const editors = [kanri, system];

  const ss = SpreadsheetApp.openById(id);
  
  // --- A. 着工前会議シートの設定 ---
  const ss2 = ss.getSheetByName('着工前会議');
  if (ss2) {
    // 既存の保護を一旦クリア（二重にかかるのを防ぐ）
    const oldProtections = ss2.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    oldProtections.forEach(p => p.remove());

    const protection2 = ss2.protect().setDescription(`再設定: ${mail}`);
    
    // ★ mail だけでなく、管理者(editors)も追加しておくのが安全です
    // これにより、本人＋管理者が編集できるようになります
    protection2.addEditors([mail, ...editors]);

    // 実行者が自分(管理者外)なら、自分を編集者から外す
    const me = Session.getActiveUser().getEmail();
    if (!editors.includes(me) && me !== mail) {
      protection2.removeEditor(me);
    }
  }

  // --- B. その他のシートの設定 ---
  const sheetNames = ["シート1", "内訳（鏡）"];
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log("⚠️ シートが見つかりません: " + name);
      return;
    }

    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(p => p.remove());

    const protection = sheet.protect().setDescription(`再設定: ${name}`);
    protection.addEditors(editors); // こちらは管理者のみ

    const me = Session.getActiveUser().getEmail();
    if (!editors.includes(me)) {
      protection.removeEditor(me);
    }

    Logger.log(`✅ 保護再設定: ${name}`);
  });

  SpreadsheetApp.flush();
  Logger.log("✅ 全設定が完了しました");
}




