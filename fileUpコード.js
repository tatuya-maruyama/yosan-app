//◆◆工程表、仕様書アップロード用
function onUpload(data, fileName, mimeType, upfileNo,id) {
  Logger.log("upfileNo " + upfileNo);
  const ss = 内訳鏡(id);
  const ss1 = 予算書シート(id);
  const gen = ss1.getRange('H102').getValue();
  const fn = "file1◆" + gen + "_" + fileName;
  const blob = Utilities.newBlob(Utilities.base64Decode(data.substr(data.indexOf('base64,')+7)), mimeType).setName(fn);
  //ファイルの保存先
  const fld = get添付ファイルフォルダ();
  const fId = fld.createFile(blob).getId();//ファイルIDを取得
  Logger.log("fId " + fId );
  if(upfileNo == "flUp1"){
    ss.getRange('V1').setValue(fId);
  }else if(upfileNo == "flUp2"){
    ss.getRange('W1').setValue(fId);
  }
  return fId;
}
  
//添付ファイルフォルダIDを取得
function get添付ファイルフォルダ(){
  const nenTuki = 年度年月();
  const nendo = nenTuki[0];
  const nendoFldName = nendo + "年";
  const oyaId = ドライブId("注文請書");//注文請書ドライブ
  const oya = DriveApp.getFolderById(oyaId);
  
  const nenFldF = () => {
    const folders = oya.getFoldersByName(nendoFldName);
    if(folders.hasNext()){
      const fld = folders.next();
      return fld;
    }else{
      const fld = oya.createFolder(nendoFldName);
      return fld;
    }
  }
  const nenFld = nenFldF();
  const tenpF = () => {
    const name = "添付フォルダ";
    const folders2 = nenFld.getFoldersByName(name);
    if(folders2.hasNext()){
      const fld = folders2.next();
      return fld;
    }else{
      const fld = nenFld.createFolder(name);
      return fld;
    }
  }
  const tenp = tenpF();
  return tenp;
}
//😃😃内訳のアップロード用まず保存して画像を表示させる
function convertExcelToPdfBase64(base64, fileName) {
  const decoded = Utilities.base64Decode(base64);
  
  // ファイル名から拡張子を判別して適切なMimeTypeを設定
  let contentType = MimeType.MICROSOFT_EXCEL; // デフォルト
  if (fileName.endsWith('.xlsm')) {
    contentType = "application/vnd.ms-excel.sheet.macroEnabled.12";
  } else if (fileName.endsWith('.xlsx')) {
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }

  const blob = Utilities.newBlob(decoded, contentType, fileName);
  const excelFile = DriveApp.createFile(blob);

  try {
    // 変換設定（Drive API v2）
    const resource = {
      title: "PDF変換_" + fileName,
      mimeType: MimeType.GOOGLE_SHEETS
    };

    // 第3引数に {convert: true} を指定することで Sheets に変換
    const converted = Drive.Files.copy(resource, excelFile.getId(), { convert: true });
    
    // --- 以降、スプレッドシートの編集とPDF化 ---
    const sss = SpreadsheetApp.openById(converted.id);
    const sheet = sss.getSheets()[0];
    const sheetId = sheet.getSheetId();

    // 1行目にラベル挿入（A, B, C...）
    const header = Array.from({length: 13}, (_, i) => String.fromCharCode(65 + i));
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, header.length).setValues([header])
         .setFontWeight("bold").setFontColor("red").setFontSize(14);

    SpreadsheetApp.flush();

    // PDFエクスポート
    const exportUrl = `https://docs.google.com/spreadsheets/d/${converted.id}/export?format=pdf&gid=${sheetId}&range=A1:M30&portrait=true`;
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
    });

    const base = Utilities.base64Encode(response.getBlob().getBytes());
    
    // 変換用の一時ファイルは削除
    excelFile.setTrashed(true);

    return JSON.stringify({ base: base, fileId: converted.id });

  } catch (e) {
    console.error("変換エラー: " + e.message);
    return JSON.stringify({ error: e.message });
  }
}