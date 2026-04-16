//😃😃契約書アップロード一時保管
function uploadToDrive(base64Data, fileName, contentType) {
  Logger.log(contentType);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, fileName);
  const idD = ドライブId("一時フォルダ丸山");
  const folder = DriveApp.getFolderById(idD);//一時フォルダに保管
  const file = folder.createFile(blob);
  const id = file.getId();
  Logger.log("uploadToDrive " + id);
  return id// ファイルIDを返す
}
//😃😃契約書アップロード😃😃
function onUpload2(upKeiyaku, fileName, mimeType, fId, which, sct,resizedDataUrl,parsedInfo) {
  Logger.log("parsedInfo " + parsedInfo);////金額,着工,完了,発注元,発行日,住所,担当,物件名
  Logger.log("which " + which);
  const vals = parsedInfo.join('◆');//解析内容読み込み用
  if(which == "csvBt"){
    // const fn = "hasekoCSV" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");
    // const blob = Utilities.newBlob(Utilities.base64Decode(data.substr(data.indexOf('base64,')+7)), mimeType).setName(fn);
    // //ファイルの保存＆IDを取得
    // const fldId = "12xqAGM_8vlz2LyboNvOAEXkG396-P0OO"//一時フォルダ
    // const folder = DriveApp.getFolderById(fldId);
    // const fileId = folder.createFile(blob).getId();

    //ﾌｧｲﾙIDを記録　年度ファイルに記録その都度書き換えられる
    const idN = getSPId("年度フォルダ");
    const sss = SpreadsheetApp.openById(idN);//年度ファイル
    const ss = sss.getSheetByName('生成履歴');
    Logger.log(ss.getRange('U1').getValue() + "," + fId);
    ss.getRange('U1').setValue(upKeiyaku);

  }else if(which == "pdfBt"){//😃注文書の登録（まだ登録されていなかった場合と変更の場合）
    const sss = SpreadsheetApp.openById(fId);
    const ss = sss.getSheetByName('シート1');
    const ds = (ss.getRange('C102').getValue() !== "")? ss.getRange('C102').getValue():ss.getRange('C19').getValue();
    const gen = (ds == "D" || ds == "S2")? ss.getRange('H102').getValue():ss.getRange('H19').getValue();
    const nendo = (ds == "D" || ds == "S2")? ss.getRange('D102').getValue():ss.getRange('D19').getValue();
    //ファイルの保存先
    const fld = get注文請書Fld(nendo,gen,ds);
    const file = DriveApp.getFileById(upKeiyaku);//アップロードファイルId
    file.moveTo(fld);

    const kinRaw = parsedInfo[0];
    // 1. 😃😃数字とマイナス記号以外を削除（先頭以外のハイフンも残る可能性があるため注意）
    let kinSanitized = kinRaw.replace(/[^-0-9]/g, ""); 
    // 2. 数値に変換
    // ※もし "- 1,000" のように間にスペースがあると "--1000" になる可能性があるため、
    //   Number() で判定する前に念のため文字として整えるとより安全です。
    const kinValue = Number(kinSanitized);
    // 3. 税抜計算（数値として渡す）
    const kin = get税抜(kinValue);
    //数値化できたかをチェック（必要に応じて）
    if (isNaN(kin)) {
      console.warn("金額が数値に変換できませんでした:", kinRaw);
      parsedInfo[0] = 0; // または null にするな)
    } else {
      parsedInfo[0] = kin;
    }
    //予算書に記録  
    if(ds == "D" || ds == "S2"){
      const row = Number(sct) + 87;
      ss.getRange('L' + row).setValue(upKeiyaku);
      ss.getRange('M' + row).setValue(kin);
      ss.getRange('AR' + row).setValue(vals);
      ss.getRange('J102').setFormula("=K87");//請負金額の数式をセット
      //😃😃最後にスペースへチャット送信
      const genmei = parsedInfo[7];
      const kinds = ss.getRange('I102').getValue();
      const cleanUke = String(kin).replace(/[\n,]/g, "");
      tyumonSpace(upKeiyaku, cleanUke, genmei, kinds);//コードyosanRezi.jsにある
      return [upKeiyaku,kin];
    }else if(ds == "S"){
      ss.getRange('T16').setValue(upKeiyaku);
      ss.getRange('J19').setValue(kin);
      ss.getRange('T17').setValue(vals);
      return [upKeiyaku,kin];
    }
  }else if(which == "newPdf"){//😃注文書の登録（追加の登録）
    Logger.log(which);
    const sss = SpreadsheetApp.openById(fId);
    const ss = sss.getSheetByName('シート1');
    const ds = (ss.getRange('C102').getValue() !== "")? ss.getRange('C102').getValue():ss.getRange('C19').getValue();
    const gen = (ds == "D" || ds == "S2")? ss.getRange('H102').getValue():ss.getRange('H19').getValue();
    const nendo = (ds == "D" || ds == "S2")? ss.getRange('D102').getValue():ss.getRange('D19').getValue();
    const folder = get注文請書Fld(nendo,gen,ds);
    const file = DriveApp.getFileById(upKeiyaku);
    file.moveTo(folder);
    if (which === "newPdf" && resizedDataUrl) {
      // OCR処理へ（例: Vision API or Drive OCR）
      // const ocrText = extractTextFromImage(imageBlob); // OCR処理関数を呼ぶ
      // const gpt = analyzePdfText(ocrText); // ChatGPT的な要約処理
      //↑はすでにやってる
      if(ds == "D" || ds == "S2"){
        const last = ss.getRange('K101').getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
        const num = last - 87;
        const name = (num == 1)? "注文書1（本工事）": "注文書" + num;
        const kinRaw = parsedInfo[0];
        // 1. 😃😃数字とマイナス記号以外を削除（先頭以外のハイフンも残る可能性があるため注意）
        let kinSanitized = kinRaw.replace(/[^-0-9]/g, ""); 
        const kin = get税抜(Number(kinSanitized));//税抜きにする
        // 2. 数値化できたかをチェック（必要に応じて）
        if (isNaN(kin)) {
          console.warn("金額が数値に変換できませんでした:", kinRaw);
          parsedInfo[0] = 0; // または null にするなど
        } else {
          parsedInfo[0] = kin;
        }
        ss.getRange('K' + last).setValue(name);
        ss.getRange('M' + last).setValue(kin);
        ss.getRange('L' + last).setValue(upKeiyaku);
        ss.getRange('AR' + last).setValue(vals);
        ss.getRange('J102').setFormula("=K87");//請負金額の数式をセット
        
        parsedInfo.push(upKeiyaku);
        parsedInfo.push(num);
        parsedInfo.push(name);
        Logger.log(parsedInfo);//[kin,start,finish,client,day,add,tan,fileId,num,name];
        const genmei = parsedInfo[7];
        const kinds = ss.getRange('I102').getValue();
        const cleanUke = String(kin).replace(/[\n,]/g, "");
        tyumonSpace(upKeiyaku, cleanUke, genmei, kinds);//コードyosanRezi.jsにある
        return parsedInfo;
      }else{

      }
    }
  } else if (which == "reziPdf") {
  Logger.log("Which = reziPdf");
  Logger.log("parsedInfo = " + JSON.stringify(parsedInfo));

  if (!parsedInfo || parsedInfo.length < 1) {
    Logger.log("parsedInfoが無効です");
    return null;
  }

  const kinRaw = parsedInfo[0];
  // 1. 😃😃数字とマイナス記号以外を削除（先頭以外のハイフンも残る可能性があるため注意）
  let kinSanitized = kinRaw.replace(/[^-0-9]/g, ""); 
  const kin = get税抜(Number(kinSanitized));

  if (isNaN(kin)) {
    Logger.log("金額の変換失敗: " + kinRaw);
    parsedInfo[0] = 0;
  } else {
    parsedInfo[0] = kin;
  }

  Logger.log("returning parsedInfo = " + JSON.stringify(parsedInfo));
  return parsedInfo;
}

  //return "OK";
}
function get注文請書Fld(nendo, gen, ds) {
  const genTukiArr = genTuki(gen);
  const genmei = genTukiArr[0];
  const tuki = genTukiArr[1];
  const idD = ドライブId("注文請書");
  const oya = DriveApp.getFolderById(idD); // 注文請書ドライブ

  // フォルダ取得（なければ作成）
  const getFld = (name, parent) => {
    const flds = parent.getFoldersByName(name);
    return flds.hasNext() ? flds.next() : parent.createFolder(name);
  };

  const nenFld = getFld(nendo + "年", oya);

  if (ds === "D") {
    const fldD = getFld("注文書大規模（ｽﾍﾟｰｽにもあげる）", nenFld);
    return getFld(genmei, fldD);
  } else if (ds === "S" || ds === "S2") {
    const fldS = getFld("注文書小規模（ｽﾍﾟｰｽにもあげる）", nenFld);
    const tukiFld = getFld(tuki + "月", fldS);
    return getFld(genmei, tukiFld);
  }

  // 想定外のdsのとき
  return null;
}

//小規模の現場名から現場名と月を取り出す
function genTuki(gen){
  //gen = "加藤　隆康　邸 192-5月";
  if(gen.indexOf("-") == -1){
    return [gen,""];//大規模の場合
  }else{
    const name = gen.slice(0,gen.indexOf(" "));
    const tuki = gen.slice(gen.indexOf("-") + 1,-1);
    const res = [name,tuki]; 
    Logger.log(res);
    return res;
  }
}

//↑で登録したIDをもとにCSVの情報を取得
function getCSVdata(ds,upKeiyaku){
  // const sss = SpreadsheetApp.openById('1yHLjRhNNS-Ht22JltcfxBb8G6_YaEjHTxI-ExL58ZcI');//年度ファイル
  // const ss = sss.getSheetByName('生成履歴');
  // const id = ss.getRange('U1').getValue();
  // Logger.log("getCSVdata() " + id);
  const blob = DriveApp.getFileById(upKeiyaku).getBlob();
  const file = blob.getDataAsString('MS932');
  const values = Utilities.parseCsv(file);
  const idC = ひな形("CSV変換用");
  const sssCSV = SpreadsheetApp.openById(idC);//CSV変換用SSファイル（管理フォルダのひな形フォルダ内）
  const csv = sssCSV.getSheetByName('シート1');
  //一旦リセット
  csv.clear();
  csv.getRange(1, 1, values.length, values[0].length).setValues(values);

  const genkinds = csv.getRange('AC2').getValue();//現場名
  const gen = genkinds.slice(0, genkinds.lastIndexOf('　'));
  const kinds = genkinds.slice(genkinds.lastIndexOf('　') + 1);
  const kin = csv.getRange('AM2').getValue();//金額（別）

  const eikou = csv.getRange('AS2').getValue();//備考: 
    //"調査（改行）営業担当：佐野　映　工事担当：佐野　映
  const tanA = eikou.slice(eikou.lastIndexOf('：') + 1);
  const tan = tanA.replace(/[\r\n]+/g,"");
  const idH = ひな形("長谷工完了報告");
  const sssH = SpreadsheetApp.openById(idH);

  const rw = csv.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  Logger.log(rw);

  //メールアドレス検索
  const ssM = sssH.getSheetByName('メールアドレス');
  const last2 = ssM.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const end = last2 + 1;
  const mailF = (tan) => {
      for(let i=2; i<last2+2; i++){
        if(i == end){
          return "なし";
        }else if(ssM.getRange('A' + i).getValue() == tan){
          const m = ssM.getRange('C' + i).getValue();
          return m;
        }
      }
    }
  if(rw == 2){//CSVデータが一つしかない場合
    //現場名
    const num = csv.getRange('A2').getValue();//発注番号
    const add = csv.getRange('AF2').getValue();//住所
    const startA = csv.getRange('AI2').getValue();//開始
    const start = Utilities.formatDate(startA,"JST","yyyy-MM-dd");
    const finishA = csv.getRange('AJ2').getValue();//終了
    const finish = Utilities.formatDate(finishA,"JST","yyyy-MM-dd");
    const keiyakuA = csv.getRange('AT2').getValue();//契約日（請日）
    const keiyaku = Utilities.formatDate(keiyakuA,"JST","yyyy-MM-dd");
    
    const mail = mailF(tan);
    //記録発注番号がなかったら記録
    const ssH = sssH.getSheetByName("発注管理表");//長谷工の工事完了報告管理表（【管理】請求支払にある）
    const last = ssH.getRange('W1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
    for(let i=2; i<last+1; i++){
      if(i == last){//記録なし
        const data = csv.getRange('A2:AU2').getValues();
        ssH.getRange('W' + last + ":BQ" + last).setValues(data);
        //DSの記録
        if(ds == "D"){
          ssH.getRange('A' + last).setValue("大規模");
        }else if(ds == "S"){
          ssH.getRange('A' + last).setValue("小規模");
        }
        //スプレットシート記録
        ssH.getRange('B' + last).setValue(gen);
        ssH.getRange('C' + last).setValue(kinds);
        ssH.getRange('D' + last).setValue(kin);
        ssH.getRange('BS' + last).setValue(tan);//工事担当
      
        const valA = num + "," + gen + "," + kinds + "," + add + "," + start + "," + finish + "," + kin + "," + keiyaku + "," + tan + "," + mail;
        Logger.log(valA);
        return valA;
      }else if(ssH.getRange('W' + i).getValue() == num){
        const valA = num + "," + gen + "," + kinds + "," + add + "," + start + "," + finish + "," + kin + "," + keiyaku + "," + tan + "," + mail;
        Logger.log(valA);
        return valA;
      }
    }  
  }else{//CSVデータが複数ある
    //一応記録はしておく
    //複数ある現場を選択してもらうため返す
    const ssH = sssH.getSheetByName("発注管理表");//長谷工の工事完了報告管理表（【管理】請求支払にある）
    const namA = rw - 1;
    let vals = "複数," + namA;
    const end2 = rw + 1
    for(let i=2; i<rw+2; i++){
      if(i == end2){
        return vals;
      }else{
          last = ssH.getRange('W1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
          for(let s=2; s<last+1; s++){
            hNum = csv.getRange('A' + i).getValue();
            if(s == last){
              const data = csv.getRange('A' + i + ':AU' + i).getValues();
              ssH.getRange('W' + last + ":BQ" + last).setValues(data);
              //DSの記録
              if(ds == "D"){
                ssH.getRange('A' + last).setValue("大規模");
              }else if(ds == "S"){
                ssH.getRange('A' + last).setValue("小規模");
              }
              //スプレットシート記録
              let genkinds2 = csv.getRange('AC' + i).getValue();//現場名
              let gen2 = genkinds2.slice(0, genkinds2.lastIndexOf('　'));
              let kinds2 = genkinds2.slice(genkinds2.lastIndexOf('　') + 1);
              let kin2 = csv.getRange('AM' + i).getValue();//金額（別）

              let eikou2 = csv.getRange('AS' + i).getValue();//備考: 
                //"調査（改行）営業担当：佐野　映　工事担当：佐野　映
              let tan2 = eikou2.slice(eikou2.lastIndexOf('：') + 1);

              ssH.getRange('B' + last).setValue(gen2);
              ssH.getRange('C' + last).setValue(kinds2);
              ssH.getRange('D' + last).setValue(kin2);
              ssH.getRange('BS' + last).setValue(tan2);//工事担当
            }else if(ssH.getRange('W' + s).getValue() == hNum){
              break;
            }
          }
        const uke = csv.getRange('AM' + i).getValue() + "円"; 
        val = csv.getRange('AC' + i).getValue() + "◆" + uke + "◇" + csv.getRange('A' + i).getValue() ;
        vals = vals + "," + val;
      }
    }
  }
}
