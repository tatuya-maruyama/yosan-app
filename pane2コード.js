//◆小規模注文書管理用
function getGyoListS(fId){
  //fId = "11j_3VeN-d_yfq-E-3tKjyC66_BwyGI5qw5qSSBCIG14";
  const ss1 = 予算書シート(fId);
  let rn1 = [];  
  for(i=2; i<=11; i++){
    if(i == 11){
      Logger.log(rn1);
      return rn1;
    }else{
      let rn2 = [];
      for(let s=1; s<=7; s++){
        if(s == 7){
          rn1.push(rn2);
        }else if(s == 1){
          rn2.push(i);
        }else if(s == 2){
          const val = ss1.getRange('AI' + i).getValue();
          rn2.push(val);
        }else{
          const col = 34 + s;
          const val = ss1.getRange(i,col).getValue();
          rn2.push(val);
        }
      }
    } 
  }
}
//◆◆小規模発注書発行
function syokiboTyumonGo(fId,gyo,gyoRow,kin,tIndex,day,fsRe,kosyu,syatyo){//fsReがreの場合は再発行（gyoRowない）
  Logger.log(fId + "," + gyo + "," + gyoRow + "," + kin + "," + tIndex + "," + day + "," + fsRe + "," + kosyu);
  const ss1 = 予算書シート(fId);
  const kRow = Number(tIndex) + 1;//rowに直すので＋１
  //再発行の場合はgyoRowがないので調べる必要がある
  const gyoRowAF = () => {
    if(gyoRow == "同じ"){//再発行の場合
      return ss1.getRange('AO' + kRow).getValue();
    }else{
      return gyoRow;
    }
  }
  const gyoRowA = gyoRowAF();
  const res = 小規模注文書作成(gyo,gyoRowA,kin,day,fsRe,ss1,kosyu,kRow);//pdfのIDとスプレットシートのIDが返ってくる
  const pdfId = res[0];
  const ssId = res[1];

  const gen = ss1.getRange('H19').getValue();
  const tMail = ss1.getRange('Z19').getValue();
  const tan = ss1.getRange('U19').getValue();

  const gInfo = 業者情報(gyoRowA);//[id,株式,（株）,〒,住所,tel,メール]
  const gName = gInfo[2];//略称
  const gyomei = gInfo[1];//正式
  //◆◇◆◇mentenance◆◇◆◇
  const gMail = gInfo[6];
  //const gMail = "test@anemoworks.com";
  //◆◇◆◇mentenance◆◇◆◇
  const ccF = () => {
    //◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ
    const syatyoMail = メアド("社長");
    //const syatyoMail = "test2@anemoworks.com";
    //◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ
    if(syatyo == true){
      m = tMail + "," + syatyoMail;
      return m;
    }else{
      return tMail;
    }
  } 
  const cc = ccF();

  //hakoBの記録とパスワード記録
  const hakoB = "契約";
  ss1.getRange('AQ' + kRow).setValue(hakoB);
  const ramdam = Math.floor(100000 + Math.random() * 900000); // 100000〜999999
  const code = ramdam.toString();
  ss1.getRange('AR' + kRow).setValue(code);
  ss1.getRange('AT' + kRow).setValue(kosyu);//工種の記録
  
  const subjectF = () => {
    let sub = "【注文書処理ください】" + gen + "（" + kosyu + "）" + gyo + "様";
    //注文書が既に発行しているのか調べる
    const re = ss1.getRange('AK' + kRow).getValue();//""以外は再発行になる
    if(re !== ""){
      const result = "（再発行）" + sub;
      return result;
    }else{
      return sub;
    }
  }
  const subject = subjectF(); 
  const linkPdf = pdfLink("電子契約やり方");//◆やり方説明PDFリンク
  const pdf = DriveApp.getFileById(pdfId);
  const blob = pdf.getBlob();
  const keiyakuLinkF = () => {
    const urlA = getSetUrl("契約書署名");
    const url = urlA + "?param1=" + fId + "&param2=" + gyoRowA + "&param3=" + kRow + "&param4=" + pdfId + "&param5=" + code;
    return url;
  }
  const keiyakuLink = keiyakuLinkF();
  const mailSyomei  = メール署名();

  const body =
    gyomei + "様\n\n" +
    "いつも大変お世話になっております。本日注文書を発行しました。下記リンクより署名処理をお願いいたします\n" +
    "◇契約書に署名する◇\n" + keiyakuLink + "\n\n" +
    "署名方法の説明はこちらから\n" + linkPdf + "\n" +
    "※こちらは自動配信です。問合せについては下記へお願いします\n" +
    mailSyomei;

  // HTML署名（※メール署名(2)がHTMLならそのままでOK）

  // 2-2) 署名（HTML版）※おすすめ：プレーンをHTML化
  const signatureHtml = nl2br_(escapeHtml_(mailSyomei));

  // ★ロゴ
  const LOGO_FILE_ID = ロゴId();
  const hasLogo = !!LOGO_FILE_ID;
  const logoBlob = hasLogo ? DriveApp.getFileById(LOGO_FILE_ID).getBlob() : null;
  const logoHtml = hasLogo
    ? `<br><img src="cid:ishiiLogo" alt="株式会社石井工業" width="240" style="display:block;max-width:240px;height:auto;">`
    : "";

  // htmlBody（ロゴは署名の下に）
  const htmlBody = `
    <p>${gyomei}様</p>
    <p>いつも大変お世話になっております。本日注文書を発行しました。下記リンクより署名処理をお願いいたします</p>
    <p>◇契約書に署名する◇</p>
    <a href="${keiyakuLink}" style="
      display:inline-block;
      padding:12px 24px;
      background-color:#4CAF50;
      color:white;
      text-decoration:none;
      border-radius:6px;
      font-size:16px;">
      契約書に署名する
    </a><br>
    <p>署名方法の説明はこちらから</p>
    <a href="${linkPdf}">説明リンク</a><br>
    ${logoHtml}
    <p>※こちらは自動配信です。問合せについては下記へお願いします</p>
    ${signatureHtml}
  `;

  const systemMail = 管理Mail("契約請求");
  const mailName = 会社名() + "【契約請求】";

  const option = {
    attachments: blob,
    from: systemMail,
    name: mailName,
    htmlBody: htmlBody,
    cc: cc,
    ...(hasLogo ? { inlineImages: { ishiiLogo: logoBlob } } : {})
  };

  GmailApp.createDraft(gMail, subject, body, option);
  
  //◆担当者出来高SSの処理をする
  //担当者出来高SSの修正
  const rn1 = [ss1.getRange('AL' + kRow).getValue(),ss1.getRange('AM' + kRow).getValue(),ss1.getRange('AN' + kRow).getValue(),];
  const rn2 = [];
  const kingaku = [rn1,rn2];
  const kosyuA = "小規模";
  const res2 = tandeki(gen,kosyuA,tMail,gMail,hakoB,gyomei,tan,kingaku);//paneコード.gs,単,金額);//tandeki()はpaneコード.gsにある
  
  const pdfUrl = pdf.getUrl();
  const chat = chatUrl("管理2")//アプリメッセージ(maruyama)
  const msg = "注文請書の署名処理メール送信しました\n現場名：" + gen + "\n業者名：" + gyo + "\n業者メール：" + gMail + "\n担当メール：" + tMail + "\nCCメール：" + cc + "\n\n" + pdfUrl;
  sendChat(chat,msg);

  //発行を記録
  const today = "発行" + Utilities.formatDate(new Date,"JST","M/d");
  const kinA = Number(kin);
  if(fsRe !== "re"){//再発行でない場合
    ss1.getRange('U' + kRow).setValue(gName);//略称
    ss1.getRange('AO' + kRow).setValue(gyoRow);
  }
  ss1.getRange('AL' + kRow).setValue(kinA); 
  ss1.getRange('AK' + kRow).setValue(today);
  ss1.getRange('AS' + kRow).setValue(ssId);//スプレットシートのIDを記録する
  //古い契約書PDFIdがあれば削除
  ss1.getRange('AJ' + kRow).setValue("");//小規模
}

//◆◆小規模注文書作成
function 小規模注文書作成(gyo,gyoRow,kin,day,fsRe,ss1,kosyu,kRow){
  const gen = ss1.getRange('H19').getValue(); 
  const id = ひな形("業者注文請書");
  const sssH = DriveApp.getFileById(id);//注文請書ひな形シート
  //一時フォルダにコピー
  const flId = ドライブId("3ヶ月フォルダ");
  const sFld = DriveApp.getFolderById(flId);//一時フォルダ（３ヶ月）
  const hName = gen + "_" + gyo + "【ひな形】注文請書" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");
  const hId = sssH.makeCopy(hName,sFld).getId();
  const sss = SpreadsheetApp.openById(hId);
  const ssH1 = sss.getSheetByName('注文請書発行（石井⇔業者)ひな型');
  const ssH4 = sss.getSheetByName('変更履歴');

  const gInfo = 業者情報(gyoRow);//[id,株式,（株）,〒,住所,tel,メール]
  const gyomei = gInfo[1];
  const post = gInfo[3];
  const gAdd = gInfo[4];
  const gTel = gInfo[5];
  const gMail = gInfo[6];

  const kId = ss1.getRange('A19').getValue();
  const add = ss1.getRange('T19').getValue();
  const st = ss1.getRange('R19').getValue();
  const fh = ss1.getRange('S19').getValue();
  const kinA = Number(kin);

  const hakkoF = () => {
    if(day == "本日"){
      const d = Utilities.formatDate(new Date,"JST","yyyy年M月d日");
      return d;
    }else{
      return day;
    }
  }
  const hakko = hakkoF();
  ssH1.getRange('M4').setValue(hakko);//発行日
  ssH1.getRange('B5').setValue(gyomei);//業者名（株式）
  ssH1.getRange('B7').setValue("〒" + post);//〒
  ssH1.getRange('B8').setValue(gAdd);//業者住所
  ssH1.getRange('B9').setValue("Tel：" + gTel);//業者Tel
  ssH1.getRange('B10').setValue("mail：" + gMail);//業者メール
  ssH1.getRange('D13').setValue(kId);//工事No
  ssH1.getRange('D14').setValue(gen);//工事名
  ssH1.getRange('D15').setValue(add);//現場住所
  ssH1.getRange('D17').setValue(st);//着工  
  ssH1.getRange('F17').setValue(fh);//完了

  ssH1.getRange('D41').setValue(kosyu);//工種
  ssH1.getRange('M41').setValue(kinA);//金額
  if(fsRe == "fs"){
    ssH1.getRange('O41').setValue("");//工種
  }else if(fsRe == "re"){
    ssH1.getRange('O41').setValue("再発行");//工種
  }
  //業者署名のとこクリア
  ssH1.getRange('J63').setValue("");
  ssH1.getRange('J64').setValue("");
  ssH1.getRange('J64:N64').setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

  ssH1.getRange('J63').setValue("未署名（まだ契約は成立してません)").setFontColor("#FF8C00");

  //◆◆履歴の作成
  const rirekiF = () => {
    const time = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
    let reki = ss1.getRange('AP' + kRow).getValue();
    const hakoA = "契約";
    const newEntry = `${hakoA},${time},,発行,,`;

    if (reki == "") {
      reki = newEntry;
    } else {
      reki = reki + "\n" + newEntry;
    }

    // ここで配列に変換
    let data = reki.split('\n').map(row => row.split(','));

    // 👇データが10行を超えてたら、古いものから削除する
    if (data.length > 10) {
      data = data.slice(data.length - 10); // 最後の10行だけ残す
    }

    // 文字列に戻して再保存
    reki = data.map(row => row.join(',')).join('\n');
    ss1.getRange('AP' + kRow).setValue(reki);

    // 変更履歴シートにも貼り付け
    const row = data.length + 7;
    ssH4.getRange('B8:G' + row).setValues(data);
    ssH4.getRange('C1').setValue(gen + " 【" + kosyu + "】");
    ssH4.getRange('C4').setValue(gyomei);
    ssH4.getRange('D4').setValue(gMail);

    const last = ssH4.getRange('B17').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    for (let i = 13; i <= last; i++) {
      if (ssH4.getRange('E' + i).getValue() == "プロセス完了") {
        ssH4.getRange('E' + i).setFontColor("#1e90ff").setFontWeight("bold"); // ←ここ直しました
      }
    }

    return "OK";
  };
  rirekiF();//履歴の実行

  SpreadsheetApp.flash;
  const nendo = ss1.getRange('A18').getValue();
  const pdfId = 小規模注文PDF(hId,gyomei,gen,nendo);
  return [pdfId,hId];
}

//◆◆小規模注文書PDF作成
function 小規模注文PDF(hId,gyomei,gen,nendo){
  //注文書S請書★パシフィック魚藍坂【S】株式会社リプル様20240101
  const fn = "注文書S請書★" + gen + "【S】" + gyomei + "様" + Utilities.formatDate(new Date,"JST","yyyyMMdd");
  const anTyu = "発注";
  const save = 保存フォルダ(nendo,gen,anTyu);
  
  const sheets = SpreadsheetApp.openById(hId);
  const ssH1 = sheets.getSheetByName('業者案内用内訳');
  const ssH2 = sheets.getSheetByName('業者案内用内訳 （追加）');

  sheets.deleteSheet(ssH1);//内訳をを削除
  sheets.deleteSheet(ssH2);//内訳をを削除

  const pdf = sheets.getAs('application/pdf');
  pdf.setName(fn);
  const pId = save.createFile(pdf).getId();
  return pId;//Idを返す
}
