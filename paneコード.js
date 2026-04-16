//◆◆現場名取得（window.onload)
function genInfo(param1,param2){
  const ss1 = 予算書シート(param1);
  if(param2 == "S"){
    const gen = ss1.getRange('H19').getValue();
    const tan = ss1.getRange('U19').getValue();
    return gen + "," + tan;
  }else{
    const gen = ss1.getRange('H102').getValue();
    const tan = ss1.getRange('U102').getValue();
    return gen + "," + tan;
  }
}

//◆◆パネル表示
function getKosyuList(fId){
  const ss = 内訳鏡(fId);
  const ss1 = 予算書シート(fId);
  const sheets = SpreadsheetApp.openById(fId);
  const ssR = sheets.getSheetByName("色管理");

  if(ss1.getRange('D85').getValue() == 0){//まだ内訳登録誰てない
    throw new Error("内訳の登録がまだのようです");
  }

  const last = ss.getRange('T3').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const lastR = ssR.getRange('V100').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  const fl1 = ss.getRange('V1').getValue();
  const fl2 = ss.getRange('W1').getValue();
  const ccMail = ss.getRange('X1').getValue();
  const tan = ss1.getRange('U102').getValue();
  const rn1F = () => {
    const val = ss.getRange('T' + last).getValue();
    if(isNaN(val)){
      return ss.getRange('T3:AH' + last).getValues();
    }else{
      return "なし";
    }
  }
  const rn1 = rn1F();

  const rn2F = () => {
    if(lastR == 2){
      return "なし";
    }else{
      return ssR.getRange('T3:AH' + lastR).getValues();
    }
  }
  const rn2 = rn2F();
  const data = { fl1, fl2, ccMail, tan, rn1, rn2 };
  return JSON.stringify(data);
  
  // const end = last + 1;
  // let vals = "";
  // let k = 0;
  // for(let i=3; i<=end; i++){
  //   if(i == end){
  //     const fl1 = ss.getRange('V1').getValue();
  //     const fl2 = ss.getRange('W1').getValue();
  //     const ccMail = ss.getRange('X1').getValue();
  //     const tan = ss1.getRange('U102').getValue();
  //     vals = k + "," + vals + "," + fl1 + "," + fl2 + "," + tan + "," + ccMail;
  //     Logger.log(vals);
  //     return vals;
  //   }else if(ss.getRange('U' + i).getValue() !== 0){  
  //     k = k + 1;
  //     const kosyu = ss.getRange('T' + i).getValue();
  //     const gyo = ss.getRange('V' + i).getValue();
  //     const uke = ss.getRange('X' + i).getValue();
  //     const ann = ss.getRange('AH' + i).getValue();
  //     const hatyu = ss.getRange('AF' + i).getValue();
  //     const val = i + "," + kosyu + "," + gyo + "," + uke + "," + ann + "," + hatyu;
  //     if(k == 1){
  //       vals = val;
  //     }else{
  //       vals = vals + "," + val;
  //     }  
  //   }
  // }
}
//◆◆業者頭文字取得
function gyoKasira(){
  const id = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(id);//業者・顧客・担当台帳
  const ss = sss.getSheetByName('業者分け（検索表すべて）');
  const last = ss.getRange('CA2').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const val = ss.getRange('CA2:CA' + last).getValues();
  const num = last - 1;
  const vals = num + "," + val;
  return vals; 
}
//◆◆業者頭文字から業者リスト取得
function gyoKashiraList(kasira){
  const id = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(id);//業者・顧客・担当台帳
  const ss = sss.getSheetByName('業者分け（検索表すべて）');
  const last = ss.getRange('CC1').getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  for(let i=81; i<=last; i++){
    if(ss.getRange(1,i).getValue() == kasira){
      const data = ss.getRange(2,i,40,1).getValues();
      const rw = data.filter(data => data[0] !== "").length;
      let vals = rw;
      vals = vals + "," + ss.getRange(2,i,rw).getValues();//名前
      vals = vals + "," + ss.getRange(45,i,rw).getValues();//ROW
      Logger.log(vals);
      return vals;
    }
  }
}
//◆◆業者選択してOKをクリックした処理
function kosyuGyoRezi(gyoRow,gyo,kin,kosyuRow,fId){
  //業者名は略称をセット
  const gRow = Number(gyoRow);
  const gInfo = 業者情報(gRow);//[id,株式,（株）,〒,住所,tel,メール]
  const gyomei = gInfo[2];
  const ss = 内訳鏡(fId);
  const ss1 =  予算書シート(fId);
  const row = Number(kosyuRow);
  ss.getRange('V' + row).setValue(gyomei);
  ss.getRange('AE' + row).setValue(gyoRow);
  const num = (Number(ss.getRange('U' + row).getValue()) * 3 ) + 1;
  ss1.getRange('AS' + num).setValue(gyoRow);
  const res = "OK," + gyomei;//略称を返す 
  return res;
}

//◆◆案内発注管理パネル用データ読み込み
function gyoAnTyu(kosyuRow,anTyu,fId){
  const ss = 内訳鏡(fId);
  const row = Number(kosyuRow);
  const uke = ss.getRange('X' + row).getValue();//発注金額
  const stateU = ss.getRange('AF' + row).getValue();//状況（契約)
  const zogen = ss.getRange('AB' + row).getValue();//増減金額
  const stateZ = ss.getRange('AF' + row).getValue();//状況（増減）
  const tuika = ss.getRange('AD' + row).getValue();//追加金額
  const stateT = ss.getRange('AG' + row).getValue();//状況（追加）
  const anK = ss.getRange('AH' + row).getValue();//契約案内状況
  const henK = ss.getRange('AI' + row).getValue();//変更案内状況
  const tuiK = ss.getRange('AJ' + row).getValue();//追加案内状況
  if(anTyu == "発注"){
    const vals = uke + "," + stateU + "," + zogen + "," + stateZ + "," + tuika + "," + stateT;
    return vals
  }else if(anTyu == "案内"){
    const vals = uke + "," + anK + "," + zogen + "," + henK + "," + tuika + "," + tuiK;
    return vals
  }
}

//◆◆内訳（案内・発注)を読み込み
function anTyuKakunin(kosyu,fId){
  const ss = 内訳鏡(fId);
  const last = ss.getRange('E2100').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const dataA = ss.getRange('A24:U' + last).getValues();
  let data = dataA.filter(record => record[11] == kosyu);
  return data;
}

//◆社員メールリスト読み込み
function syaMailList(){
  const id = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(id);//業者・顧客・担当台帳
  const ss = sss.getSheetByName("担当者分け");
  const last = ss.getRange('H2').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const num = last - 1;
  const val1 = ss.getRange('H2:H' + last).getValues();//名前
  const val2 = ss.getRange('K2:K' + last).getValues();//メール
  const vals = num + "," + val1 + "," + val2;
  return vals;
}

//◆◆ccメールの追加
function ccmailadd(mail,fId){
  const ss = 内訳鏡(fId);
  const ccMail = ss.getRange('X1').getValue();
  if(ccMail == ""){
    ss.getRange('X1').setValue(mail);
    return mail;
  }else{
    const result = ccMail + "," + mail;
    ss.getRange('X1').setValue(result);
    return result;
  }
}
//◆ccメールのリセット
function ccmailReset(fId){
  const ss = 内訳鏡(fId);
  ss.getRange('X1').setValue("");
  return "OK";
}

//◆◆◆注文書＆案内発行◆◆◆
function anTyuHako(hako,anTyu,kosyuRow,re,message,tenp,syatyo,day,fId,hakoB){//hakoは契約,変更,追加　hakoBはhakoに”両方"が追加されてる hakoBは両方の時は追加も変更もあるってことです
  Logger.log("hako " + hako + ", anTyu " + anTyu + ", kosyuRow " + kosyuRow + ", re " + re + ",message " + message + ", tenp " + tenp + ", syatyo " + syatyo + ", day " + day + ", fId " + fId + ", hakoB " + hakoB);
  //try{
    const id = 注文請書SS作成(hako,kosyuRow,day,fId,hakoB);
    Logger.log("注文請書SS作成 " + id);
    const pdf = PDF作成(id,hako,anTyu,kosyuRow,fId); //DriveAppオブジェクト
    const ss = 内訳鏡(fId);
    const ss1 = 予算書シート(fId);
    const kRow = Number(kosyuRow);
    const gyoRow = ss.getRange('AE' + kRow).getValue();
    const kosyu = ss.getRange('T' + kRow).getValue();
    const gInfo = 業者情報(gyoRow)//[id,株式,（株）,〒,住所,tel,メール]  
    const gyomei = gInfo[1];
    const gyo = gInfo[2];

    const gMail = gInfo[6];
    //const gMail = "system.ishii@ebisu-ishii.co.jp";//メンテナンス用

    Logger.log(gyomei + "," + gyo + "," + gMail);
    const gen = ss1.getRange('H102').getValue();
    const kinds = ss1.getRange('I102').getValue();
    const genmei = gen + " " + kinds;
    const tan = ss1.getRange('U102').getValue();
    const nendo = ss1.getRange('D102').getValue();
    const tMail = ss1.getRange('Z102').getValue();
    const add = ss1.getRange('T102').getValue();
    const st = ss1.getRange('R102').getValue();
    const fh = ss1.getRange('S102').getValue();
    const tTel = ss1.getRange('Y102').getValue();
    const koki = Utilities.formatDate(st,"JST","yyyy年M月d日") + "～" + Utilities.formatDate(fh,"JST","yyyy年M月d日");
    const ccF = () => {
      //◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ
      const syatyoMail = 管理Mail("社長");
      //const syatyoMail = "test2@anemoworks.com";
      //◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ
      let ccM = ss.getRange('X1').getValue();
      if(ccM !== ""){
        if(syatyo == true){
          res = tMail + "," + ccM + "," + syatyoMail;
          return res;
        }else{
          return tMail + "," + ccM;
        }
      }else{
        if(syatyo == true){
          res = tMail + "," + syatyoMail;
          return res;
        }else{
          return tMail;
        }
      }
    } 
    const cc = ccF();
    if(anTyu == "案内"){
      const fn = "【内訳】" + gyomei + "様★" + gen + Utilities.formatDate(new Date,"JST","yyMMdd") + ".xlsx";
      const excel = Excel生成(id,fn,nendo,gen,anTyu);

      const attF = () => {//添付ファイル
        let rn = [pdf,excel];
        if(tenp == "添付する"){
          const f1 = ss.getRange('V1').getValue();
          const f2 = ss.getRange('W1').getValue();
          if(f1 !== ""){
            const file1 = DriveApp.getFileById(f1);
            rn.push(file1);
          }
          if(f2 !== ""){
            const file2 = DriveApp.getFileById(f2);
            rn.push(file2);
          }
          return rn;
        }else{
          return rn;
        }
      }
      const att = attF();

      const subF = () => {//件名
        let subA = gen + kosyu + "_" + gyo;
        if(hako == "契約"){
          subA = "【案内】" + subA;
          return subA;
        }else if(hako == "変更"){
          subA = "【案内（変更）】" + subA;
          return subA;
        }else if(hako == "追加"){
          subA = "【案内（追加）】" + subA;
          return subA;
        }
      }
      const sub = subF();
      const subjectF = () => {
        if(re == "済"){
          return "(再発行)" + sub;
        }else{
          return sub;
        }
      }
      const subject = subjectF();
      Logger.log(subject);

      // 署名（プレーン）
      const mailSyomei = メール署名();
      // 署名（HTML）
      const signatureHtml = nl2br_(escapeHtml_(mailSyomei));

      // ロゴ
      const LOGO_FILE_ID = ロゴId();
      const hasLogo = !!LOGO_FILE_ID;
      const logoBlob = hasLogo ? DriveApp.getFileById(LOGO_FILE_ID).getBlob() : null;
      const logoHtml = hasLogo
        ? `<br><img src="cid:ishiiLogo" alt="株式会社石井工業" width="240" style="display:block;max-width:240px;height:auto;">`
        : "";

      // 本文（プレーンを先に作る）
      const bodyPlain = (() => {
        const head =
          `${gyomei}様\n\n` +
          `いつもお世話になっております　現場案内を送信しました\n\n`;

        const msgPart = (message && message !== "") ? `${message}\n\n` : "";

        const tail =
          `工事名：${genmei}\n` +
          `住所：${add}\n` +
          `工期：${koki}\n` +
          `担当者：${tan}\n` +
          `${tTel}\n` +
          `${tMail}\n` +
          `詳細は案内記載の弊社担当までお問い合わせください。\n` +
          `内訳内容につきましては弊社の標準単価にて作成しております。金額及び必要な項目等変更が必要な場合はお見積り書を提出ください\n` +
          `【トラブル多くなってます】現場を見に行く場合は、必ず打ち合わせください。\n` +
          `（建物撮影には許可が必要）\n\n`;

        return head + msgPart + tail;
      })();

      // HTML化（ここで初めて変換）
      const bodyHtml = nl2br_(escapeHtml_(bodyPlain));

      const htmlBody = `
        ${bodyHtml}
        ${logoHtml}
        <p>※こちらは自動配信です　問合せ及びお見積りは下記へお願いします</p>
        ${signatureHtml}
      `;

      const systemMail = 管理Mail("契約請求");
      const mailName = 会社名() + "【契約請求】";

      const option = {
        attachments: att,
        cc: cc,
        from: systemMail,
        name: mailName,
        htmlBody: htmlBody,
        ...(hasLogo ? { inlineImages: { ishiiLogo: logoBlob } } : {})
      };

      // ✅ プレーン本文は bodyPlain を渡す
      GmailApp.sendEmail(gMail, subject, bodyPlain + "\n" + mailSyomei, option);


      //済処理
      if(hako == "契約"){
        ss.getRange('AH' + kRow).setValue("済");
      }else if(hako == "変更"){
        ss.getRange('AI' + kRow).setValue("済");
      }else if(hako == "追加"){
        ss.getRange('AJ' + kRow).setValue("済");
      }

      const chat = chatUrl("管理2")//アプリメッセージ(maruyama)
      const msg = "◆案内発行成功⇒" + gyo + "\n" + gen;
      sendChat(chat,msg);

    }else if(anTyu == "発注"){
      //hakoBの記録とパスワード記録
      ss.getRange('AL' + kRow).setValue(hakoB);
      const ramdam = Math.floor(100000 + Math.random() * 900000); // 100000〜999999
      const code = ramdam.toString();
      ss.getRange('AM' + kRow).setValue(code);

      const subF = () => {
        let subA = "【注文書処理ください】" + genmei + "（" + kosyu + "）" + gyo + "様";
        if(hako == "契約"){
          return subA;
        }else if(hako == "変更"){
          subA = "精算変更" + subA;
          return subA;
        }else if(hako == "追加"){
          subA = "追加変更" + subA;
          return subA;
        }
      }
      const sub = subF();
      const subjectF = () => {
        if(re == "済"){
          const result = "（再発行）" + sub;
          return result;
        }else{
          return sub;
        }
      }
      const subject = subjectF(); 
      const linkPdf = pdfLink("電子契約やり方");//◆やり方説明PDFリンク
      const pdfId = pdf.getId();
      const blob = pdf.getBlob();
      const urlA = getSetUrl("契約書署名");
      const keiyakuLinkF = () => {
        const url = urlA + "?param1=" + fId + "&param2=" + gyoRow + "&param3=" + kRow + "&param4=" + pdfId + "&param5=" + code;
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
        <p>${message}</p>
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


      //発行を記録
      const num = ss.getRange('U' + kRow).getValue();//工種番号１から１３（ROWを割り出せる）
      const r = num * 3;
      if(hako == "契約" || hako == "変更"){
        const col = 3 + (num * 3);
        const val = "発行" + Utilities.formatDate(new Date,"JST","yyyy年M月d日");
        ss1.getRange('B' + col).setValue(val);
        ss.getRange('AF' + kRow).setValue(val);
        //新しい注文書が発行されたので古い契約書PDFがあればIDを削除しておく
        ss1.getRange('AP' + r).setValue("");//大規模
      }else if(hako == "追加"){
         const col = 2 + (num * 3);
        const val = "発行" + Utilities.formatDate(new Date,"JST","yyyy年M月d日");
        ss1.getRange('B' + col).setValue(val);
        ss.getRange('AG' + kRow).setValue(val);
        ss1.getRange('AP' + r).setValue("");//大規模
      } 
      //◆担当者出来高SSの処理をする
      //担当者出来高SSの修正
      const kingakuF = () => {
        const num = ss .getRange('U' + kRow).getValue();
        const r1 = (num * 3) + 1;
        const r2 = r1 + 1;
        //本工事
        const rn1 = [ss1.getRange('G' + r1).getValue(),ss1.getRange('I' + r1).getValue(),ss1.getRange('L' + r1).getValue(),];
        const rn2 = [ss1.getRange('G' + r2).getValue(),ss1.getRange('I' + r2).getValue(),ss1.getRange('L' + r2).getValue(),];
        const rn = [rn1,rn2];
        return rn
      }
      const kingaku = kingakuF();
      tandeki(gen,kosyu,tMail,gMail,hakoB,gyomei,tan,kingaku);


      const pdfUrl = pdf.getUrl();
      const chat = chatUrl("管理2")//アプリメッセージ(maruyama)
      const msg = "注文請書の署名処理メール送信しました\n現場名：" + gen + "\n業者名：" + gyo + "\n業者メール：" + gMail + "\n担当メール：" + tMail + "\nCCメール：" + cc + "\n\n" + pdfUrl;
      sendChat(chat,msg);
    }

  // }catch(er){
  //   const chat = "https://chat.googleapis.com/v1/spaces/AAAA1s3kXVw/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=s3-qtxg5p1mS_Rc2wKl5EgBeh87KUgwK0fFDykLPpq0"//アプリメッセージ(maruyama)
  //   const msg = "◆◇ERROR案内注文書失敗⇒" + anTyu + "\n" + hako + "\n" + er.message;
  //   sendChat(chat,msg);
  // }
}
// --- helpers ---
function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
function nl2br_(s) {
  return String(s).replace(/\r\n|\n|\r/g, "<br>");
}

//◆◆担当者出来高SSへの反映
function tandeki(gen,kosyu,tanMail,gyoMail,hakoB,gName,tan,kingaku){
  //条件で分ける　①25日から15日（担当者出来高SSの期間）担出来SSの修正と担当者に更新を知らせる、請求書がすでに登録されてる場合は破棄された旨の連絡を担当者（再登録の必要性）および業者に連絡する
  //条件で分ける　②15日から20日（請求締切から支払い案内までの期間）担出来SSの修正と担当者に更新を知らせる、請求書が登録された場合は破棄された旨の連絡を担当者および業者に連絡する※請求書の再発行は担当者では行えない（必要な場合はシステムに依頼）
  //条件で分ける　①20日から25日（支払い案内発行から担出来SS発行までの期間）何もしない
  Logger.log("gen " + gen + ",kosyu " + kosyu + ",tanMail " + tanMail + ",gyoMail " + gyoMail + ",hakoB " + hakoB + ",gName " + gName + ",tan " + tan + ",kingaku " + kingaku);
  const ssT = 担当者出来高SS();//現在進行の担出来SSを取得する

  //担当者に出来高処理のリンクをつけてチャットに送信する
  const tanChat = () => {
    const body = "新規注文書が発行されました\n登録された請求書はリセットされました\n" + "現場名：" + gen + "\n工種：" + kosyu + "\n会社名：" + gName + "\n\n請求書プロセスが完了次第再度登録が可能になります\n\n（" + tan + "さんに送信）"
    //sendChat(url,body);//チャット送信
    //アプリメッセージにも送る
    const chat2 = chatUrl("管理2");//アプリメッセージ
    sendChat(chat2,body);
  }
  //keiyaku-seikyuと業者に契約書プロセスが完了したことを知らせるメール
  const gyoMailSend = () => {//契約書の送付
    //◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ
    //const keiyakuMail = "test@anemoworks.com";
    const keiyakuMail = メアド("請求支払");
    //◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ◆◇メンテ
    const mailSyomei = メール署名();
    const sub = "※ご請求についてのお知らせ※" + gen + "【" + kosyu + "】の" + gName + "様";
    const body = gName + "様\n\nいつもお世話になっております\n以下の現場について新しい契約書が発行されたため登録の請求書がリセットされました\n新しい契約書のご署名が完了後、再度処理を開始しますのでご対応よろしくお願いいたします（※請求受付は毎月15日までです）\n現場名：" + gen + "\n工種：" + kosyu + "\n\nこちらは自動送信されました不明な点は下記へお問い合わせください\n" + mailSyomei;
    const ccMail = tanMail + "," + keiyakuMail;
    const option = {
      cc:ccMail,
    }
    GmailApp.sendEmail(gyoMail,sub,body,option);
  };
 
  const day = new Date();
  const niti = day.getDate();
  const patternF = () => {
    if(niti >= 25 || niti <= 15){
      return 1;
    }else if(niti > 15 && niti < 20){
      return 2;
    }else if(niti >= 20 && niti < 25 ){
      return 3;
    }
    return null; // 万一の保険
  }
  const pattern = patternF();
  const last = ssT.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const end = last + 1;
  if(pattern == 1){
    const genKousyu = gen + kosyu;
    for(let i=2; i<=end; i++){
      const gk = ssT.getRange('B' + i).getValue() + ssT.getRange('D' + i).getValue();
      if(i == end){
        const msg = "契約書発行\n現現：" + gen + "\n工種：" + kosyu + "\n業者：" + gName + "\nは現在の担当者出来高登録にはなかったのです\ntandeki()は処理してません";
        const chat3 = chatUrl("管理2");//アプリメッセージ
        sendChat(chat3,msg);
      }else if(genKousyu == gk){
        hakobi = "発行" + Utilities.formatDate(new Date,"JST","yyyy年M月d日");
        if(hakoB == "両方"){
          //本工事
          ssT.getRange('H' + i).setValue(hakobi);//発行日
          //追加工事
          const r2 = i + 1;
          ssT.getRange('H' + r2).setValue(hakobi);//発行日
          ssT.getRange('E' + r2 + ":G" + r2).setValues([kingaku[1]]);//追加
          //前回までの出来高も再計算
          const deki1 = Math.round(kingaku[0][1] / kingaku[0][0] * 100);
          const deki2 = Math.round(kingaku[1][1] / kingaku[1][0] * 100);
          ssT.getRange('I' + i).setValue(deki1);//前回まで出来高
          ssT.getRange('I' + r2).setValue(deki2);//前回まで出来高
          ssT.getRange('J' + i).setValue(deki1);//今回出来高リセット
          ssT.getRange('J' + r2).setValue(deki2);//今回出来高リセット
          //既に請求登録がある場合のリセット
          const sei1 = ssT.getRange('M' + i).getValue();
          const sei2 = ssT.getRange('M' + r2).getValue();
          if(sei1 !== 0 || sei2 !== 0){//請求書が登録されてる場合
            if(sei1 !== 0){
              ssT.getRange('K' + i).setValue(kingaku[0][1]);//出来高は前回までに戻す
              ssT.getRange('L' + i).setValue(0);
              ssT.getRange('M' + i).setValue(0);
              ssT.getRange('N' + i).setValue("");
              ssT.getRange('O' + i).setValue("");
              ssT.getRange('P' + i).setValue("未発行");
            };
            if(sei2 !== 0){
              ssT.getRange('K' + r2).setValue(kingaku[1][1]);//出来高は前回までに戻す
              ssT.getRange('L' + r2).setValue(0);
              ssT.getRange('M' + r2).setValue(0);
              ssT.getRange('N' + r2).setValue("");
              ssT.getRange('O' + r2).setValue("");
              ssT.getRange('P' + r2).setValue("未発行");
            };
            tanChat();
            gyoMailSend()//登録が解除されましたのメール
          }
        }else if(hakoB == "契約" || hakoB == "変更"){
           //本工事
          ssT.getRange('H' + i).setValue(hakobi);//発行日
          ssT.getRange('E' + i + ":G" + i).setValues([kingaku[0]]);//本工事
          //前回までの出来高も再計算
          const deki1 = Math.round(kingaku[0][1] / kingaku[0][0] * 100);
          ssT.getRange('I' + i).setValue(deki1);//前回まで出来高
          ssT.getRange('J' + i).setValue(deki1);//今回出来高リセット
          //既に請求登録がある場合のリセット
          const sei1 = ssT.getRange('M' + i).getValue();
          if(sei1 !== 0){
            ssT.getRange('K' + i).setValue(kingaku[0][1]);//出来高は前回までに戻す
            ssT.getRange('L' + i).setValue(0);
            ssT.getRange('M' + i).setValue(0);
            ssT.getRange('N' + i).setValue("");
            ssT.getRange('O' + i).setValue("");
            ssT.getRange('P' + i).setValue("未発行");
            tanChat();
            gyoMailSend()//登録が解除されましたのメール
          };  
        }else if(hakoB == "追加"){
          const r2 = i + 1;
          ssT.getRange('H' + r2).setValue(hakobi);//発行日
          ssT.getRange('E' + r2 + ":G" + r2).setValues([kingaku[1]]);//追加
          //前回までの出来高も再計算
          const deki2 = Math.round(kingaku[1][1] / kingaku[1][0] * 100);
          ssT.getRange('I' + r2).setValue(deki2);//前回まで出来高
          ssT.getRange('J' + r2).setValue(deki2);//今回出来高リセット
          //既に請求登録がある場合のリセット
          const sei2 = ssT.getRange('M' + i).getValue();
          if(sei2 !== 0){
            ssT.getRange('K' + r2).setValue(kingaku[1][1]);//出来高は前回までに戻す
            ssT.getRange('L' + r2).setValue(0);
            ssT.getRange('M' + r2).setValue(0);
            ssT.getRange('N' + r2).setValue("");
            ssT.getRange('O' + r2).setValue("");
            ssT.getRange('P' + r2).setValue("未発行");
            tanChat();
            gyoMailSend()//登録が解除されましたのメール
          };
        }
        break;
      }
    }
  }else if(pattern == 2){
   
  }else if(pattern == 3){
    
  }
}

//◆◆注文請書スプレットシート作成
function 注文請書SS作成(hakoA,kosyuRow,day,fId,hakoB){
  //hakoA(変更、追加),kosyuRow(),day(発行日),fId(ファイルId),hakoB(両方の場合)
  // hako = "追加";
  // anTyuk = "案内";
  // kosyuRow = "11";
  // day = "2025年8月1日";
  // fId = "1854A1jYyu-imdtMWYsjNcQEiK4U3wygwFqKyCSJjtzQ";
  //発注書（単独の場合）kosyuRow = "単独," + gyo + "," + gyoRow + "," + nin + "," + kosyu + "," + kin;//カンマ繋ぎで送る
  const ss = 内訳鏡(fId);
  const ss1 = 予算書シート(fId);
  const kRow = Number(kosyuRow);
  Logger.log("kRow :" + kRow);
  const gyo = ss.getRange('V' + kRow).getValue();
  const kosyu = ss.getRange('T' + kRow).getValue();
  const gen = ss1.getRange('H102').getValue();
  const gRow = ss.getRange('AE' + kRow).getValue();
  const gInfo = 業者情報(gRow);//[id,株式,（株）,〒,住所,tel,メール]  

  const hId = ひな形("業者注文請書");
  const sssH = DriveApp.getFileById(hId);//注文請書ひな形シート
  //一時フォルダにコピー
  const flId = ドライブId("3ヶ月フォルダ");
  const sFld = DriveApp.getFolderById(flId);//一時フォルダー(3か月)　【管理】請求支払の【契約署名】開発フォルダ内
  const hName = gen + "_" + gyo + "【ひな形】注文請書" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");
  const hIdA = sssH.makeCopy(hName,sFld).getId();
  ss.getRange('AN' + kRow).setValue(hIdA);//idは署名後の作成にも使用するので記録
  const sss = SpreadsheetApp.openById(hIdA);
  const ssH1 = sss.getSheetByName('注文請書発行（石井⇔業者)ひな型');
  const ssH2 = sss.getSheetByName('業者案内用内訳');
  const ssH3 = sss.getSheetByName('業者案内用内訳 （追加）');
  const ssH4 = sss.getSheetByName('変更履歴');

  //表紙反映
  const gName = gInfo[1];//業者名（株式会社)
  const pCode = gInfo[3];//〒
  const gAdd = gInfo[4];//業者住所
  const gTel = gInfo[5];//業者TEL
  const gMail = gInfo[6];//業者メール
  const syomeibi = "発行署名：" + Utilities.formatDate(new Date,"JST","yyyy年M月d日 hh:mm:ss");

  const kId = ss1.getRange('A102').getValue();//工事No
  const kinds = ss1.getRange('I102').getValue();//工種
  const genmei = gen + " " +  kinds;//件名
  const add = ss1.getRange('T102').getValue();//現場住所
  const st = ss1.getRange('R102').getValue();//着工
  const fh = ss1.getRange('S102').getValue();//完了

  const hakoF = () => {
    if(hakoA == "追加"){
      const tuiHako = ss1.getRange('AG' + kRow).getValue();
      if(tuiHako == ""){//追加は新規
        return hakoA;
      }else{
        return "追変";//追加の変更
      }
    }else{
      return hakoA;
    }
  }
  const hako = hakoF();

  const judgeKF = () => {//契約の状況を返す（状態は、本（本契約、追加なし）、本追（本契約（変更ない）、追加あり）、変（変更契約、追加なし）、変追（変更契約、追加あり）
    const hen = ss.getRange('AB' + kRow).getValue();
    const tui = ss.getRange('AD' + kRow).getValue();
    Logger.log('hen ：' + hen + ', tui ：' + tui);
    if(hen == 0 && tui == 0){
      return "本";//内訳は本契約のみ（初回）
    }else if(hen !== 0 && tui == 0){
      return "変";//内訳は変更のみ
    }else if(hen == 0 && tui !== 0){
      return "本追";//内訳は契約と追加
    }else if(hen !== 0 && tui !== 0){
      return "変追";//内訳は変更と追加
    }
  };
  const judgeK = judgeKF();//値によって内訳書他変更する

  const hakobiF = () => {
    if(day == "着工日"){
      const result = st;
      return result;
    }else{
      return day;
    }
  }
  const hakobi = hakobiF();
  ssH1.getRange('M4').setValue(hakobi);//発行日
  ssH1.getRange('B5').setValue(gName);//業者名（株式）
  ssH1.getRange('B7').setValue("〒" + pCode);//〒
  ssH1.getRange('B8').setValue(gAdd);//業者住所
  ssH1.getRange('B9').setValue("Tel：" + gTel);//業者Tel
  ssH1.getRange('B10').setValue("mail：" + gMail);//業者メール
  ssH1.getRange('D13').setValue(kId);//工事No
  ssH1.getRange('D14').setValue(genmei);//件名
  ssH1.getRange('D15').setValue(add);//現場住所
  ssH1.getRange('D17').setValue(st);//着工  
  ssH1.getRange('F17').setValue(fh);//完了

  ssH1.getRange('D41').setValue(kosyu + " 契約工事");//工種

  ssH1.getRange('L17').setValue(syomeibi);//発行署名

  //業者署名のとこクリア
  ssH1.getRange('J63').setValue("");
  ssH1.getRange('J64').setValue("");
  ssH1.getRange('J64:N64').setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

  ssH1.getRange('J63').setValue("未署名（まだ契約は成立してません)").setFontColor("#FF8C00");


  const ukeF = () => {//追加がない場合はは契約または変更内訳のみ、追加がある場合は契約または変更内訳と追加内訳をつける
    if(judgeK == "本" || judgeK == "本追"){
      ssH1.getRange('O41').setValue("");
      ssH2.getRange('A1').setValue("内　訳　書　（契約）");
      const result = ss.getRange('X' + kRow).getValue();
      return result;
    }else if(judgeK == "変" || judgeK == "変追"){
      ssH1.getRange('O41').setValue("【変更】");
      ssH2.getRange('A1').setValue("内　訳　書　（変更）");
      ssH2.getRange('A1:C1').setBackground("#DDDDDD");
      const result = ss.getRange('Z' + kRow).getValue();
      return result;
    }
  }
  const uke = ukeF();

  ssH1.getRange('M41').setValue(uke);//本工事のセット

  if(judgeK == "本追" || judgeK == "変追"){//追加がある場合２行目に工種＋追加の項目を追記する
    const tui = ss.getRange('AD' + kRow).getValue();
    ssH1.getRange('D43').setValue(kosyu + " 追加工事");
    ssH1.getRange('H43').setValue(1);
    ssH1.getRange('J43').setValue("式");

    if(hako == "追変"){//追加が変更の場合は
      ssH1.getRange('O43').setValue("【変更】");
      if(hakoB == "両方"){
        ssH1.getRange('O41').setValue("【変更】");
      }
    }
    ssH1.getRange('M43').setValue(tui);
  }

  //内訳書の反映
  ssH2.getRange('B2').setValue(genmei);//現場名
  ssH2.getRange('B3').setValue(add);//現場住所
  ssH2.getRange('D1').setValue(gName);//業者名
  ssH2.getRange('I1').setValue(st);//着工
  ssH2.getRange('I2').setValue(fh);//完了
  ssH2.getRange('E2').setValue(kosyu);//工種
  //追加内訳
  ssH3.getRange('B2').setValue(genmei);//現場名
  ssH3.getRange('B3').setValue(add);//現場住所
  ssH3.getRange('D1').setValue(gName);//業者名
  ssH3.getRange('I1').setValue(st);//着工
  ssH3.getRange('I2').setValue(fh);//完了
  ssH3.getRange('E2').setValue(kosyu);//工種
  const tan = ss1.getRange('U102').getValue();//担当者
  const tel = ss1.getRange('Y102').getValue();//担当者Tel
  const tMail = ss1.getRange('Z102').getValue();//担当者mail
  ssH2.getRange('E3').setValue(tan);
  ssH2.getRange('H3').setValue(tel);
  ssH2.getRange('I3').setValue(tMail);
  //追加内訳
  ssH3.getRange('E3').setValue(tan);
  ssH3.getRange('H3').setValue(tel);
  ssH3.getRange('I3').setValue(tMail);

  const rF = (r) => {//見出し項目確認
    const nums = [41,81,121,161,201,241,281];
    const len = nums.length;
    for(let i=0; i<=len; i++){
      if(i == len){
        return r;
      }else if(r == nums[i]){
        const res = r + 1;
        return res;
      }
    }
  };
  const colF = (s,uti) => {
    if(uti == "契約"){
      if(s == 6){//数量
        return 7;
      }else if(s == 7){//単価
        return 12;
      }else if(s == 8){//金額
        return 13;
      }else if(s == 9){//備考
        return 10;
      }
    }else if(uti == "変更"){
      if(s == 6){//数量
        return 18;
      }else if(s == 7){//単価
        return 12;
      }else if(s == 8){//金額
        return 20;
      }else if(s == 9){//備考
        return 17;
      }
    }else if(uti == "追加"){
      if(s == 6){//数量
        return 7;
      }else if(s == 7){//単価
        return 12;
      }else if(s == 8){//金額
        return 13;
      }else if(s == 9){//備考
        return 17;
      }
    }
  }

  //◆内訳書作成関数
  const utiwake = (data,k,uti) => {//内訳書作成
    const length = data.length;
    const sheetF = () => {//内訳書を選択
      if(uti == "追加"){
        return ssH3;
      }else{
        return ssH2;
      }
    }
    const sheet = sheetF();
    
    for(let i=0; i<=length; i++){//lengthは1始まり
      if(i == length){
        //不要行の削除
        const rw = sheet.getRange('B' + k).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();//反映ページ最終行
        const rw2 = rw - 1;//残す最終行
        const maxRow = sheet.getMaxRows(); // 現在のシートの最大行数を取得
        const rw3 = maxRow - rw2;        // 残すべき最終行より下の行をすべて計算
        if (rw3 > 0) {
          sheet.deleteRows(rw, rw3);
        }
        //数式の代入 数量×単価
        const mas1 = '=IFERROR(IF(SUM(F6*G6)=0,"",SUM(F6*G6)),"金額")';
        sheet.getRange('H6').setFormula(mas1);
        let rn = sheet.getRange('H6');
        let rn2 = sheet.getRange('H7:H' + rw2);
        rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);//数式
        //合計の数式
        const mas2 = 'ROUNDDOWN(SUM(H6:H' + rw2 + '),-3)';
        const mas3 = 'ROUNDDOWN(SUM(H6:H' + rw2 + '),-3)-SUM(H6:H' + rw2 + ')';
        sheet.getRange('D4').setFormula(mas2);
        sheet.getRange('C4').setFormula(mas3);
        SpreadsheetApp.flush();
        Logger.log(k);
        return k;
      }else{
        k = k + 1;
        k = rF(k);
        for(let s=2; s<=9; s++){
          if(s == 2){
            sheet.getRange(k,s).setValue(data[i][0]);
          }else if(s == 3 ||s == 4 || s == 5){
            const col = s + 1;
            sheet.getRange(k,s).setValue(data[i][col]); 
          }else{
            const col = colF(s,uti);
            sheet.getRange(k,s).setValue(data[i][col]); 
          }
        }
      }
    }
  }
  const rirekiF = () => {
    const time = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
    let reki = ss.getRange('AK' + kRow).getValue();
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
    ss.getRange('AK' + kRow).setValue(reki);

    // 変更履歴シートにも貼り付け
    const row = data.length + 7;
    ssH4.getRange('B8:G' + row).setValues(data);
    ssH4.getRange('C1').setValue(genmei + " 【" + kosyu + "】");
    ssH4.getRange('C4').setValue(gName);
    ssH4.getRange('D4').setValue(gMail);

    const last = ssH4.getRange('B17').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    for (let i = 13; i <= last; i++) {
      if (ssH4.getRange('E' + i).getValue() == "プロセス完了") {
        ssH4.getRange('E' + i).setFontColor("#1e90ff").setFontWeight("bold"); // ←ここ直しました
      }
    }

    return "OK";
  }

  
  //◆judgeK（本、本追、変、変追）によって内訳作成をする
  Logger.log('judgeK ：' + judgeK)
  if (judgeK == "本" || judgeK == "本追") {
    const data1 = anTyuKakunin(kosyu, fId);
    let k = 5;
    let uti = "契約";
    k = utiwake(data1, k, uti);  // ①最初の内訳

    if (hako == "追加" || hako == "追変") {
      const kosyuT = "追加" + kosyu;
      const data2 = anTyuKakunin(kosyuT, fId);
      uti = "追加";
      k = 5;
      k = utiwake(data2, k, uti);  // ②追加内訳
    }

    const rireki = rirekiF();  // ③履歴処理
    if (rireki == "OK") {
      return hIdA;
    } else {
      return rireki;
    }

  } else if (judgeK == "変" || judgeK == "変追") {
    const data1 = anTyuKakunin(kosyu, fId);
    let k = 5;
    let uti = "変更";
    k = utiwake(data1, k, uti);  // ①変更内訳

    if (judgeK == "変追") {
      const kosyuT = "追加" + kosyu;
      const data2 = anTyuKakunin(kosyuT, fId);
      uti = "追加";
      k = 5;
      k = utiwake(data2, k, uti);  // ②追加内訳
    }

    const rireki = rirekiF();  // ③履歴処理
    if (rireki == "OK") {
      return hIdA;
    } else {
      return rireki;
    }
  }
}

//◆◆PDF作成（発注、案内)
function PDF作成(id,hako,anTyu,kosyuRow,fId){
  Logger.log("id:" + id + ",hako:" + hako + ",anTyu:" + anTyu + ",kosyuRow:" + kosyuRow + ",fId:" + fId);
  const ss = 内訳鏡(fId);
  const kRow = Number(kosyuRow);
  const kosyu = ss.getRange('T' + kRow).getValue();
  const sheets = SpreadsheetApp.openById(id);  
  const ssH1 = sheets.getSheetByName('注文請書発行（石井⇔業者)ひな型');
  const ssH2 = sheets.getSheetByName('業者案内用内訳');
  const ssH3 = sheets.getSheetByName('業者案内用内訳 （追加）');
  const ssH4 = sheets.getSheetByName('変更履歴');

  const judgeKF = () => {//契約の状況を返す（状態は、本（本契約、追加なし）、本追（本契約（変更ない）、追加あり）、変（変更契約、追加なし）、変追（変更契約、追加あり）
    const hen = ss.getRange('AB' + kRow).getValue();
    const tui = ss.getRange('AD' + kRow).getValue();
    if(hen == 0 && tui == 0){
      return "本";//内訳は本契約のみ（初回）
    }else if(hen !== 0 && tui == 0){
      return "変";//内訳は変更のみ
    }else if(hen == 0 && tui !== 0){
      return "本追";//内訳は契約と追加
    }else if(hen !== 0 && tui !== 0){
      return "変追";//内訳は変更と追加
    }
  };
  const judgeK = judgeKF();//値によって内訳書他変更する

  const genmei = ssH1.getRange('D14').getValue();
  const gyomei = ssH1.getRange('B5').getValue();
  const ss1 = 予算書シート(fId);
  const gen = ss1.getRange('H102').getValue();
  const nendo = ss1.getRange('D102').getValue();
  const sFld = 保存フォルダ(nendo,gen,anTyu);
  const pNameF = () => {
    if(anTyu == "案内"){
      let fn = gyomei + "様★" + genmei + Utilities.formatDate(new Date,"JST","yyMMdd") + ".pdf";
      if(hako == "契約"){
        fn = "【案内】" + fn;
        return fn;
      }else if(hako == "変更"){
        fn = "【案内（変更）】" + fn;
        return fn;
      }else if(hako == "追加"){
        fn = "【案内（追加）】" + fn;
        return fn;
      }
    }else if(anTyu == "発注"){
      let fn = "【未署名】注文書&請書★" + genmei + "【" + kosyu + "】" + gyomei + "様" + Utilities.formatDate(new Date,"JST","yyMMdd") + ".pdf";
      return fn;
    }
  }
  const pName = pNameF();
  Logger.log(pName);
  if(anTyu == "案内"){
    if(hako == "追加"){
      sheets.deleteSheet(ssH1);//表紙を削除
      sheets.deleteSheet(ssH2);//契約内訳を削除
      sheets.deleteSheet(ssH4);//履歴
    }else{
      sheets.deleteSheet(ssH1);//表紙を削除
      sheets.deleteSheet(ssH3);//追加内訳を削除
      sheets.deleteSheet(ssH4);//履歴
    }
    const pdf = sheets.getAs('application/pdf');
    pdf.setName(pName);
    const pId = sFld.createFile(pdf).getId();
    const pfl = DriveApp.getFileById(pId);
    return pfl;//ファイルを返す
  }else if(anTyu == "発注"){
    if(judgeK == "本" || judgeK == "変"){
      sheets.deleteSheet(ssH3);//追加内訳を削除
      //◆◇◆◇◆◇◆◇◆◇◆◇準備が整ったら消す◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
      //sheets.deleteSheet(ssH4);//履歴まだ準備中なので
      //◆◇◆◇◆◇◆◇◆◇◆◇準備が整ったら消す◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
    }else if(judgeK == "本追" || judgeK == "変追"){
      //◆◇◆◇◆◇◆◇◆◇◆◇準備が整ったら消す◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
      //sheets.deleteSheet(ssH4);//履歴まだ準備中なので
      //◆◇◆◇◆◇◆◇◆◇◆◇準備が整ったら消す◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
    }
    const pdf = sheets.getAs('application/pdf');
    pdf.setName(pName);
    const pId = sFld.createFile(pdf).getId();
    const pfl = DriveApp.getFileById(pId);
    return pfl;//ファイルを返す
  }
}

//保存先フォルダを返す（注文請書ドライブの石井⇒業者フォルダ内の現場フォルダ)
function 保存フォルダ(nendo,gen,anTyu){
  // nendo = 2026;
  // gen = "テスト予算書物件";
  // anTyu = "案内";
  const drive = ドライブId("注文請書");
  const oya = DriveApp.getFolderById(drive)//注文請書ドライブ
  //年度フォルダ
  const nFldF = () => {
    const fName = nendo + "年";
    const nFls = oya.getFoldersByName(fName);
    if(nFls.hasNext()){
      const result = nFls.next();
      return result;
    }else{
      const result = oya.createFolder(fName);
      return result;
    }
  }
  const nFld = nFldF();
  //石井→業者フォルダ
  const hFldF = () => {
    const fName = "石井→業者";
    const fls = nFld.getFoldersByName(fName);
    if(fls.hasNext()){
      const result = fls.next();
      return result;
    }else{//ない場合は年度フォルダがなかった場合なので作成
      const result = nFld.createFolder(fName);
      nFld.createFolder('注文書大規模（ｽﾍﾟｰｽにもあげる）');
      nFld.createFolder('注文書小規模（ｽﾍﾟｰｽにもあげる）');
      return result;
    }
  }
  const hFld = hFldF();
  //現場フォルダ
  const gFldF = () => {
    const fName = gen;
    const fls = hFld.getFoldersByName(fName);
    if(fls.hasNext()){
      const result = fls.next();
      return result;
    }else{//ない場合は他フォルダも作成
      const result = hFld.createFolder(fName);
      return result;
    }
  }
  const gFld = gFldF();
  //現場フォルダ
  const sFldF = () => {
    const fNameF = () => {
      if(anTyu == "案内"){
        return "案内";
      }else if(anTyu == "発注"){
        return "注文書発行";
      }
    }
    const fName = fNameF();
    const fls = gFld.getFoldersByName(fName);
    if(fls.hasNext()){
      const result = fls.next();
      return result;
    }else{//ない場合は他フォルダも作成
      const result = gFld.createFolder(fName);
      gFld.createFolder("業者署名済");
      if(fName == "案内"){
        gFld.createFolder("注文書発行");
      }else{
        gFld.createFolder("案内");
      }
      return result;
    }
  }
  const sFld = sFldF();
  Logger.log(sFld.getId());
  return sFld;
}

//◆◆案内用Excelファイル生成
function Excel生成(id,fn,nendo,gen,anTyu){
  // 変換したいスプレッドシートのID（URLの /d/ と /edit の間の文字列）
  // Excel形式（.xlsx）でエクスポートするURL
  const url = "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx";

  // 認証トークンの取得
  const token = ScriptApp.getOAuthToken();
  
  // URLFetchAppでファイルを取得
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  });
  
  // Blobに変換し、ファイル名をつけてDriveに作成
  const blob = response.getBlob().setName(fn);
  
  // DriveApp でファイルを作成
  const sFld = 保存フォルダ(nendo,gen,anTyu);
  const file = sFld.createFile(blob);
  return file
}

//◆◆発注書URLの取得
function hatyuGetUrl(fId,ds,gyo,kRow,kinds){//kindsがhaTの場合は発注単独、haDの場合は通常の注文書
  const row = Number(kRow);
  Logger.log(fId + "," + ds + "," + gyo + "," + kRow);
  const ss = 予算書シート(fId);
  const idF = () => {
    if(ds == "D"){
      if(kinds == "haT"){
        const sheets = SpreadsheetApp.openById(fId);
        const sheet = sheets.getSheetByName("色管理");
        return sheet.getRange('U' + row).getValue();
      }else{
        const ssK = 内訳鏡(fId);
        const num = ssK.getRange('U' + row).getValue();
        const r = (num * 3) + 1;
        return ss.getRange('AP' + r).getValue();
      }
    }else if(ds == "S"){
      return ss.getRange('AJ' + row).getValue();
    }
  }  
  const id = idF();
  Logger.log(id);
  if(id == ""){//まだ登録がない場合
    return "なし";
  }else{
    const file = DriveApp.getFileById(id);
    const url = file.getUrl();
    return url;
  }
}

//🔹🔹発注書発行（単独）
function tanHRezist(json){//nin(一式、人工)
  const data = JSON.parse(json);
  const gen = data[0];
  const nin = data[1];
  const kosyu = data[2];
  const gyo = data[3];
  const kin = Number(data[4].replace(/,/g,""));
  const message = data[5];
  const fId = data[6];
  const day = data[7];
  const gyoRow = Number(data[8]);
  
  Logger.log("gen " + gen + ",nin " + nin + ",kosyu " + kosyu + ",gyo " + gyo + ",kin " + kin + ",message " + message + ",fId " + fId + ",day " + day + ",gyoRow" + gyoRow);
  const hako = "契約";
  const kosyuRow = "単独," + gyo + "," + gyoRow + "," + nin + "," + kosyu + "," + kin;//カンマ繋ぎで送る
  const hakoB = "契約";
  const ss = 内訳鏡(fId);
  const ss1 = 予算書シート(fId);
  const sheets = SpreadsheetApp.openById(fId);
  const ssA = sheets.getSheetByName("色管理");
  const kRowF = () => {
    let r = ssA.getRange('V100').getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
    if(r < 3){
      r = 3;
    }
    return r; 
  }
  const kRow = kRowF();
  Logger.log("kRow :" + kRow);

  const id = 注文単独SS作成(hako,kosyuRow,day,fId,hakoB,kRow);//hakoA(契約、変更、追加),kosyuRow(),day(発行日),fId(ファイルId),hakoB(両方の場合)

  Logger.log("注文単独SS作成 " + id);
  const pdf = 単独PDF作成(id,kosyuRow,fId); //DriveAppオブジェクト
  
  //const gyoRow = ss.getRange('AE' + kRow).getValue();
  //const kosyu = ss.getRange('T' + kRow).getValue();
  const gInfo = 業者情報(gyoRow)//[id,株式,（株）,〒,住所,tel,メール]  
  const gyomei = gInfo[1];
  //const gyo = gInfo[2];
  const gMail = gInfo[6];
  Logger.log(gyomei + "," + gyo + "," + gMail);
  //const gen = ss1.getRange('H102').getValue();
  const kinds = ss1.getRange('I102').getValue();
  const genmei = gen + " " + kinds;
  const tan = ss1.getRange('U102').getValue();
  const nendo = ss1.getRange('D102').getValue();
  const tMail = ss1.getRange('Z102').getValue();
  const add = ss1.getRange('T102').getValue();
  const st = ss1.getRange('R102').getValue();
  const fh = ss1.getRange('S102').getValue();
  const tTel = ss1.getRange('Y102').getValue();
  const koki = Utilities.formatDate(st,"JST","yyyy年M月d日") + "～" + Utilities.formatDate(fh,"JST","yyyy年M月d日");

  const ccF = () => {
    let ccM = ss.getRange('X1').getValue();
    if(ccM !== ""){
      return tMail + "," + ccM;
    }else{
      return tMail;
    }
  } 
  const cc = ccF();
  
  //hakoBの記録とパスワード記録
  ssA.getRange('AL' + kRow).setValue(hakoB);
  const ramdam = Math.floor(100000 + Math.random() * 900000); // 100000〜999999
  const code = ramdam.toString();
  ssA.getRange('AM' + kRow).setValue(code);
  const subject = "【注文書処理ください】" + genmei + "（" + kosyu + "）" + gyo + "様";
  const linkPdf = pdfLink("電子契約やり方");//◆やり方説明PDFリンク
  const kRowA = "単独" + kRow;//注文書単独URL貼り付け用
  const pdfId = pdf.getId();
  const blob = pdf.getBlob();
  const keiyakuLinkF = () => {
    const urlA = getSetUrl("契約書署名");
    const url = urlA + "?param1=" + fId + "&param2=" + gyoRow + "&param3=" + kRowA + "&param4=" + pdfId + "&param5=" + code;
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
    <p>${message}</p>
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

  const val = "発行" + Utilities.formatDate(new Date,"JST","yyyy年M月d日");
  ssA.getRange('AF' + kRow).setValue(val);

  const pdfUrl = pdf.getUrl();
  const chat = chatUrl("管理2")//アプリメッセージ(maruyama)
  const msg = "注文請書の署名処理メール送信しました\n現場名：" + gen + "\n業者名：" + gyo + "\n業者メール：" + gMail + "\n担当メール：" + tMail + "\nCCメール：" + cc + "\n\n" + pdfUrl;
  sendChat(chat,msg);
}

//🔹🔹発注書単独用注文請書スプレットシート作成
function 注文単独SS作成(hako,kosyuRow,day,fId,hakoB,kRow){
  //hako,hakoAは"契約"
  //発注書（単独の場合）kosyuRow = "単独," + gyo + "," + gyoRow + "," + nin + "," + kosyu + "," + kin;//カンマ繋ぎで送る
  const rn = kosyuRow.split(',');
  const sheets = SpreadsheetApp.openById(fId);
  const ss = sheets.getSheetByName("色管理");
  const ss1 = 予算書シート(fId);
  
  const gyo = rn[1];
  const gRow = Number(rn[2]);
  const nin = rn[3];
  const kosyu = rn[4];
  const kin = Number(rn[5]);
  const gen = ss1.getRange('H102').getValue();
  
  const gInfo = 業者情報(gRow);//[id,株式,（株）,〒,住所,tel,メール]  
  const hId = ひな形("注文請書単独");
  const sssH = DriveApp.getFileById(hId);//注文請書ひな形シート（単独用）
  //一時フォルダにコピー
  const flId = ドライブId("3ヶ月フォルダ");
  const sFld = DriveApp.getFolderById(flId);//一時フォルダー(3か月)　【管理】請求支払の【契約署名】開発フォルダ内
  const hName = gen + "_" + gyo + "【ひな形】注文請書" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");
  const hId2 = sssH.makeCopy(hName,sFld).getId();
  ss.getRange('AN' + kRow).setValue(hId2);//idは署名後の作成にも使用するので記録
  ss.getRange('X' + kRow).setValue(kin);
  ss.getRange('V' + kRow).setValue(gyo);
  ss.getRange('AE' + kRow).setValue(gRow);
  ss.getRange('T' + kRow).setValue(kosyu); 
  ss.getRange('AL' + kRow).setValue(hako); 
  
  const sss = SpreadsheetApp.openById(hId2);
  const snF = () => {
    if(nin == "一式"){
      return "業者注文書（単発一式)";
    }else if(nin == "人工"){
      return "業者注文書（人工契約）";
    }
  }
  const sn = snF();
  const ssH1 = sss.getSheetByName(sn);
  // const ssH2 = sss.getSheetByName('業者案内用内訳');
  // const ssH3 = sss.getSheetByName('業者案内用内訳 （追加）');
  const ssH4 = sss.getSheetByName('変更履歴');

  //表紙反映
  const gName = gInfo[1];//業者名（株式会社)
  const pCode = gInfo[3];//〒
  const gAdd = gInfo[4];//業者住所
  const gTel = gInfo[5];//業者TEL
  const gMail = gInfo[6];//業者メール
  const syomeibi = "発行署名：" + Utilities.formatDate(new Date,"JST","yyyy年M月d日 hh:mm:ss");

  const kId = ss1.getRange('A102').getValue();//工事No
  const kinds = ss1.getRange('I102').getValue();//工種
  const genmei = gen + " " +  kinds;//件名
  const add = ss1.getRange('T102').getValue();//現場住所
  const st = ss1.getRange('R102').getValue();//着工
  const fh = ss1.getRange('S102').getValue();//完了
  
  const hakobiF = () => {
    if(day == "着工日"){
      const result = st;
      return result;
    }else{
      return day;
    }
  }
  const hakobi = hakobiF();
  ssH1.getRange('M4').setValue(hakobi);//発行日
  ssH1.getRange('B5').setValue(gName);//業者名（株式）
  ssH1.getRange('B7').setValue("〒" + pCode);//〒
  ssH1.getRange('B8').setValue(gAdd);//業者住所
  ssH1.getRange('B9').setValue("Tel：" + gTel);//業者Tel
  ssH1.getRange('B10').setValue("mail：" + gMail);//業者メール
  ssH1.getRange('D13').setValue(kId);//工事No
  ssH1.getRange('D14').setValue(genmei);//件名
  ssH1.getRange('D15').setValue(add);//現場住所
  ssH1.getRange('D17').setValue(st);//着工  
  ssH1.getRange('F17').setValue(fh);//完了

  ssH1.getRange('D41').setValue(kosyu + " 契約工事");//工種

  ssH1.getRange('L17').setValue(syomeibi);//発行署名

  //業者署名のとこクリア
  ssH1.getRange('J63').setValue("");
  ssH1.getRange('J64').setValue("");
  ssH1.getRange('J64:N64').setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

  ssH1.getRange('J63').setValue("未署名（まだ契約は成立してません)").setFontColor("#FF8C00");

  ssH1.getRange('M41').setValue(kin);//請負金額

  const time = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
  const reki = `${hako},${time},,発行,,`;
  ss.getRange('AK' + kRow).setValue(reki);

  const data = [reki.split(',')];  // 1行分の配列
  // 変更履歴シートにも貼り付け
  ssH4.getRange('B8:G8').setValues(data);
  ssH4.getRange('C1').setValue(genmei + " 【" + kosyu + "】");
  ssH4.getRange('C4').setValue(gName);
  ssH4.getRange('D4').setValue(gMail);

  return hId2;
}
//🔹🔹発注書単独PDF作成
function 単独PDF作成(id,kosyuRow,fId){
  const rn = kosyuRow.split(',');//kosyuRow = "単独," + gyo + "," + gyoRow + "," + nin + "," + kosyu + "," + kin;//カンマ繋ぎで送る
  const kosyu = rn[4];
  const nin = rn[3];
  const sheets = SpreadsheetApp.openById(id);  
  const ssH1 = sheets.getSheetByName('業者注文書（人工契約）');
  const ssH2 = sheets.getSheetByName('業者注文書（単発一式)');

  const genmeiF = () => {
    if(nin == "人工"){
      return ssH1.getRange('D14').getValue();
    }else if(nin == "一式"){
      return ssH2.getRange('D14').getValue();
    }
  }
  const genmei = genmeiF();
  const gyomeiF = () => {
    if(nin == "人工"){
      return ssH1.getRange('B5').getValue();
    }else if(nin == "一式"){
      return ssH2.getRange('B5').getValue();
    }
  }
  const gyomei = gyomeiF();
  const ss1 = 予算書シート(fId);
  const gen = ss1.getRange('H102').getValue();
  const nendo = ss1.getRange('D102').getValue();
  const sFld = 保存フォルダ(nendo,gen,"発注");

  const pName = "【未署名】注文書&請書★" + genmei + "【" + kosyu + "】" + gyomei + "様" + Utilities.formatDate(new Date,"JST","yyMMdd") + ".pdf";;
  Logger.log(pName);
  
  if(nin == "人工"){
    sheets.deleteSheet(ssH2);//シートを削除
  }else if(nin == "一式"){
    sheets.deleteSheet(ssH1);//シートを削除
  }
  const pdf = sheets.getAs('application/pdf');
  pdf.setName(pName);
  const pId = sFld.createFile(pdf).getId();
  const pfl = DriveApp.getFileById(pId);
  return pfl;//ファイルを返す
}