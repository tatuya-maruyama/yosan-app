//◆◆担当者メール検索
function getMotoMail(taiHase,name){
  const id = getSPId("長谷工完了報告");
  const sss = SpreadsheetApp.openById(id);//【管理】請求支払い→【管理表長谷工】工事完了報告書ファイル
  const ss = sss.getSheetByName('メールアドレス');
  if(taiHase == "長谷工"){
    const last = ss.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    const end = last + 1;
    for(let i=2; i<end+1; i++){
      if(i == end){
        return "なし";
      }else if(ss.getRange('A' + i).getValue() == name){  
        const mail = ss.getRange('C' + i).getValue();
        return mail;
      }
    }
  }else if(taiHase == "大成"){
    const last = ss.getRange('G1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    const end = last + 1;
    for(let i=2; i<end+1; i++){
      if(i == end){
        return "なし";
      }else if(ss.getRange('G' + i).getValue() == name){  
        const mail = ss.getRange('H' + i).getValue();
        return mail;
      }
    }
  }
}
//◆◆担当者メール登録なかったので登録
function reziMotoMail(name,mail,taiHase){
  const id = getSPId("長谷工完了報告");
  const sss = SpreadsheetApp.openById(id);
  const ss = sss.getSheetByName('メールアドレス');
  if(taiHase == "長谷工"){
    const row = ss.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
    ss.getRange('A' + row).setValue(name);
    ss.getRange('C' + row).setValue(mail);
  }else if(taiHase == "大成"){
    const row = ss.getRange('G1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
    ss.getRange('G' + row).setValue(name);
    ss.getRange('H' + row).setValue(mail);
  }
}
//◆😃◆😃◆😃◆😃現場新規登録◆😃◆😃◆😃◆😃◆
function newGenRezi(json){
  const data = JSON.parse(json);
  const ds = data.rezi[2];
  const genA = data.rezi[3];
  const kinds = data.rezi[4];
  const uke = data.rezi[5];
  const start = data.rezi[13];
  const finish = data.rezi[14];
  const add = data.rezi[15];
  const tan = data.rezi[16];
  const client = data.rezi[17];
  const motoTan = data.rezi[20];
  const motoMail = data.rezi[21];
  const ekimu = data.rezi[22];
  const keiyaku = data.rezi[24]
  const nendo = data.rezi[27];
  const huri = data.rezi[28];
  const hatyuNo = data.rezi[23];
  const memo = data.rezi[19];
  const toroku = data.rezi[25];
  const spasceUrl = data.rezi[26];//新規チャットURL
  const kFileId = data.fId;
  const tyaku = data.tyaku;//日付で判断（着工ORnot)

  //◆◆担当者からメアド＆番号取得
  const tanInfo = Get社員情報(tan);//名前⇒[row,mail,tel,chat,出面chat];
  const mail = tanInfo[1];
  const tel = tanInfo[2];

  Logger.log(data);
  Logger.log("uke " + uke);
  Logger.log("🔸newGenRezi()のkFileId " + kFileId);
  
  //ファイル移動先フォルダ取得(フォルダIdと工事台帳Rowを返す)
  const saveFld = 保存フォルダと台帳(nendo,ds);
  if (!saveFld) {
    throw new Error("保存先フォルダが見つかりません（年度:" + nendo + "）");
  }
  const sFldId = saveFld[0];
  const sFld = DriveApp.getFolderById(sFldId);
  const kId = saveFld[1];
  const sssK = SpreadsheetApp.openById(kId);
  const ssK = sssK.getSheetByName('工事リスト');
  const rw = ssK.getRange('A2').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;

   //codeの作成
  const creatCode = 現場コード(nendo,rw,genA,ds);
  const genmei = creatCode[1];
  const code = creatCode[0];
  契約書移動(kFileId,nendo,genmei,ds);//契約書ファイルがある場合のファイルを移動する関数

  Logger.log("ds " + ds);
  if(ds == "D" || ds == "S2"){//◆◆大規模ファイル作成保存◆◆
    const name = (ds == "D")? genmei + "◆D支払管理" : genmei + "◆S支払管理";
    const idY = ひな形("予算書雛形");
    const hinaSp = SpreadsheetApp.openById(idY);//大規模予算書ひな形
    const newSpId = hinaSp.copy(name).getId();

    シート権限_再作成_2シートだけ(newSpId,mail)//権限を追加
    
    let file = DriveApp.getFileById(newSpId);
    file.moveTo(sFld);
    //編集開始
    const sheetYs = SpreadsheetApp.openById(newSpId);
    const sheetY = sheetYs.getSheetByName('シート1');//大規模予算書（作成）
    const ss2 = sheetYs.getSheetByName('内訳（鏡）');
    
    //コントロールパネルハイパーリンク
    const urlA = getSetUrl("予算書アプリ");
    const url = urlA + "?page=pane&param1=" + newSpId;
    const displayText = "コントロールパネル";
    ss2.getRange('R4').setFormula('=HYPERLINK("' + url + '", "' + displayText + '")');

    sheetY.getRange('Z102').setValue(mail);
    sheetY.getRange('Y102').setValue(tel);

    sheetY.getRange('A102').setValue(code);
    sheetY.getRange('B102').setValue(tyaku);
    sheetY.getRange('C102').setValue(ds);
    sheetY.getRange('E102').setValue(newSpId);
    sheetY.getRange('D102').setValue(nendo);
    sheetY.getRange('H102').setValue(genmei);
    const kasira = huri ? huri.slice(0, 1) : "";
    sheetY.getRange('V102').setValue(kasira);
    sheetY.getRange('W102').setValue(huri);
    sheetY.getRange('I102').setValue(kinds);
    sheetY.getRange('T102').setValue(add);
    sheetY.getRange('X102').setValue(client);
    sheetY.getRange('U102').setValue(tan);
    if(uke == ""){
      sheetY.getRange('M88').setValue(0);
    }else{
      const kin = Number(String(uke).replace(/[\n,]/g, ""));
      sheetY.getRange('M88').setValue(kin);
    }
    sheetY.getRange('R102').setValue(start);
    sheetY.getRange('S102').setValue(finish);
    sheetY.getRange('AJ102').setValue(keiyaku);
    sheetY.getRange('AK102').setValue(spasceUrl);
    sheetY.getRange('Z2').setValue(toroku);
    sheetY.getRange('A101').setValue(nendo);

    sheetY.getRange('AF102').setValue(motoTan);
    sheetY.getRange('AG102').setValue(motoMail);
    sheetY.getRange('AH102').setValue(ekimu);
    sheetY.getRange('AI102').setValue(hatyuNo);

    sheetY.getRange('AB102').setValue(memo);//メモ
    
    sheetY.getRange('AR102').setValue("請求用");//請求用にする

    sheetY.getRange('L88').setValue(kFileId);//契約書ファイルID記録初回用
    //金額,着工,完了,発注元,発行日,住所,担当,物件名
    const vals = uke + "◆" + start + "◆" + finish + "◆" + client + "◆" + keiyaku + "◆" + add + "◆" + motoTan + "◆" + genmei;
    sheetY.getRange('AR88').setValue(vals);//注文書内容の記録
    
    //OneDriveのURLのとこにスプレットシートのURLを登録する
    const newFile = DriveApp.getFileById(newSpId);
    const ssUrl = newFile.getUrl();
    sheetY.getRange('AC102').setValue(ssUrl);//OneDriveUrl(スプレットシートのURL)
    //請負金額に数式をセットする
    sheetY.getRange('J102').setFormula("=K87");
        
    //◆◆//工事台帳データベースに記録◇◇
    const row = rw;//◆◆工事台帳のRowを送る
    //とりあえず年度row作成して記録
    const nenRow = nendo + rw;
    sheetY.getRange('A1').setValue(nenRow);
    sheetY.getRange('AD102').setValue(row);//ROWの記録

    //台帳に登録(もうGAS化始まったので)
    const val1 = sheetY.getRange('A102:AJ102').getValues();
    ssK.getRange('A' + rw + ':AJ' + rw).setValues(val1);
    SpreadsheetApp.flush();
    現場検索リスト更新(nendo);//リスト更新
    //ここで注文書をチャットにあげる
    const cleanUke = String(uke).replace(/[\n,]/g, "");
    tyumonSpace(kFileId, cleanUke, genmei, kinds);//注文書をスペースにあげる
  }
}
//😃😃注文書をスペースに投稿する
function tyumonSpace(kFileId, uke, genmei, kinds, ts, fUrl) {
  // ts や fUrl が渡されない場合、これらは undefined になります
  const webhookUrl = chatUrl("注文書丸山");

  try {
    // 1. 金額の整形（カンマ区切り）
    const cleanUke = String(uke).replace(/[^\d]/g, ""); 
    const formattedUke = cleanUke ? Number(cleanUke).toLocaleString() : "0";

    // 2. ファイルURLの決定
    let fileUrl = "";
    if (kFileId) {
      // kFileId がある場合は最優先で取得
      fileUrl = DriveApp.getFileById(kFileId).getUrl();
    } else if (fUrl) {
      // kFileId がなくて fUrl がある場合（請求書などの予備）
      fileUrl = fUrl;
    }

    // 3. タイトルの決定
    // ts が未定義(注文書時) or 空文字なら「注文書」、そうでなければ ts の中身を使う
    const title = (ts === "請求書") ? "請求書" : "注文書";

    // 4. メッセージ作成と送信
    const msg = `【${title}】${genmei} ${kinds} ${formattedUke}円\n※保管お願いします\n${fileUrl}`;
    
    sendChat(webhookUrl, msg);

  } catch (e) {
    Logger.log("tyumonSpaceエラー: " + e.message);
  }
}

//😃現場コード(nendo)生成
function 現場コード(nen,rw,genA,ds){
  const date = new Date();
  const day = Utilities.formatDate(date,"JST","yyMMdd");
  if(ds == "D"){
    const genId = "D-" + nen + "-" + day + "-" + rw;
    const genmei = genA
    return [genId,genmei];
  }else{
    const genId = "S-" + nen + "-" + day + "-" + rw;
    let getu = Utilities.formatDate(date,"JST","M月");
    const genmei = genA + " " + rw + "-" + getu;
    return [genId,genmei];
  };
}

//😃😃予算書保存フォルダIDと工事台帳ROWを返す
function 保存フォルダと台帳(nendoA,ds){
  Logger.log({nendoA,ds});
  const nendo = Number(nendoA);
  const idN = getSPId("年度フォルダ");
  const sheets = SpreadsheetApp.openById(idN);
  const sheet = sheets.getSheetByName('生成履歴');
  let rw = sheet.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i = 1; i < rw + 1; i++){
    if(sheet.getRange('A' + i).getValue() == nendo){
      const foldId = sheet.getRange('C' + i).getValue();//予算書フォルダID
      const koziId = sheet.getRange('F' + i).getValue();//工事台帳DBファイルID
      if(ds == "D"){
        return [foldId,koziId];
      }else{
        const oya = DriveApp.getFolderById(foldId);
        const folders = oya.getFoldersByName("小規模");
        if(folders.hasNext()){
          const folder = folders.next();
          const fId = folder.getId();
          return [fId,koziId];
        }
      }
    }
  }
  Logger.log("年度が見つかりませんでした: " + nendo);
  return null;
}   
//😃😃初回現場登録時の契約書ファイルを一時フォルダから移動する
function 契約書移動(kFileId,nendo,gen,ds){
  Logger.log("💠kFileId " + kFileId);
  if(kFileId !== ""){
    const file = DriveApp.getFileById(kFileId);
    const folder = get注文請書Fld(nendo, gen, ds);//コードyosan2にある
    file.moveTo(folder);
  }
}

//😃😃新規現場が登録されたので※※検索用リスト更新※※する
function 現場検索リスト更新(nendo){
  //nendo = 2025;
  const id = 工事台帳(nendo);
  const sheets = SpreadsheetApp.openById(id);
  const sheetL = sheets.getSheetByName('工事リスト');
  const sheetK = sheets.getSheetByName('検索用1');
  const sheetK2 = sheets.getSheetByName('検索用2');
  const sheetK3 = sheets.getSheetByName('大小現場分け');
  let row = sheetL.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let val = sheetL.getRange('A1:H' + row).getValues();
  sheetK.getRange('A1:I1000').clearContent();//クリア
  sheetK.getRange('A1:H' + row).setValues(val);
  val = sheetL.getRange('V1:V' + row).getValues();//頭文字
  sheetK.getRange('I1:I' + row).setValues(val);
  val = sheetL.getRange('AD1:AD' + row).getValues();//ROW
  sheetK.getRange('F1:F1000').clearContent();//クリア
  sheetK.getRange('F1:F' + row).setValues(val);

  //◆◆フィルター昇順◆◆◆
  if(sheetK.getFilter() !== null){
    sheetK.getFilter().remove();
  };
  let filter = sheetK.getRange('A1:I' + row).createFilter();
  filter.sort(9, true);
  //検索表作成（関数入り表をコピー値だけを別セルにペースト）
  SpreadsheetApp.flush();
  sheetK.getRange('K1:BI185').copyTo(sheetK.getRange('BL1:DJ185'), {contentsOnly:true});

  //全現場検索用（大規模小規模分けリスト）
  let rowK = sheetK.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let rowA = sheetK.getRange('DP1000').getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
  sheetK.getRange('DM2:DR' + rowA).clearContent();
  
  let dataA = sheetK.getRange('A1:I' + row).getValues();
  let dataB = dataA.filter(record => record[2] === "D")
  let y = 0
  for(let i = 2; i < dataB.length + 2; i++){
    sheetK.getRange('DM' + i).setValue(dataB[y][7]);
    sheetK.getRange('DN' + i).setValue(dataB[y][4]);
    sheetK.getRange('DO' + i).setValue(dataB[y][5]);
    y = y + 1;
  }
  let dataC = dataA.filter(record => record[2] === "S" || record[2] == "S2");
  y = 0
  for(let s = 2; s < dataC.length + 2; s++){
    sheetK.getRange('DP' + s).setValue(dataC[y][7]);
    sheetK.getRange('DQ' + s).setValue(dataC[y][4]);
    sheetK.getRange('DR' + s).setValue(dataC[y][5]);
    y = y + 1;
  }
  //◆◆◆終了工事を除くフィルター◆◆◆
  dataA = sheetK.getRange('A1:I' + row).getValues();
  dataB = dataA.filter(record => record[1] === "職人報告" || record[1] === "注文書あり" || record[1] === "着工" || record[1] === "請求あり" || record[1] === "終了" || record[1] === "工事着工" )
  let row2 = sheetK2.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  sheetK2.getRange('A2:I' + row2).clearContent();
  sheetK2.getRange(2, 1, dataB.length, dataB[0].length).setValues(dataB);
  //検索表作成（関数入り表をコピー値だけを別セルにペースト）
  SpreadsheetApp.flush();
  sheetK2.getRange('K1:BI84').copyTo(sheetK2.getRange('BL1:DJ84'), {contentsOnly:true});

  //◆◆◆大規模小規模の分け◆◆◆
  dataA = sheetK2.getRange('A2:I' + row2).getValues();
  dataB = dataA.filter(record => record[2] === "D")
  let row3 = sheetK3.getLastRow();
  if(row3 !== 0){
    sheetK3.getRange('A2:S' + row3).clearContent();
  }
  sheetK3.getRange(2, 1, dataB.length, dataB[0].length).setValues(dataB);

  dataA = sheetK2.getRange('A2:I' + row2).getValues();
  dataB = dataA.filter(record => record[2] === "S" || record[2] === "S2")
  sheetK3.getRange(2, 11, dataB.length, dataB[0].length).setValues(dataB); 

  return "OK";
}

//😃😃顧客の登録😃😃
function rezistClientGetList(dt){
  const idK = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(idK);//業者・顧客・担当台帳
  const sheetK = sss.getSheetByName('顧客台帳');
  const row = sheetK.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
  sheetK.getRange('A' + row).setValue(row);//ROW
  sheetK.getRange('B' + row).setValue(dt[1]);//顧客
  sheetK.getRange('C' + row).setValue(dt[2]);//フリガナ
  sheetK.getRange('D' + row).setValue(dt[3]);//種別
  sheetK.getRange('E' + row).setValue(dt[4]);//郵便
  sheetK.getRange('F' + row).setValue(dt[5]);//住所
  sheetK.getRange('H' + row).setValue(dt[6]);//TEL
  sheetK.getRange('I' + row).setValue(dt[7]);//FAX
  sheetK.getRange('L' + row).setValue(dt[8]);//担当
  sheetK.getRange('M' + row).setValue(dt[9]);//担当携帯
  sheetK.getRange('N' + row).setValue(dt[10]);//担当メール
  sheetK.getRange('O' + row).setValue(dt[11]);//メモ
  sheetK.getRange('P' + row).setValue(dt[12]);//登録者
  SpreadsheetApp.flush();
  //顧客リストの更新
  const judge = clientSenbetu();
  if(judge == "OK"){
    const vals = sheetK.getRange('B2:B' + row).getValues().flat();
    vals.unshift("選択");
    return vals;
  }
}

//◆◆◆顧客区別◆◆◆
function clientSenbetu(){
  const idK = getSPId("業者顧客担当");
  const sheets = SpreadsheetApp.openById(idK);
  const sheetA = sheets.getSheetByName('顧客台帳');
  const sheetB = sheets.getSheetByName('顧客分け（元請）');
  const sheetC = sheets.getSheetByName('顧客分け（検索表）');
  let rw = sheetA.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let dataA = sheetA.getRange('A1:K' + rw).getValues();
  let dataB = dataA.filter(record => record[3] === "元請")
  let rw2 = sheetB.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  sheetB.getRange('A2:K' + rw2).clearContent();
  sheetB.getRange(2, 1, dataB.length, dataB[0].length).setValues(dataB);
  //検索表作成
  val = sheetA.getRange('A1:K' + rw).getValues();
  sheetC.getRange('A1:K200').clearContent();
  sheetC.getRange('A1:K' + rw).setValues(val);
  let filter = sheetC.getRange('A2:K' + rw);
  filter.sort(3);
  SpreadsheetApp.flush();
  sheetC.getRange('N1:CG72').copyTo(sheetC.getRange('CK1:FB72'),{contentsOnly:true});
  return "OK"; 
}