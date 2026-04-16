//初期読み込みパスワード確認＆インフォ＆担当リスト＆顧客リスト取得
function passKaku(fId, access, syaId, pass) {
  Logger.log(fId + "," + access + "," + syaId + "," + pass);
  const accessB = Number(access);
  const passB = Number(pass);
  const id = getSPId("業者顧客担当") 
  const sss = SpreadsheetApp.openById(id);//業者・顧客・担当者台帳
  const ss2 = sss.getSheetByName('顧客分け（検索表）');
  const last2 = ss2.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const ss3 = sss.getSheetByName('担当者分け');
  const last3 = ss3.getRange("H1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const row = Number(syaId) - 100;
  Logger.log("last3 " + last3 + ", row " + row);

  const rnRow = ss3.getRange('W1:W' + last3).getValues().flat();
  const rnName = ss3.getRange('H2:H' + last3).getValues().flat();
  const rnPass = ss3.getRange('V1:V' + last3).getValues().flat();
  const accessA = ss3.getRange('F2').getValue();
  const tan = ss3.getRange('A' + row).getValue();
  if (access !== "gibson" && accessA !== accessB) {
    Logger.log("access NG");
    return "NG";
  }
  const nentuki = 年度年月();
  const nendo = nentuki[0];
  let ds = "";
  for (let i = 0; i <= last3; i++) {
    if(i == last3){
      Logger.log("pass NG2");
      return "NG";
    }else if (Number(rnRow[i]) === row) {
      const passA = Number(rnPass[i]);
      if (passA !== passB) {
        Logger.log("pass NG1");
        return "NG";
      }else{
        break;
      }
    }
  }
  if(fId == "rezi"){
    const data = {
      nameList: rnName,
      kokyakuList: ss2.getRange('B2:B' + last2).getValues().flat(),
      rezistant: tan,
      nendo: nendo
    }
    const json = JSON.stringify(data);
    return json;
  }else{
    const sssY = SpreadsheetApp.openById(fId);
    const ssY = sssY.getSheetByName('シート1');
    const dsF = () => {
      if(ssY.getRange('C102').getValue() == "D" || ssY.getRange('C102').getValue() == "S2"){
        return "D";
      }else{
        return "S";
      }
    }
    const ds = dsF();
      const rowF = () => {
      if (ds === "D") return 102;
      if (ds === "S") return 19;
      return null;
    };
    const r = rowF();
    const rn = ssY.getRange('A' + r + ':AJ' + r).getValues()[0];

    // 日付整形
    const formatIfDate = (val) => val ? Utilities.formatDate(new Date(val), "Asia/Tokyo", "yyyy年M月d日") : "";

    //const start = Utilities.formatDate(new Date(ssY.getRange('R').getValue()),"JST","yyyy年M月d日");
    //const finish = Utilities.formatDate(new Date(ssY.getRange('S').getValue()),"JST","yyyy年M月d日");

    const data = {
      info: rn.map((v, idx) => {
        // 日付変換対象のみ整形
        if ([17, 18, 26, 35].includes(idx)) {
          return formatIfDate(v);
        }
        return v;
      }),
      nameList: rnName,
      kokyakuList: ss2.getRange('B2:B' + last2).getValues().flat(),
      selectedTanto: rn[6], // 担当者
      selectedKokyaku: rn[23], // 顧客名
      otherFields: {
        seikyukanryo: formatIfDate(rn[26]),
        memo: rn[27],//メモ
        mototan: rn[31],//元請け担当
        tantoMail: rn[32],//元請担当メール
        gyomu: rn[33],//工事・役務(大成）
        haseko: rn[34],//長谷工CSV
        keiyaku: formatIfDate(rn[35])//契約日
      },
      nendo: nendo,//現在の年度
      ds: ds,
      toroku: ssY.getRange('Z2').getValue(),//登録者
      chatUrl: ssY.getRange('AK102').getValue()//チャットURL
    };
    const json = JSON.stringify(data);
    Logger.log(json);
    return json;
  }
}

//◆◆現場内容の変更
function genInfoHen(json) {
  // try {
    const data = JSON.parse(json);
    Logger.log(data);

    const sss = SpreadsheetApp.openById(data.fId);
    const ss = sss.getSheetByName("シート1");

    const rF = () => {
      if (data.ds == "D" || data.ds == !"S2") {
        return 102;
      } else if (data.ds == "S") {
        return 19;
      }
    };
    const r = rF();

    const old = ss.getRange('U' + r).getValue();

    ss.getRange('H' + r).setValue(data.gen);
    ss.getRange('I' + r).setValue(data.naiyo);
    //請負金額は（J102)は複数契約書に対応するため数式になってる場合位があるので確認して、変更する必要がある
    const formula = ss.getRange('J' + r).getFormula();
    if(formula == ""){
      ss.getRange('J' + r).setValue(data.uke);
    }
    //const start = new Date(data.st);
    ss.getRange('R' + r).setValue(data.st);

    //const finish = new Date(data.fn);
    ss.getRange('S' + r).setValue(data.fn);

    ss.getRange('T' + r).setValue(data.add);
    ss.getRange('U' + r).setValue(data.tan);
    ss.getRange('X' + r).setValue(data.client);
    ss.getRange('AB' + r).setValue(data.memo);
    ss.getRange('AF' + r).setValue(data.motoTan);
    ss.getRange('AG' + r).setValue(data.mMail);
    ss.getRange('AI' + r).setValue(data.csv);
    ss.getRange('AJ' + r).setValue(data.keiyaku);
    ss.getRange('AK' + r).setValue(data.chatUrl);

    if (old !== data.tan) {
      const tInfo = Get社員情報(data.tan);
      const tel = tInfo[2];
      const mail = tInfo[1];
      ss.getRange('Y' + r).setValue(tel);
      ss.getRange('Z' + r).setValue(mail);
    }

    SpreadsheetApp.flush();

    const vals = ss.getRange('A' + r + ':AJ' + r).getValues(); 
    const nentuki = String(ss.getRange('A1').getValue());
    if(ss.getRange('A1').getValue() !== "ID" && ss.getRange('A1').getValue() !== ""){
      const nen = Number(nentuki.slice(0, 4));
      const row = Number(nentuki.slice(4));
      const ssK = 工事リストシート(nen); // 工事台帳の工事リストを取得

      ssK.getRange('A' + row + ':AJ' + row).setValues(vals);
    }
    return "OK";
    
  // } catch (er) {
  //   return er.message; 
  // }
}
//◆◆◆◆業者支払管理呼出◆◆◆◆
  function getPay(id, ds, access, genName){
    // id = "1g1fMAmUch7Fl12ajwo4HRDyZf3FlHt59qoF6p0Bqv6E";
    // ds = "D";
    // access = 34316;
    // genName = "サントラビル"
    const sssY = SpreadsheetApp.openById(id);
    const ssY = sssY.getSheetByName('シート1');
    const accessP = 安全パスワード(access);
    if (accessP !== "OK") {
      Logger.log("NGアクセス");
      return "NGアクセス";
    } else {
      Logger.log(ds);
      let vals = "";
      if (ds == "D" || ds == "S2") {
        const allVals = ssY.getRange('A1:AJ41').getValues(); // ★まとめて一気に取る
        const gyoKanriF = () => {
          const numF = () => {
            const names = ssY.getRange('F1:F40').getValues().flat();
            for (let i = 3; i < names.length; i++) {
              if (names[i] == "") {
                return i;//配列は０からなので+1さらに追加まで入れたいので+1
              }
            }
          };
          const num = numF();
          const numA = (num - 3) / 3 * 2;
          let rnG = [[numA]];
          const rgF = (g) => {
            const n = g - 1;
            const index = [4,5,7,8,10,11,13,14,16,17,19,20,22,23,25,26,28,29,31,32,34,35,37,38,40,41];
            return index[n];
          };
          let kindsA = "";
          let gyoA = "";
          for (let g = 0; g <= numA; g++) {
            if (g == 0) {
              let rn = ["業者名", "工種","発注金額", "支払累計", "支払残金", "1月支払", "2月支払", "3月支払", "4月支払", "5月支払", "6月支払", "7月支払", "8月支払", "9月支払", "10月支払", "11月支払", "12月支払"];
              rnG.push(rn);
            } else {
              const rg = rgF(g);
              const rowVals = allVals[rg - 1]; // ★事前取得したallValsから行を取る
              const kindsF = () => {
                if(rg == 5 || rg == 8 || rg == 11 || rg == 14 || rg == 17 || rg == 20 || rg == 23 || rg == 26 || rg == 29 || rg == 32 || rg == 35 || rg == 38 || rg == 41){
                  return  "追加 " + kindsA;
                }else{
                  kindsA = rowVals[0]
                  return kindsA;
                }
              }
              const kinds = kindsF();
              const gyoF = () => {
                if(rg == 5 || rg == 8 || rg == 11 || rg == 14 || rg == 17 || rg == 20 || rg == 23 || rg == 26 || rg == 29 || rg == 32 || rg == 35 || rg == 38 || rg == 41){
                  return  gyoA;
                }else{
                  gyoA = rowVals[5]
                  return gyoA;
                }
              }
              const gyo = gyoF();
              let rn = [gyo, kinds,rowVals[6], rowVals[8], rowVals[11], rowVals[13], rowVals[15], rowVals[17], rowVals[19], rowVals[21], rowVals[23], rowVals[25], rowVals[27], rowVals[29], rowVals[31], rowVals[33], rowVals[35]];
              rnG.push(rn);
            }
          }
          return rnG;
        };

        const seikyuF = () => {
          const num = ssY.getRange('M101').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
          const num2 = (num - 87) + 2;
          const end = num2 + 87;
          let rnS = [[num2]];
          let rn = ["請負契約","請負金額", "支払計", "売上計", "請負残金", "1月請求", "2月請求", "3月請求", "4月請求", "5月請求", "6月請求", "7月請求", "8月請求", "9月請求", "10月請求", "11月請求", "12月請求"];
          rnS.push(rn);

          const vals = ssY.getRange('K87:AP87').getValues()[0];
          const zan = vals[0] - vals[29];
          let rn2 = ["請負合計", vals[0], vals[31], vals[29], zan, vals[3], vals[5], vals[7], vals[9], vals[11], vals[13], vals[15], vals[17], vals[19], vals[21], vals[23], vals[25]];
          rnS.push(rn2);

          for (let s = 88; s <= end; s++) {
            const vals = ssY.getRange('K' + s + ':AQ' + s).getValues()[0];
            const pay = "";
            const rn3 = [vals[0],vals[2], pay, vals[30], vals[32], vals[3], vals[5], vals[7], vals[9], vals[11], vals[13], vals[15], vals[17], vals[19], vals[21], vals[23], vals[25]];
            rnS.push(rn3);
          }
          return rnS;
        };

        const val = ssY.getRange('H102').getValue();
        const data = {
          gyoKanri: gyoKanriF(),
          seikyu: seikyuF(),
          gen: val
        };
        const json = JSON.stringify(data);
        // Logger.log(json); // 必要なら本番でコメントアウト
        return json;
      }else if(ds == "S"){//小規模予算書//小規模予算書//小規模予算書//小規模予算書//小規模予算書//小規模予算書//小規模予算書
        const allVals = ssY.getRange('U1:AT41').getValues(); // ★まとめて一気に取る
        const gyoKanriF = () => {
          const numF = () => {
            const names = ssY.getRange('U1:U10').getValues().flat();
            for (let i = 1; i < names.length; i++) {
              if (names[i] == "") {
                return i;//配列は０からなので+1さらに追加まで入れたいので+1
              }
            }
          };
          const num = numF() + 1;//""のROW
          const numA = num - 2;//登録の数
          let rnG = [[numA]];
          
          for (let g = 0; g <= numA; g++) {
            if (g == 0) {
              let rn = ["業者名", "工種","発注金額", "支払累計", "支払残金", "10月支払", "11月支払", "12月支払", "1月支払", "2月支払", "3月支払", "4月支払", "5月支払", "6月支払", "7月支払", "8月支払", "9月支払"];
              //請負金額,支払計,売上計,請求残金,10月請求,11月請求,12月請求,1月請求,2月請求,3月請求,4月請求,5月請求,6月請求,7月請求,8月請求,9月請求払
              rnG.push(rn);
            } else {
              const gyo = allVals[g][0];
              const kinds = allVals[g][25]
              let rn = [gyo, kinds,allVals[g][17],allVals[g][18],allVals[g][19],allVals[g][1],allVals[g][2],allVals[g][3],allVals[g][4],allVals[g][5],allVals[g][6],allVals[g][7],allVals[g][8],allVals[g][9],allVals[g][10],allVals[g][11],allVals[g][12]];
              rnG.push(rn);
            }
          }
          return rnG;
        };

        const seikyuF = () => {
          let rnS = [[2]];
          let rn = ["請負契約","請負金額", "支払計", "売上計", "請負残金", "10月請求", "11月請求", "12月請求", "1月請求", "2月請求", "3月請求", "4月請求", "5月請求", "6月請求", "7月請求", "8月請求", "9月請求"];
          rnS.push(rn);

          const vals = ssY.getRange('U16:AHP16').getValues()[0];
          const uke = ssY.getRange('J19').getValue();
          const pay = ssY.getRange('AH15').getValue();
          const sales = ssY.getRange('AH16').getValue();
          let rn2 = ["注文書1（本工事）", uke, pay, sales, "", vals[1], vals[2], vals[3], vals[4], vals[5], vals[6], vals[7], vals[8], vals[9], vals[10], vals[11], vals[12]];
          rnS.push(rn2);
          return rnS;
        };

        const val = ssY.getRange('H19').getValue();
        const data = {
          gyoKanri: gyoKanriF(),
          seikyu: seikyuF(),
          gen: val
        };
        const json = JSON.stringify(data);
        Logger.log(json); // 必要なら本番でコメントアウト
        return json;
      }
    }
  }

  //😃◇😃請求登録ゾーン😃◇😃請求登録ゾーン😃◇😃請求登録ゾーン😃◇😃請求登録ゾーン😃◇😃請求登録ゾーン😃◇😃請求登録ゾーン
  //◆◆請求登録◆◆◆◆請求登◆◆
  function climeRezi(genS,tukiS,kindsS,ukeoiS,dekiS,madeS,seiS,zanS,memoS,fId,ds,nendo,sctA){
    Logger.log(genS + ',' + tukiS + ',' + kindsS + ',' + ukeoiS + ',' + dekiS + ',' + madeS + ',' + seiS + ',' + zanS + ',' + memoS + ',' + fId + ',' + ds + "," + sctA);
    //記録するROWをsctから判断
    const rF = () => {
      const sct = Number(sctA);
      return 87 + sct
    }
    const r = rF();
    const sss = SpreadsheetApp.openById(fId);
    const ss = sss.getSheetByName('シート1');
    
    const dId = 工事台帳(nendo);
    Logger.log(dId);
    const sssD = SpreadsheetApp.openById(dId);
    const ssD = sssD.getSheetByName('工事リスト');
    const rowF = () => {
      const last = ssD.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
      const valH = ssD.getRange('H1:H' + last).getValues().flat();
      for(let i=0; i<last; i++){
        if(valH[i] == genS){
          return i + 1;
        }
      }
    }
    const row = rowF();
    //日付記録(請求書複数登録のため)
    const day = Utilities.formatDate(new Date,"JST","yyyy年M月d日");
    const tukiColF = () => {
      if(ds == "D" || ds =="S2"){
        return (Number(tukiS) * 2) + 12;
      }else if(ds == "S"){
        const tukiN = Number(tukiS);
        if(tukiN > 9){
          return Number(tukiS) + 12;
        }else{
          return Number(tukiS) + 24;
        }
      }
    }
    const tukiCol = tukiColF(); 

    if(ds == "D" || ds =="S2"){
      const gname = ss.getRange('H102').getValue();
      if(gname == genS){
        const nameD = ssD.getRange('H' + row).getValue();
        Logger.log(nameD + "," + genS);
        if(nameD == genS){
          const seiN = Number(seiS.replace(/,/g, "")) / 1.1;//税抜きに変更する
          ss.getRange(r,tukiCol).setValue(seiN);//請求記録
          const dayCol = tukiCol + 1;
          ss.getRange(r,dayCol).setValue(day);//請求日記録
          ss.getRange('AB102').setValue(memoS);//メモの記録
          ss.getRange('AD102').setValue(row);//期またぎの場合の対応
          if(kindsS == "請求済"){//請求完了の場合の日付記録
            ss.getRange('AA102').setValue(day);
            //ss.getRange('B102').setValue("請求済");請求済にしてしまうと稼働現場から消えてしまう
          }else{

          }
          vals = ss.getRange('A102:AC102').getValues();
          ssD.getRange('A' + row + ':AC' + row).setValues(vals);//台帳記録
          ssD.getRange('AD' + row).setValue(row);//期またぎ対応
          //◆VBAメールを入れる
          //請求登録メール送信VBA(nendoA,seiN,ds,genS,tukiS,ukeoiS,dekiS,madeS,zanS,nendo);
          return "OK"
        }else{
          return "台帳とファイルが違う可能性あり";
        }
      }else{
        return "現場名とファイルIDが違う可能性あり";
      }
    }else if(ds == "S"){
      const gname = ss.getRange('H19').getValue();
      Logger.log("gname " + gname + " genS " + genS);
      Logger.log(row);
      if(gname == genS){
        const nameD = ssD.getRange('H' + row).getValue();
        Logger.log(nameD + "," + genS);
        if(nameD == genS){
          const seiN = Number(seiS.replace(/,/g, "")) / 1.1;//税抜きに変更する
          ss.getRange(16,tukiCol).setValue(seiN);//請求記録
          ss.getRange(17,tukiCol).setValue(day);//請求日記録
          ss.getRange('AB19').setValue(memoS);//メモの記録
          ss.getRange('AD19').setValue(row);//期またぎ対応
          if(kindsS == "請求済"){//請求完了の場合の日付記録
            ss.getRange('AA19').setValue(day);
            //ss.getRange('B19').setValue("請求済");
          }else{
            
          }
          vals = ss.getRange('A19:AB19').getValues();
          ssD.getRange('A' + row + ':AB' + row).setValues(vals);//台帳記録
          ssD.getRange('AD' + row).setValue(row);//期またぎ対応
          //◆VBAメールを入れる
          //請求登録メール送信VBA(nendoA,seiN,ds,genS,tukiS,ukeoiS,dekiS,madeS,zanS,nendo);
          return "OK"
        }else{
          return "台帳とファイルが違う可能性あり";
        }
      }else{
        return "現場名とファイルIDが違う可能性あり";
      }
    }
  }
  //😃😃契約書の読み込み時の記録を読み込む
  function getKeiyakuInfo(fId,sct,ds){
    const sss = SpreadsheetApp.openById(fId);
    const ss = sss.getSheetByName('シート1');
    const rowF = () => {
      const res = 87 + Number(sct);
      return res;
    }
    const row = rowF();
    if(ds == "D" || ds == "S2"){
      const vals = ss.getRange('AR' + row).getValue();
      if(vals == ""){
        return "なし";
      }else{
        return vals;
      }
    }else{
      const vals = ss.getRange('T17').getValue();
      if(vals == ""){
        return "なし";
      }else{
        return vals;
      }  
    }
  }

  //◆◆請求処理◆大成有楽不動産◆
  function taiseiSei(genmei,kousyu,tukiS,ukeoiS,dekiS,madeS,seiS,zanS,fId,ds,tanM,mailM,ke,start,finish,seibi,add,num,end,zouGen,henSt,henFn,keiyaku,firstEnd,moto,manBuild,motoSt,motoFn,motoSct){
    Logger.log("genmei:" + genmei + ", kousyu:" + kousyu + ", tukiS:" + tukiS + ", ukeoiS:" + ukeoiS + ", dekiS:" + dekiS + ", madeS:" + madeS + ", madeS:" + madeS + ", zanS:" + zanS + ", fId:" + fId + ", ds:" + ds + ", tanM:" + tanM + ", mailM:" + mailM + ", ke:" + ke + ", start:" + start + ", finish:" + finish + ", seibi:" + seibi + ", add:" + add + ", num:" + num + ", end:" + end + ", zouGen:" + zouGen + ", henSt:" + henSt + ", henFn:" + henFn + ", keiyaku:" + keiyaku + ", firstEnd:" + firstEnd + ", moto:" + moto + ", manBuild:" + manBuild + ", motoSt:" + motoSt + ", motoFn:" + motoFn + ", motoSct:" + motoSct);
    
    入力情報反映(fId,ds,tanM,mailM,ke,start,finish,add,henSt,henFn,keiyaku);//請求時に登録のなかった入力情報を反映する
    //請求金額
    const seikyu = Number(seiS.replace(/,/g, "")) / 1.1//請求金額を税別に
    //工事か役務か判断して請求書PDFを作成メール下書き保存迄
    const kouzi = genmei + kousyu;
    if(ke == 1){
      const hId = ひな形("大成工事請求");
      const sss = SpreadsheetApp.openById(hId);//工事ひな形
      const ss1 = sss.getSheetByName('様式C（工事）');
      const ss1B = sss.getSheetByName('印なし１');
      const ss2 = sss.getSheetByName('工事検査願（工事）');
      const ss2B = sss.getSheetByName('印なし２');
      const ssD1 = sss.getSheetByName('様式D（工事）');
      const ssD2 = sss.getSheetByName('様式D（工事）減');
      const ssD3 = sss.getSheetByName('様式D（工事）工期のみ');
      const ssD4 = sss.getSheetByName('印なし３増');
      const ssD5 = sss.getSheetByName('印なし３減');
      const ssD6 = sss.getSheetByName('印なし工期のみ');
      
      //請求書
      ss1.getRange('AY4').setValue(seibi);//発行日
      ss1B.getRange('AY4').setValue(seibi);//発行日
      ss2.getRange('P6').setValue(seibi);//発行日
      ss2B.getRange('P6').setValue(seibi);//発行日
      ssD1.getRange('AS4').setValue(seibi);//発行日
      ssD2.getRange('AS4').setValue(seibi);//発行日
      ssD3.getRange('AS4').setValue(seibi);//発行日
      ssD4.getRange('AS4').setValue(seibi);//発行日
      ssD5.getRange('AS4').setValue(seibi);//発行日
      ssD6.getRange('AS4').setValue(seibi);//発行日

      const kouki = start + "～" + finish;
      ss1.getRange('B31').setValue(kouki);//工期
      ss1B.getRange('B31').setValue(kouki);//工期

      //出来高査定日（請求が1回のみの場合は入力しない）
      if(firstEnd == "1回のみ"){
        ss1.getRange('K33').setValue("");
        ss1B.getRange('K33').setValue("");
      }else{
        ss1.getRange('K33').setValue(seibi);//請求日
        ss1B.getRange('K33').setValue(seibi);//請求日
      }
      
      
      ss1.getRange('T31').setValue(kouzi);//工事名
      ss1.getRange('T35').setValue(ukeoiS);//請負金額
      ss1B.getRange('T31').setValue(kouzi);
      ss1B.getRange('T35').setValue(ukeoiS);//請負金額
      ss1.getRange('T37').setValue(zouGen);//精算増減
      ss1B.getRange('T37').setValue(zouGen);//精算増減
      ss1.getRange('T39').setValue(dekiS);//出来高
      ss1.getRange('T41').setValue(madeS);//前回迄
      ss1B.getRange('T39').setValue(dekiS);//出来高
      ss1B.getRange('T41').setValue(madeS);//前回迄
      const nankai = "（第" + num + "回）";
      ss1.getRange('G43').setValue(nankai);//第何回
      ss1B.getRange('G43').setValue(nankai);//第何回
      

      //
      ss2.getRange('H23').setValue(kouzi);//工事名
      ss2B.getRange('H23').setValue(kouzi);//工事名
      ss2.getRange('H26').setValue(add);//工事場所
      ss2B.getRange('H26').setValue(add);//工事場所

      //精算増減
      ssD1.getRange('L31').setValue(kouzi);//工事名
      ssD2.getRange('L31').setValue(kouzi);//工事名
      ssD3.getRange('L31').setValue(kouzi);//工事名
      ssD4.getRange('L31').setValue(kouzi);//工事名
      ssD5.getRange('L31').setValue(kouzi);//工事名
      ssD6.getRange('L31').setValue(kouzi);//工事名

      ssD1.getRange('AV31').setValue(keiyaku);//契約日
      ssD2.getRange('AV31').setValue(keiyaku);//契約日
      ssD3.getRange('AV31').setValue(keiyaku);//契約日
      ssD4.getRange('AV31').setValue(keiyaku);//契約日
      ssD5.getRange('AV31').setValue(keiyaku);//契約日
      ssD6.getRange('AV31').setValue(keiyaku);//契約日

      ssD1.getRange('L33').setValue(add);//住所
      ssD2.getRange('L33').setValue(add);//住所
      ssD3.getRange('L33').setValue(add);//住所
      ssD4.getRange('L33').setValue(add);//住所
      ssD5.getRange('L33').setValue(add);//住所
      ssD6.getRange('L33').setValue(add);//住所

      const zoGenF = () => {
        if(zouGen !== 0){
          ssD1.getRange('E23').setValue("レ");//レ点
          ssD2.getRange('E23').setValue("レ");//レ点
          ssD3.getRange('E23').setValue("レ");//レ点
          ssD4.getRange('E23').setValue("レ");//レ点
          ssD5.getRange('E23').setValue("レ");//レ点
          ssD6.getRange('E23').setValue("レ");//レ点
          if(zouGen > 0){
            return "増";
          }else{
            return "減";
          }
        }else{
          ssD1.getRange('E23').setValue("");//レ点
          ssD2.getRange('E23').setValue("");//レ点
          ssD3.getRange('E23').setValue("");//レ点
          ssD4.getRange('E23').setValue("");//レ点
          ssD5.getRange('E23').setValue("");//レ点
          ssD6.getRange('E23').setValue("");//レ点
          return "なし";
        }
      }
      const zoGen = zoGenF();//精算増減judge
      if(zoGen == "増" || zoGen == "減"){
        //請求書に精算増減の事由
        ss1.getRange('AF37').setValue("別紙のとおり");

        ssD1.getRange('AX38').setValue(ukeoiS);//請負
        ssD2.getRange('AX38').setValue(ukeoiS);//請負
        ssD3.getRange('AX38').setValue(ukeoiS);//請負
        ssD4.getRange('AX38').setValue(ukeoiS);//請負
        ssD5.getRange('AX38').setValue(ukeoiS);//請負
        ssD6.getRange('AX38').setValue(ukeoiS);//請負

        ssD1.getRange('AX40').setValue(zouGen);//請負
        ssD2.getRange('AX40').setValue(zouGen);//請負
        ssD3.getRange('AX40').setValue(zouGen);//請負
        ssD4.getRange('AX40').setValue(zouGen);//請負
        ssD5.getRange('AX40').setValue(zouGen);//請負
        ssD6.getRange('AX40').setValue(zouGen);//請負
      }else{
        //請求書に精算増減の事由
        ss1.getRange('AF37').setValue("");

        ssD1.getRange('AX38').setValue("");//請負
        ssD2.getRange('AX38').setValue("");//請負
        ssD3.getRange('AX38').setValue("");//請負
        ssD4.getRange('AX38').setValue("");//請負
        ssD5.getRange('AX38').setValue("");//請負
        ssD6.getRange('AX38').setValue("");//請負

        ssD1.getRange('AX40').setValue("");//請負
        ssD2.getRange('AX40').setValue("");//請負
        ssD3.getRange('AX40').setValue("");//請負
        ssD4.getRange('AX40').setValue("");//請負
        ssD5.getRange('AX40').setValue("");//請負
        ssD6.getRange('AX40').setValue("");//請負
      }

      ssD1.getRange('AV31').setValue(keiyaku);

      if(henSt !== "なし"){//工期変更
        //請求書の工期も変更する必要がある
        const henKouki = henSt + "～" + henFn;
        ss1.getRange('B31').setValue(henKouki);
        ss1B.getRange('B31').setValue(henKouki);

        ssD1.getRange('E25').setValue("レ");
        ssD1.getRange('AE41').setValue(motoSt);
        ssD1.getRange('AU41').setValue(motoFn);
        ssD1.getRange('AE43').setValue(henSt);
        ssD1.getRange('AU43').setValue(henFn);

        ssD2.getRange('E25').setValue("レ");
        ssD2.getRange('AE41').setValue(motoSt);
        ssD2.getRange('AU41').setValue(motoFn);
        ssD2.getRange('AE43').setValue(henSt);
        ssD2.getRange('AU43').setValue(henFn);

        ssD3.getRange('E25').setValue("レ");
        ssD3.getRange('AE41').setValue(motoSt);
        ssD3.getRange('AU41').setValue(motoFn);
        ssD3.getRange('AE43').setValue(henSt);
        ssD3.getRange('AU43').setValue(henFn);

        ssD4.getRange('E25').setValue("レ");
        ssD4.getRange('AE41').setValue(motoSt);
        ssD4.getRange('AU41').setValue(motoFn);
        ssD4.getRange('AE43').setValue(henSt);
        ssD4.getRange('AU43').setValue(henFn);

        ssD5.getRange('E25').setValue("レ");
        ssD5.getRange('AE41').setValue(motoSt);
        ssD5.getRange('AU41').setValue(motoFn);
        ssD5.getRange('AE43').setValue(henSt);
        ssD5.getRange('AU43').setValue(henFn);

        ssD6.getRange('E25').setValue("レ");
        ssD6.getRange('AE41').setValue(motoSt);
        ssD6.getRange('AU41').setValue(motoFn);
        ssD6.getRange('AE43').setValue(henSt);
        ssD6.getRange('AU43').setValue(henFn);
      }else{//工期変更ない
        ssD1.getRange('E25').setValue("");
        ssD1.getRange('AE41').setValue("");
        ssD1.getRange('AU41').setValue("");
        ssD1.getRange('AE43').setValue("");
        ssD1.getRange('AU43').setValue("");

        ssD2.getRange('E25').setValue("");
        ssD2.getRange('AE41').setValue("");
        ssD2.getRange('AU41').setValue("");
        ssD2.getRange('AE43').setValue("");
        ssD2.getRange('AU43').setValue("");

        ssD3.getRange('E25').setValue("");
        ssD3.getRange('AE41').setValue("");
        ssD3.getRange('AU41').setValue("");
        ssD3.getRange('AE43').setValue("");
        ssD3.getRange('AU43').setValue("");

        ssD4.getRange('E25').setValue("");
        ssD4.getRange('AE41').setValue("");
        ssD4.getRange('AU41').setValue("");
        ssD4.getRange('AE43').setValue("");
        ssD4.getRange('AU43').setValue("");

        ssD5.getRange('E25').setValue("");
        ssD5.getRange('AE41').setValue("");
        ssD5.getRange('AU41').setValue("");
        ssD5.getRange('AE43').setValue("");
        ssD5.getRange('AU43').setValue("");

        ssD6.getRange('E25').setValue("");
        ssD6.getRange('AE41').setValue("");
        ssD6.getRange('AU41').setValue("");
        ssD6.getRange('AE43').setValue("");
        ssD6.getRange('AU43').setValue("");
      }

      SpreadsheetApp.flush();
      //請求フォルダ作成
      const fldId = 請求書保存先(tukiS,genmei);//IDを返す
      Logger.log(fldId);
      const pFld = DriveApp.getFolderById(fldId);
      const fn = "【請求書】" + genmei + "(工事)大成有楽不動産" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss") + ".pdf"; 
      const pName1 = "ﾒｰﾙ" + fn;//印ありメール用
      const pName2 = "ﾌﾟﾘﾝﾄ" + fn;//印なしプリント用
      //パターン１請求書だけ,2請求書と報告書,3請求書と報告書とD様式（増、減、工期だけの3パターン）
      const patternF = () => {
        if(firstEnd == "1回のみ" || end == true){
          if(zoGen == "なし"){
            if(henSt == "なし"){
              return "請求報";//
            }else{
              return "請求報工";
            }
          }else if(zoGen == "増"){
            return "請求報増";
          }else if(zoGen == "減"){
            return "請求報減";
          }
        }else if(end == false){
          if(zoGen == "なし"){
            if(henSt == "なし"){
              return "請求";//
            }else{
              return "請求工";
            }
          }else if(zoGen == "増"){
            return "請求増";
          }else if(zoGen == "減"){
            return "請求減";
          }
        }
      }
      const pattern = patternF();//パターンによって選択シートを変える（請求、請求増、請求減、請求工、請求報、請求報増、請求報減、請求報工）
      const att1 = createPdfT(pattern,pName1,pFld,ke,manBuild);//メール用
      const att2 = createPdfT(pattern,pName2,pFld,ke,manBuild);//プリント用

      mp = "ﾒｰﾙ";
      const cb1 = taiseiMaiSend(mp,att1,moto,manBuild,mailM,tanM,kouzi,tukiS,ke)
      mp = "ﾌﾟﾘﾝﾄ";
      const cb2 = taiseiMaiSend(mp,att2,moto,manBuild,mailM,tanM,kouzi,tukiS,ke)
      const systemMail = 管理Mail("システム");
      if(cb1 !== "" && cb2 !== ""){
        //成功チャット送信
        const msg = "大成請求書発行メール生成に成功しました\n請求書：" + fn + "\n" + systemMail + "のメール下書きを用意しました\n" + cb1;
        const url = chatUrl("管理1");//丸山
        sendChat(url,msg);
        //😃😃請求書スペース送信
        const cleanUke = String(seikyu).replace(/[\n,]/g, "");
        const ts = "請求";
        const kFileId = "";
        tyumonSpace(kFileId, cleanUke, genmei, kousyu, ts, cb2);//注文書をスペースにあげる
      }else{
        //失敗チャット送信
        const msg = "※失敗※大成請求書発行メール生成失敗\n請求書：" + fn + "\n確認ください";
        const url = chatUrl("管理1");//丸山
        sendChat(url,msg);
      }
      
    }else if(ke == 2){
      Logger.log("🔷keは2です");
      const hId2 = ひな形("大成役務請求");
      const sss = SpreadsheetApp.openById(hId2);//役務ひな形
      const ss1 = sss.getSheetByName('様式A（役務）');
      const ss1B = sss.getSheetByName('印なし１');

      const ss2 = sss.getSheetByName('様式B（内訳）');
      const ss2B = sss.getSheetByName('内訳控え');
      const ss3 = sss.getSheetByName('作業完了報告書');
      const ss3B = sss.getSheetByName('印なし３');
      
      ss1.getRange('AP29').setValue(seikyu);
      ss1B.getRange('AP29').setValue(seikyu);

      //日付
      ss1.getRange('BA2').setValue(seibi);//日付請求日
      ss1B.getRange('BA2').setValue(seibi);//日付請求日
      ss2.getRange('G1').setValue(seibi);//日付請求日
      ss2B.getRange('G1').setValue(seibi);//日付請求日
      ss3.getRange('I1').setValue(seibi);//日付請求日
      ss3B.getRange('I1').setValue(seibi);//日付請求日
      
      ss1.getRange('K29').setValue(kouzi);//物件＋内容
      ss1B.getRange('K29').setValue(kouzi);//物件＋内容
      

      //工期
      const finishA = "～" + finish;
      ss1.getRange('D29').setValue(start);
      ss1B.getRange('D29').setValue(start);
      ss1.getRange('D31').setValue(finishA);
      ss1B.getRange('D31').setValue(finishA);

      ss2.getRange('C7:H7').clearContent();
      ss2B.getRange('C7:H7').clearContent();
      ss2.getRange('C8').clearContent();
      ss2B.getRange('C8').clearContent();

      ss2.getRange('C7').setValue(start);//着工
      ss2B.getRange('C7').setValue(start);//着工
      ss2.getRange('D7').setValue(kouzi);//物件＋内容
      ss2B.getRange('D7').setValue(kouzi);//物件＋内容
      
      ss2.getRange('C8').setValue(finishA);//終了
      ss2B.getRange('C8').setValue(finishA);//終了

      const utiwake = ss2.getRange('C9:H31').getValues();//内訳を控えの方にコピペ
      Logger.log("utiwake；" + utiwake);
      ss2B.getRange('C9:H31').setValues(utiwake);

      ss3.getRange('D13').setValue(genmei);//物件名
      ss3B.getRange('D13').setValue(genmei);//物件名
      ss3.getRange('D14').setValue(kousyu);//作業名
      ss3B.getRange('D14').setValue(kousyu) ;//作業名
      ss3.getRange('F15').setValue(start);//着工
      ss3B.getRange('F15').setValue(start) ;//着工
      ss3.getRange('F16').setValue(finish);//終了
      ss3B.getRange('F16').setValue(finish) ;//終了
      SpreadsheetApp.flush();

      //請求フォルダ作成
      const fldId = 請求書保存先(tukiS,genmei);//IDを返す
      const pFld = DriveApp.getFolderById(fldId);
      const fn = "【請求書】" + genmei + "(役務)大成有楽不動産" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss") + ".pdf"; 
      const pName1 = "ﾒｰﾙ" + fn;//印ありメール用
      const pName2 = "ﾌﾟﾘﾝﾄ" + fn;//印なしプリント
      const pattern = "";
      const att1 = createPdfT(pattern,pName1,pFld,ke,manBuild);//メール用
      const att2 = createPdfT(pattern,pName2,pFld,ke,manBuild);//プリント用

      mp = "ﾒｰﾙ";
      const cb1 = taiseiMaiSend(mp,att1,moto,manBuild,mailM,tanM,kouzi,tukiS,ke)
      mp = "ﾌﾟﾘﾝﾄ";
      const cb2 = taiseiMaiSend(mp,att2,moto,manBuild,mailM,tanM,kouzi,tukiS,ke)
      const systemMail = 管理Mail("システム");
      if(cb1 !== "" && cb2 !== ""){
        //成功チャット送信
        const msg = "大成請求書発行メール生成に成功しました\n請求書：" + fn + "\n" + systemMail + "のメール下書きを用意しました\n" + cb2;
        const url = chatUrl("管理1");//丸山
        sendChat(url,msg);

        //😃😃請求書スペース送信
        const cleanUke = String(seikyu).replace(/[\n,]/g, "");
        const ts = "請求";
        const kFileId = "";
        tyumonSpace(kFileId, cleanUke, genmei, kousyu, ts, cb2);//注文書をスペースにあげる

      }else{
        //失敗チャット送信
        const msg = "※失敗※大成請求書発行メール生成失敗\n請求書：" + fn + "\n確認ください";
        const url = chatUrl("管理1");//丸山
        sendChat(url,msg);
      }
    }
  }

  //◆◆PDF作成◆◆
function createPdfT(pattern,pName,pFld,ke,manBuild){
  //必要なシートをコピーして
  Logger.log(pattern + "," + pName + "," + pFld + "," + ke);
  if(ke == 1){  
    const fileId = ひな形("大成工事請求");//ひな形フォルダ
    const file = DriveApp.getFileById(fileId);
    const newFileId = file.makeCopy(pName,pFld).getId();
    const sss = SpreadsheetApp.openById(newFileId);
    const ss1 = sss.getSheetByName('様式C（工事）');
    const ss1B = sss.getSheetByName('印なし１');
    const ss2 = sss.getSheetByName('工事検査願（工事）');
    const ss2B = sss.getSheetByName('印なし２');
    const ssD1 = sss.getSheetByName('様式D（工事）');
    const ssD2 = sss.getSheetByName('様式D（工事）減');
    const ssD3 = sss.getSheetByName('様式D（工事）工期のみ');
    const ssD4 = sss.getSheetByName('印なし３増');
    const ssD5 = sss.getSheetByName('印なし３減');
    const ssD6 = sss.getSheetByName('印なし工期のみ');
    const kasira = pName.slice(0,1);
    Logger.log(kasira);
    if(kasira == "ﾒ"){
      //いらないシートを削除（ﾒｰﾙ用）
      if(pattern == "請求"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求増"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        // sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求減"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        // sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求工"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        // sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        // sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報増"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        // sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        // sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報減"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        // sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        // sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報工"){
        // sss.deleteSheet(ss1);
        sss.deleteSheet(ss1B);
        // sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        // sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }
    }else if(kasira == "ﾌ"){
      //いらないシートを削除（ﾌﾟﾘﾝﾄ用)
      if(pattern == "請求"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求増"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        // sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求減"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        // sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求工"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        // sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        // sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報増"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        // sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        // sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報減"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        // sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        // sss.deleteSheet(ssD5);
        sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }else if(pattern == "請求報工"){
        sss.deleteSheet(ss1);
        // sss.deleteSheet(ss1B);
        sss.deleteSheet(ss2);
        // sss.deleteSheet(ss2B);
        sss.deleteSheet(ssD1);
        sss.deleteSheet(ssD2);
        sss.deleteSheet(ssD3);
        sss.deleteSheet(ssD4);
        sss.deleteSheet(ssD5);
        // sss.deleteSheet(ssD6);
        let pdf = sss.getAs('application/pdf');
        pdf.setName(pName);
        pId = pFld.createFile(pdf).getId();
        return pId;
      }
    }
  }else if(ke == 2){
    const fileId = ひな形("大成役務請求");//ひな形フォルダ
    const file = DriveApp.getFileById(fileId);
    const newFileId = file.makeCopy(pName,pFld).getId();
    const sss = SpreadsheetApp.openById(newFileId);
    const ss1 = sss.getSheetByName('様式A（役務）');
    const ss1B = sss.getSheetByName('印なし１');
    const ss2 = sss.getSheetByName('様式B（内訳）');
    const ss2B = sss.getSheetByName('内訳控え');
    const ss3 = sss.getSheetByName('作業完了報告書');
    const ss3B = sss.getSheetByName('印なし３');
    const kasira = pName.slice(0,1);
    if(kasira == "ﾒ" && manBuild !== "千葉"){
      //sss.deleteSheet(ss1);
      sss.deleteSheet(ss1B);
      //sss.deleteSheet(ss2);
      sss.deleteSheet(ss2B);
      // sss.deleteSheet(ss3);
      sss.deleteSheet(ss3B);
      let pdf = sss.getAs('application/pdf');
      pdf.setName(pName);
      pId = pFld.createFile(pdf).getId();
      return pId;
    }else if(kasira == "ﾌ" && manBuild !== "千葉"){
      sss.deleteSheet(ss1);
      //sss.deleteSheet(ss1B);
      //sss.deleteSheet(ss2);
      sss.deleteSheet(ss2B);
      sss.deleteSheet(ss3);
      // sss.deleteSheet(ss3B);
      let pdf = sss.getAs('application/pdf');
      pdf.setName(pName);
      pId = pFld.createFile(pdf).getId();
      return pId;
    }else if(kasira == "ﾒ" && manBuild == "千葉"){
      //sss.deleteSheet(ss1);
      sss.deleteSheet(ss1B);
      //sss.deleteSheet(ss2);
      sss.deleteSheet(ss2B);
      sss.deleteSheet(ss3);
      sss.deleteSheet(ss3B);
      let pdf = sss.getAs('application/pdf');
      pdf.setName(pName);
      pId = pFld.createFile(pdf).getId();
      return pId;
    }else if(kasira == "ﾌ" && manBuild == "千葉"){
      sss.deleteSheet(ss1);
      //sss.deleteSheet(ss1B);
      //sss.deleteSheet(ss2);
      sss.deleteSheet(ss2B);
      sss.deleteSheet(ss3);
      sss.deleteSheet(ss3B);
      let pdf = sss.getAs('application/pdf');
      pdf.setName(pName);
      pId = pFld.createFile(pdf).getId();
      return pId;
    }
  }
}

//◆◆請求書保存先フォルダ作成◆◆
function 請求書保存先(tukiS,genmei){
  Logger.log("請求保存先" + "," + tukiS + "," + genmei);
  const nentuki = 本日の年度();
  Logger.log("請求書保存先のtukiS " + nentuki[0] + "," + nentuki[1]);
  const kongetu = nentuki[1];
  const tuki = Number(tukiS);
  Logger.log("請求書保存先のtukiS " + tuki);
  const nendoF = () => {
    if(tuki == kongetu){
      const n = nentuki[0];
      return n;
    }else if(kongetu > 9 && tuki < 10){
      const n = nentuki[0] - 1;
      return n;
    }else{
      const n = nentuki[0];
      return n; 
    }
  }
  const nendo = nendoF();
  Logger.log(nendo);
  const idN = getSPId("年度フォルダ") 
  const sssN = SpreadsheetApp.openById(idN);//年度ファイル
  const ssN = sssN.getSheetByName('月フォルダ管理');
  const last = ssN.getRange('R100').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const end = last + 1
  for(let i=2; i<last+2; i++){
    if(i == end){
      //年度ファルダもない状態
      const oyaId = ssN.getRange('S2').getValue();
      const oya = DriveApp.getFileById(oyaId); 
      const nameN = nendo + "年";
      const fdN = oya.createFolder(nameN);
      const fdNId = fdN.getId();
      ssN.getRange('R' + end).setValue(nendo);
      ssN.getRange('T' + end).setValue(fdNId);
      const nameT = tuki + "月";//年度フォルダがないので月フォルダもない
      const fdTId =  fdN.createFolder(nameT).getId();
      const col = tuki + 20;
      ssN.getRange(end,col).setValue(fdTId);
      const fdT = DriveApp.getFolderById(fdTId);
      const id = fdT.createFolder(genmei).getId();
      return id;
    }else if(ssN.getRange('R' + i).getValue() == nendo){
      const col = tuki + 20;
      Logger.log(col);
      const tId = ssN.getRange(i,col).getValue();
      Logger.log(tId);
      if(tId == ""){
        const nId = ssN.getRange('T' + i).getValue();
        const fdN = DriveApp.getFolderById(nId);
        const nameT = tuki + "月";
        const fdTId = fdN.createFolder(nameT).getId();
        ssN.getRange(i,col).setValue(fdTId);
        const fdT = DriveApp.getFolderById(fdTId);
        const id = fdT.createFolder(genmei).getId();
        return id;
      }else{
        const tFld = DriveApp.getFolderById(tId);
        const id = tFld.createFolder(genmei).getId();//現場名のフォルダを生成しID取得
        Logger.log(id);
        return id;
      }
    }
  }
}

//◆◆◆請求書メール作成下書き保存する◆◆◆※メール送信はsystem.ishii@ebisu-ishii.co.jpなので。。どうするか
function taiseiMaiSend(mp,att,moto,manBuild,mailM,tanM,kouzi,tukiS,ke){;//相手の担当者へメールする (mpはﾒｰﾙかﾌﾟﾘﾝﾄか,pdf2は添付するPDF,元請名,マンションかビルか千葉,担当者ﾒｰﾙ,担当者名（苗字スペース名前)） ※※attはファイルID
  Logger.log(mp + "," + att + "," + moto + "," + manBuild + "," + mailM + "," + tanM + "," + kouzi + "," + tukiS + "," + ke)
  //mailM = "test@anemoworks.com";
  Logger.log("pdfのid " + att);
  const pdf = DriveApp.getFileById(att);
  const pUrl = pdf.getUrl(); 
  Logger.log(pUrl);
  if(mp == "ﾒｰﾙ"){//元請担当者宛

    //◆◇◆◇準備が出来たら外す◆◇◆◇準備が出来たら外す◆◇◆◇準備が出来たら外す◆◇◆◇
    const mailSF = () => {//CC請求書担当事務のメールを選択
      if(manBuild == "マンション"){
        return "fukumoto.k@taisei-yuraku.co.jp";//福元様
      }else if(manBuild == "ビル"){
        return "yoshimi.r@taisei-yuraku.co.jp";//能見様
      }else if(manBuild == "千葉"){
        return "ueuchi.t@taisei-yuraku.co.jp"//上内様
      }
    } 
    const mailS = mailSF();

    const mailS2 = 管理Mail("社長");
    
    // const mailS = "test@anemoworks.com";
    // const mailS2 = "test2@anemoworks.com";
    //◆◇◆◇準備が出来たら外す◆◇◆◇準備が出来たら外す◆◇◆◇準備が出来たら外す◆◇◆◇

    const tanF = () => {
      name = tanM.slice(0,tanM.indexOf("　"));
      return name;
    }
    const tan = tanF();
    const bodyF = () => {
      const syomei = メール署名();
      if(manBuild == "マンション"){
        msg = moto + "\n" + tan + "様　" + "CC 福元様\n\nいつもお世話になっております\n" + kouzi + "の請求書PDFを発行しました(" + tukiS + "月請求)ご確認お願いいたします\n" +  "もし不備がありましたらお申し付けください\n\n" + syomei;
        return msg;//福元様
      }else if(manBuild == "ビル"){
        msg = moto + "\n" + tan + "様　" + "CC 能見様\n\nいつもお世話になっております\n" + kouzi + "の請求書PDFを発行しました(" + tukiS + "月請求)ご確認お願いいたします\n" +  "もし不備がありましたらお申し付けください\n\n" + syomei;
        return msg;//能見様
      }else if(manBuild == "千葉" && ke == 1){//ke:1は工事　2は役務
        msg = moto + "\n" + tan + "様　" + "CC 上内様\n\nいつもお世話になっております\n" + kouzi + "の請求書PDFを発行しました(" + tukiS + "月請求)ご確認お願いいたします\n" +  "もし不備がありましたらお申し付けください\n\n" + syomei;
        return msg;//上内様
      }else if(manBuild == "千葉" && ke == 2){//ke:1は工事　2は役務
        msg = moto + "\n" + tan + "様　" + "CC 上内様\n\nいつもお世話になっております\n" + kouzi + "の請求書PDFを発行しました(" + tukiS + "月請求)ご確認お願いいたします\n" +  "もし不備がありましたらお申し付けください\n作業完了確認書（4枚綴り）は別送信いたします\n\n" + syomei;
        return msg;//上内様
      }
    }
    const body = bodyF();

    const sub = kouzi + "　請求書(" + tukiS + "月)"
    const cc = mailS + "," + mailS2;
    const option = {
      "cc":cc,
      "attachments":pdf
    }
    GmailApp.createDraft(mailM,sub,body,option);

    return pUrl;

  }else if(mp == "ﾌﾟﾘﾝﾄ"){//事務員宛
  
    ////◆◇準備が出来たら外すやつ◆◇準備が出来たら外すやつ◆◇準備が出来たら外すやつ
    const zimu = 管理Mail("事務");
    // const zimu = "test@anemoworks.com";
    //◆◇準備が出来たら外すやつ◆◇準備が出来たら外すやつ◆◇準備が出来たら外すやつ

    const sub2 = kouzi + " 請求書（" +　tukiS + "月）";
    const bodyF2 = () => {
      if(manBuild == "千葉" && ke == 2){
        msg = sub2 + "\n" + moto + " 担当者：" + tanM + "さん\n\n大成千葉支店　作業完了確認書(4枚綴り)の送信（PDF）のみです"
        return msg;
      }else{
        msg = sub2 + "\n" + moto + "  担当者：" + tanM + "さん\n\nプリント＆捺印＆郵送をお願いします";
        return msg;
      }
    }
    const body2 = bodyF2();
    const option2 = {
      'attachments':pdf
    }
    Logger.log(zimu + "," + sub2 + "," + body2 + "," + option2);
    if(manBuild == "千葉" && ke == 1){
      return pUrl;
    }else{
      GmailApp.createDraft(zimu,sub2,body2,option2);
      return pUrl;
    }
  }
}
//◆◆大成請求書発行時未入力の項目の記録を反映させる
function 入力情報反映(fId,ds,tanM,mailM,ke,start,finish,add,henSt,henFn,keiyaku){
  const sss = SpreadsheetApp.openById(fId);
  const ss = sss.getSheetByName('シート1');
  if(ds == "D"){
    ss.getRange('AF102').setValue(tanM);//元請担当
    ss.getRange('AG102').setValue(mailM);//元請メール
    ss.getRange('AH102').setValue(ke);//工事/役務

    if(henSt == "なし"){//工期変更があった場合
      ss.getRange('R102').setValue(start);//着工
      ss.getRange('S102').setValue(finish);//完了
    }else{
      ss.getRange('R102').setValue(henSt);//着工
      ss.getRange('S102').setValue(henFn);//完了
    }
    
    ss.getRange('T102').setValue(add);//住所
    ss.getRange('AJ102').setValue(keiyaku);//契約日

    const nenRow = String(ss.getRange('A1').getValue());
    Logger.log(nenRow);
    const nendo = Number(nenRow.slice(0,4));
    const row = Number(nenRow.slice(4));
    const nId = getSPId("年度フォルダ");
    const sssN = SpreadsheetApp.openById(nId);//年度ファイル
    const ssN = sssN.getSheetByName('生成履歴');
    const last = ssN.getRange('B1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    for(let i=2; i<last+1;i++){
      if(ssN.getRange('A' + i).getValue() == nendo){
        const dId = ssN.getRange('F' + i).getValue();
        const sssD = SpreadsheetApp.openById(dId);
        const ssD = sssD.getSheetByName('工事リスト');
        vals = ss.getRange('A102:AJ102').getValues();
        ssD.getRange('A' + row + ":AJ" + row).setValues(vals);
      }
    }
  }else if(ds == "S"){
    ss.getRange('AF19').setValue(tanM);//元請担当
    ss.getRange('AG19').setValue(mailM);//元請メール
    ss.getRange('AH19').setValue(ke);//工事/役務

    ss.getRange('R19').setValue(start);//着工
    ss.getRange('S19').setValue(finish);//完了

    ss.getRange('T19').setValue(add);//住所
    ss.getRange('AJ19').setValue(keiyaku);//契約日

    const nenRow = String(ss.getRange('A1').getValue());
    const nendo = Number(nenRow.slice(0,4));
    const row = Number(nenRow.slice(4));
    const nId = getSPId("年度フォルダ");
    const sssN = SpreadsheetApp.openById(nId);//年度ファイル
    const ssN = sssN.getSheetByName('生成履歴');
    const last = ssN.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    for(let i=2; i<last+1;i++){
      if(ssN.getRange('A' + i).getValue() == nendo){
        const dId = ssN.getRange('F' + i).getValue();
        const sssD = SpreadsheetApp.openById(dId);
        const ssD = sssD.getSheetByName('工事リスト');
        vals = ss.getRange('A19:AJ19').getValues();
        ssD.getRange('A' + row + ":AJ" + row).setValues(vals);
      }
    }
  }
}

//長谷工工事完了報告書発行
function reziCsvGet(gen,reziNo){
  Logger.log(gen + " & " + reziNo);
  const id = getSPId("長谷工完了報告");
  const sss = SpreadsheetApp.openById(id);
  const ss = sss.getSheetByName('発注管理表');
  const last = ss.getRange('W1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i=2; i<last+1; i++){
    if(ss.getRange('W' + i).getValue() == reziNo){//登録の現場名とCSVが違う可能性があるので
      const hatyuNo = reziNo;
      const kousyu = ss.getRange('C' + i).getValue();
      const genmei = ss.getRange('B' + i).getValue();
      const add = ss.getRange('BB' + i).getValue();
      const startA = ss.getRange('BE' + i).getValue();
      const start = Utilities.formatDate(startA,"JST","yyyy年M月d日");
      const finishA = ss.getRange('BF' + i).getValue();
      const finish = Utilities.formatDate(finishA,"JST","yyyy年M月d日");
      const ukeoi = ss.getRange('BJ' + i).getValue();//税込み
      const keiyakuA = ss.getRange('BP' + i).getValue();
      const keiyaku = Utilities.formatDate(keiyakuA,"JST","yyyy年M月d日");
      //担当
      const bikou = ss.getRange('BO' + i).getValue();
      const tanA = bikou.slice(bikou.lastIndexOf('：') + 1);
      const tan = tanA.replace(/\r?\n/g, "");
      const mailF = () => {
        const ss2 = sss.getSheetByName('メールアドレス');
        const last2 = ss2.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
        const end = last2 + 1
        for(let s=2; s<last2+2; s++){
          t = ss2.getRange('A' + s).getValue()
          if(s == end){
            m = "なし";//メールなし
            return m;
          }else if(t.indexOf(tan) !== -1){
            m = ss2.getRange('C' + s).getValue();
            return m;
          }
        }
      }
      const mail = mailF();
      const tel = ss.getRange('BD' + i).getValue();
      //num + "," + gen + "," + kinds + "," + add + "," + start + "," + finish + "," + kin + "," + keiyaku + "," + tan + "," + mail
      const vals = hatyuNo + "," + genmei + "," + kousyu + "," + add + "," + start + "," + finish + "," + ukeoi + "," + keiyaku + "," + tan + "," + mail + "," + tel; 
      Logger.log(vals);
      return vals;
    }
  }
  Logger.log("Noがない");
}
//請求書保存先フォルダ取得
function 請求書保存月フォルダ(tukiS){
    const nentuki = 本日の年度();
    const kongetu = nentuki[1];
    const nendoF = () => {
      if(tukiS == kongetu){
        return nentuki[0];
      }else if(kongetu > 9 && tukiS < 10){
        n = nentuki[0] - 1;
        return n;
      }else{
        return nentuki[0];
      }
    }
    const nendo = nendoF();
    Logger.log("nendo " + nendo);
    const nId = getSPId("年度フォルダ");
    const sssN = SpreadsheetApp.openById(nId);//年度ファイル
    const ssN = sssN.getSheetByName('月フォルダ管理');
    const last = ssN.getRange('R100').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const end = last + 1;
    for(let i=2; i<last+2; i++){
      if(i == end){
        //年度ファイルがない
        const oyaId = ssN.getRange('S2').getValue();
        const oya = DriveApp.getFolderById(oyaId);
        const nameN = nendo + "年";
        const fdN = oya.createFolder(nameN);
        const fdNId = fdN.getId();
        ssN.getRange('R' + end).setValue(nendo);
        ssN.getRange('T' + end).setValue(fdNId);
        const nameT = tukiS + "月";
        const fdT = fdN.createFolder(nameT);
        const fdTId = fdT.getId();
        const col = Number(tukiS) + 20;
        ssN.getRange(end,col).setValue(fdTId);
        const tFld = DriveApp.getFolderById(fdTId);
        return tFld;
      }else if(ssN.getRange('R' + i).getValue() == nendo){
        const col = Number(tukiS) + 20;
        Logger.log("col " + col);
        const tId = ssN.getRange(i,col).getValue();
        Logger.log("tId " + tId);
        if(tId == ""){//ない場合月フォルダを作成してIDを返す
          const nameT = tukiS + "月";
          const nId = ssN.getRange('T' + i).getValue();
          const nenFld = DriveApp.getFolderById(nId);
          const tId = nenFld.createFolder(nameT).getId();
          ssN.getRange(i,col).setValue(tId);
          const tFld = DriveApp.getFolderById(tId);
          return tFld;
        }else{
          const tFld = DriveApp.getFolderById(tId);//月フォルダ
          return tFld;
        }
        
      }
    }
  }

//◆◆長谷工の工事完了報告のメール生成◆◆
function hasekoKanHo(hatyuNo,ukeoiS,genmei,kousyu,start,finish,keiyaku,add,tanM,mailM,tukiS,kanryo,tel){
  //保存フォルダfld　請求書フォルダに工事完了報告書（長谷工）を作成して入れる（もしすでにある場合を考慮）
  Logger.log(hatyuNo + "," + ukeoiS + "," + genmei + "," + kousyu + "," + start + "," + finish + "," + keiyaku + "," + add + "," + tanM + "," + mailM + "," + tukiS + "," + kanryo + "," + tel);
  
  const fldF = () =>{
    const tFld = 請求書保存月フォルダ(tukiS);
    //月フォルダに内にフォルダが既にあるか調べる
    fldName = "工事完了報告書（長谷工)";
    const folders = tFld.getFoldersByName(fldName);
    if (folders.hasNext()) {
        const folder = folders.next();
        const id = folder.getId();
        Logger.log(`フォルダID: ${id}`);
        return folder;
    } else {
        const newFolder = tFld.createFolder(fldName);
        const id = newFolder.getId();
        Logger.log(`作成したフォルダID: ${id}`);
        return newFolder;
    }
  } 
  const fld = fldF();

  const telAF = () => {
    if(!tel){
      const tTel = 長谷工完了報告TEL(hatyuNo);
      return tTel;
    }else{
      return tel;
    }
  }
  const telA = telAF(); 
  const fileId = getSPId("長谷工完了報告");//工事完了報告書ひな形フォルダ
  const file = DriveApp.getFileById(fileId);
  const fName = genmei + kousyu + "◆工事完了報告書" + Utilities.formatDate(new Date,"JST","yyyyMMdd");
  const newFileId = file.makeCopy(fName,fld).getId();
  const sss = SpreadsheetApp.openById(newFileId);//コピーを作成
  const ss = sss.getSheetByName('工事完了報告書');
  const ss2 = sss.getSheetByName('発注管理表');
  const ss3 = sss.getSheetByName('メールアドレス');
  const ss4 = sss.getSheetByName('請求貼付');
  
  ss.getRange('F16').setValue(hatyuNo);//発注No
  const kouzi = genmei + kousyu;
  ss.getRange('E15').setValue(kouzi);//工事名
  ss.getRange('F37').setValue(kouzi);//工事名
  ss.getRange('E22').setValue(start);//着工

  //完了日の判断、請求日と完了日が違う場合、完工日は請求日に合わせる
  // if(finish == kanryo){//一緒の場合
  //   ss.getRange('E23').setValue(finish);//完了  
  // }else{//違う場合
  ss.getRange('E23').setValue(finish);//完了 
  // }
  ss.getRange('E20').setValue(keiyaku);//契約日
  ss.getRange('I5').setValue(kanryo);//請求日
  ss.getRange('E27').setValue(kanryo);//請求日

  ss.getRange('E18').setValue(add);//住所
  ss.getRange('F38').setValue(tanM);//担当
  ss.getRange('E25').setValue(ukeoiS);//請負金額(込）
  ss.getRange('F39').setValue(telA);

  //いらないシートの削除
  sss.deleteSheet(ss2);
  sss.deleteSheet(ss3);
  sss.deleteSheet(ss4);

  //ﾌｧｲﾙ名
  const pName = genmei + kousyu + "◆工事完了報告書" + Utilities.formatDate(new Date,"JST","yyyyMMdd") + ".pdf";

  let pdf = sss.getAs('application/pdf');
  pdf.setName(pName);//名前を付ける
  const fUrl = fld.createFile(pdf).getUrl();//確認用にPDFのurlをチャットに送る用

  const tanN = tanM.replace(/\r?\n/g,"");//名前に改行があるので

  //メール下書き保存
  const sub = genmei + kousyu +"工事完了報告書"
  const syomei = メール署名();
  const body = "長谷工リフォーム " + tanN + "様\n\nいつもお世話になっております\n" + genmei + kousyu + "の工事完了報告書です\n発注書に合わせて作成しました。不備がございましたらお申し付けください\n\n" + syomei;
  const mailCC = 管理Mail("社長");
  const option = {
    'cc':mailCC,
    'attachments':pdf
  };
  GmailApp.createDraft(mailM,sub,body,option);

  //成功チャット送信
  const systemMail = 管理Mail("システム");
  const msg = "長谷工の工事完了報告書生成しました\n工事完了報告書：" + genmei + "\n" +systemMail + "のメール下書きを用意しました\n" + fUrl;
  const url = chatUrl("管理1");//丸山
  sendChat(url,msg);

  //工事完了報告書の発行の処理発行の日付記録
  const id = getSPId("長谷工完了報告");
  const sssKH = SpreadsheetApp.openById(id);
  const ssKH = sssKH.getSheetByName('発注管理表');
  const last = ssKH.getRange('W1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i=2; i<last+1; i++){
    if(ssKH.getRange('W' + i).getValue() == hatyuNo){
      const day = Utilities.formatDate(new Date,"JST","yyyy年M月d日");
      ssKH.getRange('E' + i).setValue(day);
      break;
    }
  }
}
//◆長谷工工事完了報告書のTELを取得
function 長谷工完了報告TEL(hatyuNo){
  const id = getSPId("長谷工完了報告");
  const sss = SpreadsheetApp.openById(id);//工事完了報告書
  const ss = sss.getSheetByName('発注管理表');
  const last = ss.getRange('W1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i=2; i<=last; i++){
    if(ss.getRange('W' + i).getValue() == hatyuNo){
      const tel = ss.getRange('BD' + i).getValue();
      return tel;
    }
  }
}

//◆◆自社ひな形の請求書の作成
function zisyaSei(cellInpt1,cellInpt2,genmei,kousyu,ukeoiS,seiS,tukiS,start,finish,seibi,add,moto,id,pay,tuika,tuikin,motoSct){
  Logger.log(cellInpt1 + "," + cellInpt2 + "," + genmei + "," + kousyu + "," + ukeoiS + "," + seiS + "," + tukiS + "," + start + "," + finish + "," + seibi + "," + add + "," + moto + "," + id + "," + pay + "tuika:" + tuika + " tuikin:" + tuikin + " motoSct:" + motoSct);
  const idH = ひな形("自社請求");
  const sssH = SpreadsheetApp.openById(idH);//自社物件用請求書ひな形★SS
  const ss1 = sssH.getSheetByName('請求書');
  const ss1B = sssH.getSheetByName('印なし');
  //請求番号
  const seiban = Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");
  ss1.getRange('R1').setValue(seiban);
  ss1B.getRange('R1').setValue(seiban);
  //請求日
  ss1.getRange('R2').setValue(seibi);
  ss1B.getRange('R2').setValue(seibi);
  //請求先住所
  Logger.log(moto);
  const addIdF = () => {
    const idK = getSPId("業者顧客担当");
    const sssK = SpreadsheetApp.openById(idK);//業者・顧客・担当者SS
    const ssK = sssK.getSheetByName('顧客台帳');
    const last = ssK.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    //Logger.log("last " + last);
    for(let i=2; i<last+1; i++){
      //Logger.log(ssK.getRange('B' + i).getValue() + " & " + moto);
      if(ssK.getRange('B' + i).getValue() == moto){
        sa = ssK.getRange('F' + i).getValue();
        kid = ssK.getRange('A' + i).getValue();
        post = ssK.getRange('E' + i).getValue();
        Logger.log(sa + "," + kid + "," + post);
        rn = [sa,kid,post];
        return rn; 
      }
    }
  }
  const addId = addIdF();
  Logger.log(addId);
  const seiAdd = addId[0];
  const code = addId[1];
  const postcode = addId[2];
  // const postcodeF = () => {
  //   p = addId[2];
  //   if(isNaN(p) == true){
  //     return p;
  //   }else{
  //     p2 = String(p);
  //     return p2;
  //   }
  // } 
  // const postcode = postcodeF();

  ss1.getRange('B4').setValue(seiAdd);
  ss1B.getRange('B4').setValue(seiAdd);
  //顧客コード
  ss1.getRange('C2').setValue(code);
  ss1B.getRange('C2').setValue(code);
  //請求先
  ss1.getRange('B5').setValue(moto);
  ss1B.getRange('B5').setValue(moto);
  //工事Id
  ss1.getRange('C12').setValue(id);
  ss1B.getRange('C12').setValue(id);
  //契約工期
  const kouki = start + "～" + finish;
  ss1.getRange('C13').setValue(kouki);
  ss1B.getRange('C13').setValue(kouki);
  //郵便番号
  const yubin = "〒" + postcode;
  ss1.getRange('B3').setValue(yubin);
  ss1B.getRange('B3').setValue(yubin);
  //工事ID
  ss1.getRange('C7').setValue(id);
  ss1B.getRange('C7').setValue(id);
  //支払い条件
  const paymentF = () => {
    if(pay == 1){
      val = "契約の通り";
      return val;
    }else if(pay == 2){
      val = "工事完了後一括払い";
      return val;
    }else if(pay == 3){
      val = "従来通り";
      return val;
    }else{
      return payment;
    }
  }
  const payment = paymentF();
  ss1.getRange('C12').setValue(payment);
  ss1B.getRange('C12').setValue(payment);
  //工事名
  const kouzi = genmei + kousyu;
  ss1.getRange('C8').setValue(kouzi);
  ss1B.getRange('C8').setValue(kouzi);
  ss1.getRange('B18').setValue(kouzi);
  ss1B.getRange('B18').setValue(kouzi);
  //現場住所
  ss1.getRange('C9').setValue(add);
  ss1B.getRange('C9').setValue(add);
  //セルB19の入力
  const celin1F = () => {
    if(cellInpt1 == 1){
      val = "工事一式";
      return val;
    }else if(cellInpt1 == 2){
      uke = Number(ukeoiS.replace(/,/g, ""));
      nuki = uke / 1.1
      nukiA = nuki.toLocaleString();
      val = "請負金額" + nukiA + "円";
      return val;
    }else{
      return cellInpt1;
    }
  }
  const celin1 = celin1F();
  ss1.getRange('B19').setValue(celin1);
  ss1B.getRange('B19').setValue(celin1);
  //E19の入力
  ss1.getRange('D19').setValue(cellInpt2);
  ss1B.getRange('D19').setValue(cellInpt2);

  //金額入力（税抜き）
  const seikinF = () => {
    const seiN = Number(seiS.replace(/,/g, ""));
    const nuki = seiN / 1.1;//税抜きにする
    return nuki;
  }
  const seikin = seikinF();
  
  ss1.getRange('K19').setValue(seikin);
  ss1B.getRange('K19').setValue(seikin);

  //追加項目ある場合
  ss1.getRange('B20:K20').clearContent();//一旦リセット
  ss1B.getRange('B20:K20').clearContent();//一旦リセット
  if(tuika !== "なし"){
    ss1.getRange('H20').setValue(1);
    ss1B.getRange('H20').setValue(1);
    ss1.getRange('I20').setValue("式")
    ss1B.getRange('I20').setValue("式")
    ss1.getRange('B20').setValue(tuika);
    ss1B.getRange('B20').setValue(tuika);
    ss1.getRange('K20').setValue(tuikin);
    ss1B.getRange('K20').setValue(tuikin);
    //本請求金額から追加項目金額を引く
    const tuikinA = Number(tuikin.replace(/,/g,""));
    const newsei = seikin - tuikinA;
    ss1.getRange('K19').setValue(newsei);
    ss1B.getRange('K19').setValue(newsei);
  }
  
  SpreadsheetApp.flush();

  const pPdf = genmei + "【請求書ﾌﾟﾘﾝﾄ用】" + tukiS + "月" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss") + ".pdf";//プリント用;
  const mPdf = genmei + "【請求書M】" + tukiS + "月" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss") + ".pdf";//メール用;

  //請求書PDF作成
  const flId = 請求書保存先(tukiS,genmei);//請求書保存フォルダIdを取得
  //保存先フォルダ内に現場名のフォルダを作成
  const saveFld = DriveApp.getFolderById(flId);
  
  fn = [pPdf,mPdf];//配列で最後のpdfファイルのurlを返す 
  const pdfUrl = 自社請求書PDF作成(fn,saveFld,genmei,tukiS);

  //😃😃請求書スペース送信
  const cleanUke = String(seikin).replace(/[\n,]/g, "");
  const ts = "請求";
  const kFileId = "";
  //tyumonSpaceはコードyosanRezi.jsにある
  tyumonSpace(kFileId, cleanUke, genmei, kousyu, ts, pdfUrl);//注文書をスペースにあげる

  if(motoSct == "billOne"){//長谷工のBillOneの場合を追加😃
    const fnK = genmei + "【完了報告書】" +  tukiS + "月" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");
    長谷工BillOne報告書作成(fnK,saveFld,genmei,kousyu,seibi,seikin,tukiS);//seikinはnuki(税抜き）
  }
  return pdfUrl;
}

//◆◆自社物件用請求書PDFを作成する
function 自社請求書PDF作成(fn,saveFld,genmei,tukiS){//fnは配列配列の要素の数だけPDF作成最後のPDFのURLを返す
  const hId = ひな形("自社請求");//自社物件用請求書ひな形
  const sss = SpreadsheetApp.openById(hId);
  const ss1 = sss.getSheetByName('請求書');
  const ss2 = sss.getSheetByName('印なし');
  const last = fn.length;
  const end = last - 1;
  let pdfId = "";
  for(let i=0; i<last; i++){
    let name = fn[i];
    if(name.indexOf("【請求書ﾌﾟﾘﾝﾄ用】") !== -1){
      pdfId1 = createPdf(sss,ss2,name,saveFld);
    }else if(name.indexOf("【請求書M】") !== -1){
      pdfId2 = createPdf(sss,ss1,name,saveFld);
    }
    if(i == end){
      //チャット送信
      const pdfFile = DriveApp.getFileById(pdfId1);//チャットに送るのはプリント用
      const url = pdfFile.getUrl();
      const msg = "◇自社物件用の請求書の作成が出来ました◇\n\n保存先：共有フォルダ⇒請求書⇒" + tukiS + "月\n⇒" + genmei + "\n\n確認できます\n" + url;
      const chat = chatUrl("管理1");//丸山
      sendChat(chat,msg);
      return url;
    }
  }
}


//😃長谷工のBillOne用の完了報告書作成
function 長谷工BillOne報告書作成(fnK, saveFld, genmei, kousyu, seibi, nuki, tukiS){
  const hId = ひな形("billOne報告書");
  const hFile = DriveApp.getFileById(hId);
  const flId = ドライブId("一時フォルダ丸山");
  const fld = DriveApp.getFolderById(flId);
  const id = hFile.makeCopy(fnK, fld).getId();
  const sss = SpreadsheetApp.openById(id);
  const ss = sss.getSheetByName("作業完了報告書");

  const zei  = Math.floor(nuki * 0.1); // 消費税を整数化
  const zeiA = zei.toLocaleString();
  const seikin = nuki + zei;
  const zeiMsg = "（内消費税" + zeiA + "円含む）";

  ss.getRange('M7').setValue(seibi);
  ss.getRange('E21').setValue(genmei);
  ss.getRange('E22').setValue(kousyu);
  ss.getRange('F24').setValue(seikin);
  ss.getRange('L24').setValue(zeiMsg);

  SpreadsheetApp.flush();
  // PDF作成
  const pdfName = fnK + ".pdf";
  const portrait = "&portrait=true";
  const pdfId = createPdf(sss, ss, pdfName, saveFld, portrait);
  const pdfFile = DriveApp.getFileById(pdfId);
  const url = pdfFile.getUrl();

  // Chat通知
  const chatMsg = "◇長谷工BillOne用の完了報告書◇\n\n保存先：共有フォルダ⇒請求書⇒" + tukiS + "月\n⇒" + genmei + "\n\n確認できます\n" + url;
  const chat = chatUrl("管理1");//丸山
  sendChat(chat, chatMsg);
}


//◆◆PDF作成◆◆
function createPdf(sssP,ssP,pdfName,saveFld,portrait){
  const fId = sssP.getId();
  const sId = ssP.getSheetId();
  Logger.log(fId + "," + sId)
  const portF = () => {
    if(!portrait){
      return '&portrait=false';
    }else{
      return portrait;
    }
  }
  const port = portF();
  const options = 'exportFormat=pdf&format=pdf'
    + '&gid=' + sId       //PDFにするシートの「シートID」
    + port      //true(縦) or false(横)
    + '&size=A4'        //印刷サイズ
    + '&gridlines=false'      //グリッドラインの表示有無
    //+ '&range=' + RANGE_F + '%3A' + RANGE_T   //セル範囲を指定 %3A はコロン(:)を表す
    + '&top_margin=0.50'      //上の余白
    + '&right_margin=0.50'    //右の余白
    + '&bottom_margin=0.20'   //下の余白
    + '&left_margin=0.50'     //左の余白
    + '&sheetnames=false'     //シート名の表示有無
    + '&printtitle=false'     //スプレッドシート名の表示有無
    + '&pagenum=UNDEFINED'    //ページ番号をどこに入れるか
    + '&scale=2'              //1= 標準100%, 2= 幅に合わせる, 3= 高さに合わせる,  4= ページに合わせる
    + '&horizontal_alignment=CENTER'//水平方向の位置
    + '&vertical_alignment=CENTER'//垂直方向の位置
    + '&gridlines=false'      //グリッドラインの表示有無
    + '&fzr=false'            //固定行の表示有無
    + '&fzc=false'            //固定列の表示有無
    const url = "https://docs.google.com/spreadsheets/d/" + fId + "/export?" + options;

  Logger.log(url);
  let token = ScriptApp.getOAuthToken();
  let responce = UrlFetchApp.fetch(url,{headers:{
    "Authorization" : "Bearer " + token
  }
  });
  let blob = responce.getBlob().setName(pdfName);
  const id = saveFld.createFile(blob).getId();
  return id;
}

//◆◆現在の担当者出来高SSのIDと現場リストを取得する
function getTanDekiGenList(){
  const idN = getSPId("年度フォルダ");
  const sssN = SpreadsheetApp.openById(idN);
  const ssN = sssN.getSheetByName('生成履歴');
  const tdId = ssN.getRange('G2').getValue();
  const sssF = SpreadsheetApp.openById(tdId);
  const ssF = sssF.getSheetByName('シート1');
  const last = ssF.getRange('Y100').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const val = ssF.getRange('AH1:AH' + last).getValues();
  const vals = tdId + "," + last + "," + val;
  return vals;
}
//◆◆担当者出来高SSの現場受け取り、出来高登録用URLを返す
function getTandekiUrl(tandekiFId,sctDekiGen,gen){
  const sss = SpreadsheetApp.openById(tandekiFId);
  const ss = sss.getSheetByName('シート1');
  const row = Number(sctDekiGen) + 1;
  const genPass = ss.getRange('AB' + row).getValue();
  const r1 = ss.getRange('AC' + row).getValue();
  const r2 = ss.getRange('AD' + row).getValue();
  const urlA = getSetUrl("担当者出来高");
  const url =  urlA + "?page=index&param1=" + genPass + "&param2=" + r1 + "&param3=" + r2;
  return url;
}

//◆◆◆自社物件の注文請書の作成⇒チャットに送信
function tyumonUke(day,id,gen,naiyo,uke,terms,start,finish,client,syomei,ds,add,fId,keiyakuSct){
  Logger.log(day + "," + id + "," + gen + "," + naiyo + "," + uke + "," + terms + "," + start + "," + finish + "," + client + "," + syomei + "," + ds + "," + add);
  const idH = ひな形("石井請負注文");
  const file = DriveApp.getFileById(idH);
  const fn = "注文請書ひな形コピー" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");
  const flId = ドライブId("一時フォルダ丸山");
  const fl1 = DriveApp.getFolderById(flId);//一時フォルダ
  const newFileId = file.makeCopy(fn,fl1).getId(); 
  const sss = SpreadsheetApp.openById(newFileId);//注文請書ひな形
  const ss = sss.getSheetByName('表紙');
  const ss2 = sss.getSheetByName('請書表紙');
  const info = 顧客情報(client);
  const postC = info[0];
  const addC = info[1];
  const genmei = gen + naiyo
  const ukeA = Number(uke.replace(/,/g, ""));
  ss.getRange('M4').setValue(day);
  ss.getRange('D13').setValue(id);
  ss.getRange('D14').setValue(genmei);
  ss.getRange('D15').setValue(add);
  ss.getRange('J7').setValue("〒" + postC);
  ss.getRange('J8').setValue(addC);
  ss.getRange('J9').setValue(client);
  ss.getRange('I10').setValue(syomei);
  ss.getRange('D17').setValue(start);
  ss.getRange('F17').setValue(finish);
  ss.getRange('D18').setValue(terms);
  ss.getRange('M41').setValue(ukeA);

  const nameT = genmei + "◆" + client + "様注文書" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");//注文書 
  const nameU = genmei + "◇" + client + "様請書" + Utilities.formatDate(new Date,"JST","yyyyMMddhhmmss");//請書

  SpreadsheetApp.flush();

  const saveFld = 注文請書現場フォルダId(gen,ds);
  const portrait = '&portrait=true';
  const pdfId1 = createPdf(sss,ss,nameT,saveFld,portrait);
  const pdfId2 = createPdf(sss,ss2,nameU,saveFld,portrait);

  const pdfUrl = DriveApp.getFileById(pdfId1).getUrl();
  //チャットにできましたよの送信
  const chat = chatUrl("管理1");//丸山
  const msg = "【注文請書作成】自社物件\nの作成に成功しました\n確認お願いします\n" + pdfUrl;
  sendChat(chat,msg); 

  //予算書に記録する
  const sssY = SpreadsheetApp.openById(fId);
  const ssY = sssY.getSheetByName('シート1');
  if(keiyakuSct == "契約" || keiyakuSct == "変更"){
    const row = 88;
    ssY.getRange('J102').setValue(ukeA);
    ssY.getRange('L' + row).setValue(pdfId1);
  }else if(keiyakuSct == "追加"){
    const row = ssY.getRange('K88').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1;
    const num = row - 87;
    ssY.getRange('K' + row).setValue("注文書" + num);
    ssY.getRange('M' + row).setValue(ukeA);
    ssY.getRange('L' + row).setValue(pdfId1);
  }
}

//◆◆顧客情報を取得する（郵便番号と住所)
function 顧客情報(client){
  const id = getSPId("業者顧客担当");
  const sss = SpreadsheetApp.openById(id);
  const ss = sss.getSheetByName('顧客台帳');
  const last = ss.getRange('A1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  for(let i=2; i<=last; i++){
    if(ss.getRange('B' + i).getValue() == client){
      const post = ss.getRange('E' + i).getValue();
      const add = ss.getRange('F' + i).getValue();
      return [post,add];
    }
  }
}
//◆◆注文請書フォルダの現場フォルダIDを取得する
function 注文請書現場フォルダId(gen,ds){
  const nentuki = 本日の年度();
  const nendo = nentuki[0];
  const tuki = nentuki[1];
  const fldName = nendo + "年"
  const oyaId = ドライブId("注文請書");//注文請書ドライブ
  
  const getFldId = (id,name) => {
    const oya = DriveApp.getFolderById(id);
    const folders = oya.getFoldersByName(name);
    if(folders.hasNext()){
      const fold = folders.next()
      const fid = fold.getId();
      return fid;
    }else{
      //ない場合は作成する
      const id = oya.createFolder(name).getId();
      return id;
    }
  }
  const nenFldId = getFldId(oyaId,fldName);
  const genFldIdF = () => {
    if(ds == "D"){
      const nameF = "注文書大規模（ｽﾍﾟｰｽにもあげる）";
      const fldAId = getFldId(nenFldId,nameF);
      const gFldId = getFldId(fldAId,gen);
      return gFldId;
    }else if(ds == "S"){
      const nameF = "注文書小規模（ｽﾍﾟｰｽにもあげる）";
      const fldAId = getFldId(nenFldId,nameF);
      const tukiFName = tuki + "月";
      const tukiFldId = getFldId(fldAId,tukiFName);
      const gFldId = getFldId(tukiFldId,gen);
      return gFldId;
    }
  }
  const genFldId = genFldIdF();
  const genFld = DriveApp.getFolderById(genFldId);
  return genFld;
}
//🔸🔸契約書登録リストの取得
function getKeiyakuList(fId){
  
  const sssY = SpreadsheetApp.openById(fId);
  const ssY = sssY.getSheetByName('シート1');
  const ds = (ssY.getRange('C102').getValue() == "D" || ssY.getRange('C102').getValue() == "S2")? ssY.getRange('C102').getValue(): "S";
  if(ds == "D" || ds == "S2"){
    const lastF = () => {
      const r = ssY.getRange('K88').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
      if(r == 101){//追加はない
        return 88;
      }else{
        return r;
      }
    }
    const last = lastF();
    const num = last - 88 + 1;
    const list = () => {
      let rn = [[num]];
      for(let i = 88; i <= last; i++){
        const rn1 = [
          ssY.getRange('K' + i).getValue(),
          ssY.getRange('M' + i).getValue(),
          ssY.getRange('L' + i).getValue()
        ];
        rn.push(rn1);
      }
      return rn;
    }
    const data = {
      list:list()
    };
    const json = JSON.stringify(data);
    return json;
  }else if(ds == "S"){
    let rn = [[1]];
    const kin = ssY.getRange('J19').getValue();
    const fId = ssY.getRange('T16').getValue();
    const vals = ["",kin,fId];
    rn.push(vals);
    const data = {
      list: rn
    }
    const json = JSON.stringify(data);
    return json;
  }
}

//😃😃出来高チャット（物件チャットスペースに投稿）
function sendBukkenSpace(chatUrl, text){
  try {
    sendChat(chatUrl, text);
    return "OK"; // 成功したことをクライアントに伝える
  } catch(e) {
    return "エラー詳細: " + e.message; // messegeをmessageに修正
  }
}
//😃😃長谷工の請求書保存＆注文書＆請求スペースに投稿
function uploadHaseSeiFile(base64Data, contentType, fileName, genS, kinds, kin, tukiS) {
  // 1. PDFチェック（サーバー側でも念のため確認）
  if (contentType !== 'application/pdf') {
    throw new Error("PDFファイルのみアップロード可能です。");
  }

  try {
    // 2. 保存先フォルダの取得（genSを使用）
    const fldId = 請求書保存先(tukiS, genS); 
    const saveFld = DriveApp.getFolderById(fldId);

    // 3. ファイル名の作成とBlobの生成
    const newFileName = "【請求書】" + genS + " " + kinds + "_" + kin + "円.pdf";
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, contentType, newFileName);

    // 4. ファイルを保存
    const newFile = saveFld.createFile(blob);
    const fUrl = newFile.getUrl();

    // 5. 物件チャットスペースに投稿
    // 引数: kFileId, uke, genmei, kinds, ts, fUrl
    const kFileId = ""; // 今回は新規作成なので空
    const ts = "請求書";
    
    // 前に作った tyumonSpace を呼び出す
    tyumonSpace(kFileId, kin, genS, kinds, ts, fUrl);

    return fUrl; // 成功したらURLを返す
  } catch (e) {
    throw new Error("アップロード中にエラーが発生しました: " + e.message);
  }
}

