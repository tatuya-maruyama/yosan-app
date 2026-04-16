//clasp push && clasp deploy -i AKfycbxYYDbl2CSy6BTeAO-jTEdBbNvBclOGpuDNEzAQgwlpt8ovS1gE6WiZzkuhtY4M3ng
//😃😃アップロードされたエクセルファイル（内訳を予算書に反映する)
function readCSV(json) {
  const data = JSON.parse(json);
  const utiId = data.utiId;
  const fId = data.fId;
  const col = data.column; // 例: ["A","C","B","E","D"]

  const ss = 内訳鏡(fId);
  const rw = ss.getRange('E2000').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  const resetF = () => {
    if(rw == 1){
      return "OK";
    } else {
      const res = csvReset(fId);  
      return res;
    }
  }

  const reset = resetF();
  if (reset === "OK") {
    const convertedSpreadsheet = SpreadsheetApp.openById(utiId);
    const ssE = convertedSpreadsheet.getSheets()[0];
    const values = ssE.getDataRange().getValues();

    if (values.length > 0 && values[0].length > 0) {
      const maxCols = 12;

      // 文字 → 数値インデックス（"A"→0, "C"→2...）
      const colIndex = col.map(c => {
        return c.toUpperCase().charCodeAt(0) - 65;
      });

      // 並び替え
      const reorderedValues = values.map(row => {
        return colIndex.map(i => row[i] || "");
      });

      ss.getRange(1, 4, reorderedValues.length, reorderedValues[0].length).setValues(reorderedValues);
    } else {
      return "❌ ファイルに読み込むデータがありません";
    }

    const formula = 関数セット(ss);
    if (formula === "OK") {
      return SpreadsheetApp.openById(fId).getUrl();
    } else {
      return "⚠️ データは読み込みましたが、計の数式セットに失敗しました";
    }
  } else {
    return "❌ リセット処理に失敗しました";
  }
}


function 関数セット(ss){
    let ls = ss.getRange('D1').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    const fast = ss.getRange('D' + ls).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();//内訳項目の初めの行
    const last = ss.getRange('E2000').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();//内訳の最終行
    // 単価×数量（提出金額）と（原価金額）の数式を一括設定
    const numRows = last - fast + 1;
    ss.getRange(fast, 10, numRows).setFormulaR1C1('=IF(RC[-2]="","",IFERROR(RC[-2]*RC[-1],""))'); // J列 = H×I
    ss.getRange(fast, 14, numRows).setFormulaR1C1('=IF(RC[-6]="","",IFERROR(RC[-6]*RC[-1],""))'); // N列 = H×M
    
    //大項目、小項目計のセット
    // 小項目計のセット
    const 小項目 = (i, s) => {
      let res = "なし";
      const st = i + 1;
      const allD = ss.getRange('D' + st + ':D' + s).getValues();
      const allE = ss.getRange('E' + st + ':E' + s).getValues();

      for (let t = 0; t < allD.length; t++) {
        const en = allD.length - 1;
        if (t == en) {
          Logger.log(res);
          return res;
        } else {
          const val2 = allD[t][0];
          if (val2 !== "") {
            const kei2 = val2 + "－計";
            for (let t2 = t; t2 < allE.length; t2++) {
              if (allE[t2][0] == kei2) {
                const t3 = t2 - 1;
                const row1 = i + 1 + t;
                const row2 = i + 1 + t3;
                const row3 = i + 1 + t2;
                const masS1 = '=SUM(J' + row1 + ':J' + row2 + ')';
                const masS2 = '=SUM(N' + row1 + ':N' + row2 + ')';
                ss.getRange('J' + row3).setFormula(masS1);
                ss.getRange('N' + row3).setFormula(masS2);
                res = "あり";
                t = t2;
                break;
              }
            }
          }
        }
      }
      return res; // これ重要！
    };

    // 大項目の小計を渡す
    const 大項目セット = (val, s) => {
      const masM1 = '=J' + s;
      const masM2 = '=N' + s;
      const dVals = ss.getRange('D4:D' + ls).getValues().flat();
      for (let k = 0; k < dVals.length; k++) {
        if (dVals[k] == val) {
          const r = k + 4;
          ss.getRange('J' + r).setValue(masM1);
          ss.getRange('N' + r).setValue(masM2);
          break;
        }
      }
    };
    const rnD = ss.getRange('D' + fast + ':D' + last).getValues().flat();
    const rnE = ss.getRange('E' + fast + ':E' + last).getValues().flat();
    for(let i=0; i<rnD.length; i++){
        const val = rnD[i];
        Logger.log(i);
        if(rnE[i] == ""){
            const row = i + fast;
            ss.getRange('H' + row + ':J' + row).setValue("");//よけいなゼロを消す
        }else if(!isNaN(val) && val !==""){
            const kei = val + "－計";
            for(let s=i; s<=last; s++){  
                if(rnE[s] == kei){
                    //数式をセットする（iからｓまでの間に小項目がないかをチェックする
                    const iA = i + fast;
                    const sA = s + fast;
                    const judge = 小項目(iA,sA);
                    Logger.log("judge " + judge);
                    if(judge == "なし"){//小項目なかった
                        const s2 = sA - 1;
                        const masD1 = '=SUM(J' + iA + ':J' + s2 + ')';
                        const masD2 = '=SUM(N' + iA + ':N' + s2 + ')';
                        ss.getRange('J' + sA).setFormula(masD1);
                        ss.getRange('N' + sA).setFormula(masD2);
                    }else if(judge == "あり"){//小項目あった
                        const s2 = sA - 1;
                        const masD1 = '=SUM(J' + iA + ':J' + s2 + ')/2';
                        const masD2 = '=SUM(N' + iA + ':N' + s2 + ')/2';
                        ss.getRange('J' + sA).setFormula(masD1);
                        ss.getRange('N' + sA).setFormula(masD2);
                    } 
                    大項目セット(val,sA);//大項目番号と小計ROWを渡す  
                    break;
                }
            }
        }
    }

    //粗利計算関数セット
    const arari = '=IFERROR((J4-N4)/J4,"")';
    ss.getRange('O4').setFormula(arari);
    ss.getRange('O4').setNumberFormat('0.0%');
    rn = ss.getRange('O4')
    const last2 = last + 100;//追加の分も考量して
    rn2 = ss.getRange('O5' + ':O' + last2);
    rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);//数式


    //調整費セット
    const tyosei = last + 2;
    const tyRow = ls + 1;
    const tyoseiNum = ss.getRange('D' + ls).getValue() + 1;
    ss.getRange('D' + tyosei).setValue(tyoseiNum);
    ss.getRange('E' + tyosei).setValue("調整");
    ss.getRange('D' + tyRow).setValue(tyoseiNum);
    ss.getRange('E' + tyRow).setValue("調整");

    //本工事プルダウンセット
    const tyoEnd = tyosei + 19;
    const pulVal = ss.getRange('R1').getValue();
    const pulDataval = ss.getRange('R1').getDataValidation();
    ss.getRange('L' + fast + ':L' + tyoEnd).setValue(pulVal);
    ss.getRange('L' + fast + ':L' + tyoEnd).setDataValidation(pulDataval);//プルダウンを設定

    rn = ss.getRange('R1');
    rn2 = ss.getRange('L' + fast + ':L' + tyoEnd);
    rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
    ss.getRange('Q' + fast + ':Q' + tyoEnd).setValue('本工事');

    //追加項目セット
    const tuika = tyosei + 20;
    const tRow = tyRow + 1;
    const tuikaNum = tyoseiNum + 1;
    const tuikaEnd = tuika + 100;
    ss.getRange('D' + tuika).setValue(tuikaNum);
    ss.getRange('E' + tuika).setValue("追加");
    ss.getRange('D' + tRow).setValue(tuikaNum);
    ss.getRange('E' + tRow).setValue("追加");
    //色を変える
    ss.getRange('D' + tuika + ':K' + tuikaEnd).setBackground('#EEEEEE');
    //追加バージョンにする
    ss.getRange('Q' + tuika + ':Q' + tuikaEnd).setValue('追加');

    //追加プルダウンをセット
    const pulVal2 = ss.getRange('R2').getValue();
    const pulDataval2 = ss.getRange('R2').getDataValidation();
    ss.getRange('L' + tuika + ':L' + tuikaEnd).setValue(pulVal2);
    ss.getRange('L' + tuika + ':L' + tuikaEnd).setDataValidation(pulDataval2);//プルダウンを設定
    rn = ss.getRange('R2');
    rn2 = ss.getRange('L' + fast + ':L' + tyoEnd);
    rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);

    //注文書作成用の関数セット
    rn = ss.getRange('A24:C24');
    rn2 = ss.getRange('A25:C' + tuikaEnd);
    rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);//数式

    //調整＆追加の計数式セット
    let masT = '=IF(H' + tyosei + '="","",' + 'IFERROR(H' + tyosei + '*I' + tyosei + ',""))';
    let masT2 = '=IF(H' + tyosei + '="","",' + 'IFERROR(H' + tyosei + '*M' + tyosei + ',""))';

    next = tyosei + 1;

    rn = ss.getRange('J' + tyosei);
    rn.setFormula(masT);
    rn2 = ss.getRange('J' + next + ':J' + tuikaEnd);
    rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);//数式

    rn = ss.getRange('N' + tyosei);
    rn.setFormula(masT2);
    rn2 = ss.getRange('N' + next + ':N' + tuikaEnd);
    rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);//数式

    return "OK";
}

//◆◆EXcel反映リセット◆◆
function csvReset(id){
    const ss = 内訳鏡(id);
    const ss1 = 予算書シート(id);
    ss.getRange('D2:P2000').clearContent();
    ss.getRange('D2:K2000').setBackground('#ffffff');
    ss.getRange('A25:C2000').clearContent();
    const rn = ss.getRange('L24');
    const rn2 = ss.getRange('L25:L2000');
    rn.copyTo(rn2,SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
    const val = ss.getRange('Q3').getValue();
    ss.getRange('Q37:Q2000').setValue(val);
    ss.getRange('V3:V23').clearContent();//登録してある業者
    ss.getRange('AE3:AE23').clearContent();//登録してある業者ROW

    ss.getRange('V1:X2').clearContent();//添付＆ccメールクリア

    const judgeF = (i) => {
        const res = (i - 1) % 3
        if(res === 0){
            return "NG";
        }else{
            return "OK";
        }
    }
    for(let i=4; i<=42; i++){
        const judge = judgeF(i);
        if(judge == "OK"){
            ss1.getRange('B' + i).clearContent();
        }
    }
    SpreadsheetApp.flush();
    return "OK";
}

