//◆◆稼働中の更新したひな形ファイルを予算書ファイルに反映する
function 予算書ひな形反映() {
  const sssH = SpreadsheetApp.openById('11lGcIPD6L7QMeNRABPMD4bMKXH2gLaacHpGmOFU_dKQ'); // 予算書ひな形ファイル
  const hina = sssH.getSheetByName('シート1');

  const nentuki = 年度年月();
  const nendo = nentuki[0];
  const dId = 工事台帳(nendo);
  const sssD = SpreadsheetApp.openById(dId);
  const ssD = sssD.getSheetByName('大小現場分け');

  const last = ssD.getRange('A100').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  for (let i = 2; i <= last; i++) {
    const id = ssD.getRange('E' + i).getValue();
    Logger.log(ssD.getRange('H' + i).getValue());
    if (id) { // ファイルIDが空じゃない場合だけ処理
      const sss = SpreadsheetApp.openById(id);
      const ss = sss.getSheetByName('シート1');

      const source1 = hina.getRange('M88');
      const target1 = ss.getRange('M88');
      copyFullRangeBetweenSheets(source1,target1)
      // const source2 = hina.getRange('AQ88:AQ100');
      // const target2 = ss.getRange('AQ88:AQ100');
      // copyFullRangeBetweenSheets(source2,target2);
    }
  }
}

function copyFullRangeBetweenSheets(sourceRange, targetRange) {
  // ① 数式をコピー
  targetRange.setFormulas(sourceRange.getFormulas());

  // ② 数式がないところは値をコピー
  const formulas = sourceRange.getFormulas();
  const values = sourceRange.getValues();
  for (let i = 0; i < formulas.length; i++) {
    for (let j = 0; j < formulas[i].length; j++) {
      if (!formulas[i][j]) { // 数式がないなら
        targetRange.getCell(i + 1, j + 1).setValue(values[i][j]);
      }
    }
  }

  // ③ 書式（表示形式）コピー ※ここ修正！
  targetRange.setNumberFormats(sourceRange.getNumberFormats());

  // ④ 背景色コピー
  targetRange.setBackgrounds(sourceRange.getBackgrounds());

  // ⑤ フォント色・サイズコピー
  targetRange.setFontColors(sourceRange.getFontColors());
  targetRange.setFontSizes(sourceRange.getFontSizes());

  SpreadsheetApp.flush();
}


