// 😃😃原価設定：設定した粗利で原単価をセットする
function arariSet(fId, arariA) {
  const ss = 内訳鏡(fId);
  const last = ss.getRange('M2000').getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const tanka = ss.getRange('I2:I' + last).getValues().flat(); // 単価を取得
  const arari = 1 - (Number(arariA) / 100); // 粗利率を小数で計算

  const genka = tanka.map(val => {
    if (typeof val === 'number' && !isNaN(val)) {
      return [Math.round((val * arari) * 10) / 10]; // 小数1桁で四捨五入して2次元に戻す
    } else {
      return [null]; // 数値でない場合は空にする
    }
  });

  ss.getRange('M2:M' + last).setValues(genka); // 原価列にセット
  return "OK";
}

