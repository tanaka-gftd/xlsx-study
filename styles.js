/* Excelのシートに、様々なスタイルをあてる */


const XlsxPopulate = require('xlsx-populate');

//空のワークブックを作成し、その中のセルに様々なスタイルを設定していく
XlsxPopulate.fromBlankAsync().then(workbook => {

  //使用するシートを取得
  const sheet = workbook.sheet('Sheet1');

  //指定したセルの文字にbold（太字）を設定
  sheet.cell('A1').value('太字');
  sheet.cell('A1').style("bold", true);

  //指定したセルの文字にイタリック（斜体）を設定
  sheet.cell('B1').value('イタリック').style("italic", true);

  //指定したセルの文字にボールドとイタリック両方を当てる
  sheet.cell('C1').value('両方').style({"italic": true, "bold": true});

  //指定したセルに数値フォーマットを設定（小数点以下2位まで）
  sheet.cell('D1').value(1234.56);
  sheet.cell('D1').style("numberFormat", "0.00");

  //指定したセル範囲に背景色をまとめて設定
  sheet.cell('A2').value('水色の背景');
  sheet.range('A2:E2').style("fill", "00ffff");

  //指定したセル範囲にランダムな色の背景色を設定
  sheet.cell('A3').value('ランダムな背景');
  const Hex = '012345678abcdef';
  sheet.range('B3:F3').style({
    fill: (cell, ri, ci, range) => {
      let rgb = '';
      for(let i = 0; i < 6; i++){
        rgb += Hex[Math.floor(Math.random()*Hex.length)];
      }
      return rgb;
    }
  });

  //指定した行目に枠線のスタイルをあてる
  sheet.row(4).style("border", true);

  //指定した列に中央寄せのスタイルをあてる
  sheet.column("C").style("horizontalAlignment", "center");
  sheet.cell("C5").value('中央寄せ');

  //指定したセルに複雑なパラメーターを設定する（googleのスプレッドシートには未対応？）
  sheet.cell("A6").value('複雑な背景→');
  sheet.cell("B6").style("fill", {
    type: "pattern",
    pattern: "darkDown",
    foreground: {
      rgb: "ff0000"
    },
    background: {
      theme: 3,
      tint: 0.4
    }
  });

  //シートの幅を微調整
  sheet.column('A').width(15);

  //作成したものを新規ファイルに書き出し
  return workbook.toFileAsync("./styles.xlsx");
});