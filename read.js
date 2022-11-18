const XlsxPopulate = require('xlsx-populate');

//既存のワークブックを取得して、その値を読み込む
//XlsxPopulate.fromFileAsync(xlsxファイルのパス)...既存のxlsxファイルを取得し、workbookを作成
XlsxPopulate.fromFileAsync("./Book1.xlsx").then(workbook => {  //then以下のworkbookは、コールバック関数の引数です

  //Sheet1のA1セルの値を取得する
  const value = workbook.sheet("Sheet1").cell("A1").value();

  //取得した値をログに出力
  console.log(value);
});
