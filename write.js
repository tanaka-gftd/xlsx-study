/* セル範囲への書き込み */


const XlsxPopulate = require('xlsx-populate');

/* 
  Excel に書き込む方法は
    特定のセルを指定して書き込む方法
    範囲を指定して2次元配列を使って書き込む方法
    特定のセルを起点にして2次元配列を使って書き込む方法
  などがある
*/


//空のワークブックを作成して、指定した範囲内に2次元配列で値を書き込む
XlsxPopulate.fromBlankAsync().then(workbook => {

  //特定のセルを一つ指定して、そのセルに値を書き込む
  workbook.sheet(0).cell("A1").value('得点表');

  //range関数を使ってセル範囲を指定し、2次元配列を使うことで複数の値を値を書き込む
  //ここでは セル番地B1~D1 の３つのセルで構成された範囲を指定している
  //range...シート内のセル範囲を指定
  workbook.sheet(0).range("B1:D1").value(
    [
      ['英語', '国語', '数学']
    ]
  );

  //指定したセルを起点に、2次元配列で複数の値をまとめて書き込む
  //ここでは セル番地A2 が起点
  workbook.sheet(0).cell("A2").value(
    [
      ['aくん'],
      ['bくん'],
      ['cくん'],
      ['dくん'],
      ['eくん']
    ]
  );

  //range関数によるセル範囲の指定は、複数行、複数列でもできる
  //ここでは セル番地B2~D6 の15個のセルで構成された範囲を取得している
  const range = workbook.sheet(0).range("B2:D6");

  //range.value の引数に関数を設定すると書き込む値をプログラムで設定できる
  //ここでは、指定されたセル範囲内の各セルに、0~100までの自然数の乱数を書き込んでいる
  range.value((cell, ri, ci, range) => Math.floor(Math.random() * 101));

  //Excelファイルへの書き出し
  return workbook.toFileAsync("./points.xlsx");
});