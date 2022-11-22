/* 日付のフォーマット指定 */


const XlsxPopulate = require('xlsx-populate');

//セルに日付を入力していく
XlsxPopulate.fromBlankAsync().then(workbook => {

  //指定したセルに日付を設定
  workbook.sheet(0).cell('A1').value(new Date(2020, 8, 22)).style("numberFormat", "yyyy年 m月 dd日");

  //指定したセルに算出した日付を設定
  const date = XlsxPopulate.numberToDate(42788);
  workbook.sheet(0).cell('A2').value(date).style("numberFormat", "dddd, mmmm dd, yyyy");

  //シートの幅を微調整
  workbook.sheet(0).column('A').width(30);

  //作成したものを新規ファイルに書き出し
  return workbook.toFileAsync("./date.xlsx");
});