/* 既存のExcelファイルにシートを追加 */


const XlsxPopulate = require('xlsx-populate');

//workbookにシートを追加していく
/* 
  シートを追加する方法は、
    シートを最後に追加する方法
    指定した番号の位置にシートを追加する方法
    指定したシートの前にシートを追加する方法
  がある
*/
XlsxPopulate.fromBlankAsync().then(workbook => {

  //workbookの最後に追加
  const sheet5 = workbook.addSheet('Sheet5');

  //数値で指定した場所にシートを追加(番号は0から始まる)
  const sheet2 = workbook.addSheet('Sheet2', 1);

  //指定したシートの前に、シートを追加
  const sheet3 = workbook.addSheet('Sheet3', 'Sheet5');

  //シートの指定は、シートオブジェクトが格納された変数でもできる
  //指定したシート(シートオブジェクト)の前に、シートを追加
  const sheet4 = workbook.addSheet('Sheet4', sheet5);

  //Excelファイルの書き出し
  return workbook.toFileAsync("./xlsxFiles/sheetTest.xlsx");
});
