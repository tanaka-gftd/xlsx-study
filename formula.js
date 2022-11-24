/* Excelのセル内に計算式を設定 */

/* xlsx-populateライブラリのformula関数は、googleスプレッドシートでは未対応の模様（解析エラーになる） */


const XlsxPopulate = require('xlsx-populate');

//新しいワークブックを作成して、セル内に計算式を設定していく
XlsxPopulate.fromBlankAsync().then(workbook => {

  //簡単な足し算の例
  workbook.sheet(0).cell('A1').value(1);  //Sheet1のA1セルに1を設定
  workbook.sheet(0).cell('A2').formula('=A1+2');  //Sheet1のA2セルに、'A1+2'の式を設定 

  //異なるセルの値を参照する例
  const sheet2 = workbook.addSheet('Sheet2');
  sheet2.cell('A1').value(9);
  workbook.sheet(0).cell('B1').formula('=Sheet2!A1&" * 3は、"&Sheet2!A1*3');  //取得した値に3を掛け算する

  //範囲にまとめて式を設定する例
  workbook.sheet(0).cell('C1').value(1);
  workbook.sheet(0).range('C2:C11').formula('=INDIRECT(ADDRESS(ROW()-1,COLUMN())) * 2');  //自分自身の１つ上のセルの 2倍を計算する式を、C2~C11に設定

  //作成したものをファイルに書き出し
  return workbook.toFileAsync("./xlsxFiles/formula.xlsx");
});