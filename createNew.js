const XlsxPopulate = require('xlsx-populate');

//空のワークブックを作成し、文章を入れる
//XlsxPopulate.fromBlankAsync()...空のExcelオブジェクト(workbook)を作成する
XlsxPopulate.fromBlankAsync().then(workbook => {  //then以下のworkbookは、コールバック関数の引数です

  //ワークブックのSheet1のA1セルに文章を入れる
  /* 
    Excelオブジェクト.sheet(シート名)...Excelのシートを、名前or番号で取得
    sheet.cell(セル番地)...シートから指定されたアドレスのセルを取得
    cell.value(値)...セルに値を書き込む  
  */
  workbook.sheet('Sheet1').cell('A1').value('新しく作った Excel');

  //Excelファイルの書き出し
  //Excelオブジェクト.toFileAsync...Excelを書き出す(引数で書き出すファイル名を指定できる)
  return workbook.toFileAsync("./Book1.xlsx");
});
