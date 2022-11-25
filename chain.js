/* xlsx-populateライブラリではメソッドチェーンも使える */


const XlsxPopulate = require('xlsx-populate');

//空のワークブックを作成して、メソッドチェーンを使ってみる
XlsxPopulate.fromBlankAsync().then(workbook => {

  const sheet2 = workbook.addSheet('Sheet2');

  //複数のセルや範囲に値やスタイルを設定できる
  workbook
    .sheet(0)
      .cell("A1")
        .value("foo")
        .style("bold", true)
      .relativeCell(1, 0)
        .formula("A1")
        .style("italic", true)
  .workbook()  //チェーンの途中で、設定を行うシートを別のシートに変更できる
    .sheet(1)
      .range("A1:B3")
        .value(5)
      .cell(0, 0)
        .style("underline", "double");

    return workbook.toFileAsync("./xlsxFiles/chain.xlsx");         
});