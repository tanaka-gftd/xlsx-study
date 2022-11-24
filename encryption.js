/* パスワードで保護されたExcelファイルを読み書き */


const XlsxPopulate = require('xlsx-populate');

//パスワードで保護されたワークブックを作成してみる
XlsxPopulate.fromBlankAsync().then(workbook => {

  workbook.sheet(0).cell("A1").value("機密情報");

  //ファイル書き込み時に、読み込む際のパスワードを付与する
  workbook.toFileAsync("./xlsxFiles/encryption.xlsx", {password:"S3cret!"});
});