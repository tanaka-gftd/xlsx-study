/* 検索と置き換え */


const XlsxPopulate = require('xlsx-populate');

//既存のワークブックを読み込んで、その中から文字列を検索 ＆ 文字列を置き換え ＆ 別ファイルに書き出し
XlsxPopulate.fromFileAsync("./xlsxFiles/points.xlsx").then(workbook => {

  //文字列検索、一致したセルを配列で返す
  workbook.find("得点表");  //一致したセルを配列で返す

  //指定したシートの中から文字列を検索（ここでは先頭のシート）
  workbook.sheet(0).find("得点表");  //一致したセルを配列で返す

  //特定のセルが指定した文字列を持っているかを判定
  workbook.sheet("Sheet1").cell("A1").find("得点表");  //trueかfalseで返す

  //ワークブック全体から文字列を検索し、別の文字列に置換（ここでは"得点表" → "点数表"）
  workbook.find("得点表","点数表");  //一致したセルを配列で返す

  //ワークブック全体から文字列を正規表現を用いて検索し、メソッドを用いて置換(ここでは半角英字の小文字を、大文字に置き換え)
  workbook.find(/[a-z]+/g, match => match.toUpperCase());

  //以上の操作によって作成されたExcelファイルを、新規ワークブックに書き出し
  return workbook.toFileAsync("./xlsxFiles/points2.xlsx");
});