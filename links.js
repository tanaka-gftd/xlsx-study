/* Excelのセル内にリンクを作成 */


const XlsxPopulate = require('xlsx-populate');

//新しいワークブックを作成し、セル内にリンクを設定していく
XlsxPopulate.fromBlankAsync().then(workbook => {
  
  //リンク先となるシートを作成しておく
  const sheet2 = workbook.addSheet('Sheet2');
  sheet2.cell('A1').value("飛び先");

  //Sheet1のA1セル内に、リンクを設定(ついでにスタイルも設定)
  workbook.sheet(0).cell('A1').value("リンクテキスト")
    .style({fontColor: "0563c1", underline: true})
    .hyperlink("https://www.nicovideo.jp/");

  //Sheet1のA2セル内に、リンクに加えツールチップも設定することで、A2セルにカーソルを合わせると文字列が出るようにした(ついでにスタイルも設定)
  workbook.sheet(0).cell('A2').value("ニコニコ動画")
    .style({fontColor: "0563c1", underline: true})
    .hyperlink({hyperlink:"https://www.nicovideo.jp/", tooltip: "ニコニコ動画"});

  //A2セル内に貼られたリンク情報を取得し、コンソールに表示
  const value = workbook.sheet(0).cell('A2').hyperlink();
  console.log(value);

  //Sheet1のセルA3内に、メール送信リンクを設定
  workbook.sheet(0).cell('A3').value("クリックでメール送信")
    .hyperlink({email: "sample@nnn.ed.jp", emailSubject: "大変お忙しい所ではございますが、sampleさん..."});

  //Sheet1のA4セル内に、同じワークブック内の別シートのセル（ここではSheet2のA1セル）へのリンクを設定
  workbook.sheet(0).cell('A4').value("クリックで別シートへ遷移")
    .hyperlink("Sheet2!A1");

  //ファイルへの書き出し
  return workbook.toFileAsync("./xlsxFiles/links.xlsx");
});