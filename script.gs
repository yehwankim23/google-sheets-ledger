function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("금전출납부")
    .addItem("다음 달 만들기", "createNextMonth")
    .addToUi();
}

function createNextMonth() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = activeSpreadsheet.getSheets();
  const currentSheet = sheets[0];
  const currentSheetName = currentSheet.getSheetName();
  let [year, month] = currentSheetName.split(" ");
  year = Number(year.slice(0, -1));
  month = Number(month.slice(0, -1)) + 1;

  if (month === 13) {
    const ui = SpreadsheetApp.getUi();

    if (
      ui.alert(
        `'${year + 1}년 1월'을 만들까요? '아니요'를 선택하면 '다음 달 만들기'가 취소됩니다.`,
        ui.ButtonSet.YES_NO
      ) === ui.Button.YES
    ) {
      year++;
      month = 1;
    } else {
      return;
    }
  }

  const newSheet = sheets.at(-1).copyTo(activeSpreadsheet);
  const newSheetName = `${year}년 ${month}월`;
  newSheet.setName(newSheetName);

  newSheet
    .getRange("A1")
    .setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText(`금전출납부 (${newSheetName})`)
        .setTextStyle(0, 6, SpreadsheetApp.newTextStyle().setForegroundColor("#000").build())
        .build()
    );

  const maxRows = currentSheet.getMaxRows();
  newSheet.getRange("D3").setValue(`='${currentSheetName}'!D${maxRows}`);
  newSheet.getRange("E3").setValue(`='${currentSheetName}'!E${maxRows}`);
  newSheet.getRange("F3").setValue(`='${currentSheetName}'!F${maxRows}`);

  activeSpreadsheet.setActiveSheet(newSheet);
  activeSpreadsheet.moveActiveSheet(0);
}

function onEdit(event) {
  const range = event.range;
  const sheet = range.getSheet();
  const rowIndex = range.getRowIndex();

  if (sheet.getRange(rowIndex + 2, 3).getValue() === "월    계") {
    sheet.insertRowAfter(rowIndex + 1);
    sheet.getRange(rowIndex, 6).setValue(`=F${rowIndex - 1}+D${rowIndex}-E${rowIndex}`);
  }

  const maxRows = sheet.getMaxRows();
  sheet.getRange(maxRows - 1, 4).setValue(`=SUM(D4:D${maxRows - 4})`);
  sheet.getRange(maxRows - 1, 5).setValue(`=SUM(E4:E${maxRows - 4})`);
  sheet.getRange(maxRows, 4).setValue(`=D3+D${maxRows - 1}`);
  sheet.getRange(maxRows, 5).setValue(`=E3+E${maxRows - 1}`);
  sheet.getRange(maxRows, 6).setValue(`=F${maxRows - 4}`);
}
