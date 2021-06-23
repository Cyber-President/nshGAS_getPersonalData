function clearSuperSheet() {
    //対象のスプレッドシートを設定
    const nskobe2021bms = SpreadsheetApp.getActiveSpreadsheet();
    const superSheet = nskobe2021bms.getSheetByName('SuperSheet');

    clearRange(superSheet, 4, 10, ["B","D","F","H","I"]); // profile
    clearRange(superSheet, 13, 16, ["B"]); // numeric
    clearRange(superSheet, 20, 20, ["A","C","F","I"]); //hobbis
    clearRange(superSheet, 24, 33, ["A","B","C","D","E"]); // studentInfomation
    clearRange(superSheet, 37, 37, ["A","B","C","D","E","F","H","I","J","K"]); // monoPass&Language
    clearRange(superSheet, 41, 50, ["A","B","C","D","E","F","H","I"]); // formPrg
    clearRange(superSheet, 54, 63, ["A","B","D","F","H"]); // formExam
    clearRange(superSheet, 67, 69, ["A","B","C"]); // interviewRecord

}

function clearRange(sheet, minRow, maxRow, colList) {
    for (var i = 0; i < colList.length; i++) {
        for (var row = minRow; row <=maxRow; row++) {
            sheet.getRange(colList[i]+String(row)).clearContent();
        }
    }
}