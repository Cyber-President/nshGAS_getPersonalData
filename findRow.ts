// 検索値の行数を返す
function findRow(val, rangeValuse, col, lastRow) {

    for (var i = 1; i <= lastRow; i++) {
        if (rangeValuse[i][col] === val) {
            return i;
        }
    }
    return 0;
}