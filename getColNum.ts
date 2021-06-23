// DBの各列番号を返す
function getColNum(dbMaxCol, dbRangeValue) {
    for (var i = 1; i <= dbMaxCol; i++) {
        var getColName = dbRangeValue[0][i - 1];
        switch (getColName) {
            case '学籍番号':
                colNumStudentId = i - 1;
                continue;
            case 'コース':
                colNumCourseNow = i - 1;
                continue;
            case '通学入学年月':
                colNumEnrollmentDate = i - 1;
                continue;
            case '性別':
                colNumGender = i - 1;
                continue;
            case '生年月日':
                colNumBirthday = i - 1;
                continue;
            case '年齢':
                colNumAge = i - 1;
                continue;
            case '前年度コース':
                colNumCourseOld = i - 1;
                continue;
            case 'Slack FullName':
                colNumSlackFullName = i - 1;
                continue;
            case 'Slack Display Name':
                colNumSlackDisplayName = i - 1;
                continue;
            case 'キャンパス担任':
                colNumTeacherCampus = i - 1;
                continue;
            case 'N高 担任':
                colNumTeacherHighschool = i - 1;
                continue;
            case '映像許諾':
                colNumPicPermission = i - 1;
                continue;
            case '英語レベル':
                colNumEnglishLevel = i - 1;
                continue;
            case 'コース略':
                colNumCourseNow = i - 1;
                continue;
            case '姓名':
                colNumNameKanji = i - 1;
                continue;
            case 'セイメイ':
                colNumNameFurigana = i - 1;
                continue;
            case 'int(学年)':
                colNumGrade = i - 1;
                continue;
            case 'PJSリンク':
                colNumPjsUrl = i - 1;
                continue;
            case 'N高メールアドレス':
                colMail = i - 1;
                continue;
            default:
                continue;
        }
    }
}