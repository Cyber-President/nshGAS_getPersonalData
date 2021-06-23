function getDate() {
    //セルデータを削除
    clearSuperSheet()

    //対象のスプレッドシートを設定
    const nskobe2021bms = SpreadsheetApp.getActiveSpreadsheet();
    const superSheet = nskobe2021bms.getSheetByName('SuperSheet');
    const db = nskobe2021bms.getSheetByName('DB');
    const diaryTA = nskobe2021bms.getSheetByName('日誌');

    // DBを二次元配列化
    const dbLastRow = db.getLastRow(); // DBの最終行を取得
    const dbLastCol = db.getLastColumn();
    const dbRangeValues = db.getRange(1, 1, dbLastRow, dbLastCol).getValues();

    // 検索値（シメイ）を取得
    const getSearchName = superSheet.getRange("B1").getValue();
    console.log("SearcName:%s", getSearchName);

    // DBの各列番号を取得
    getColNum(dbLastCol, dbRangeValues);
    // console.log('colNumNameFurigana:%s',colNumNameFurigana);

    // 検索値のDB行数を取得
    const targetRowDb = findRow(getSearchName, dbRangeValues, colNumNameFurigana, dbLastRow);

    // 学籍番号を取得
    const getStudentId = dbRangeValues[targetRowDb][colNumStudentId];
    // console.log('getStudentId:%s',getStudentId);


    const diaryTALastRow = diaryTA.getLastRow(); //日誌の最終行を取得
    const diaryTARangeValue = diaryTA.getRange(1, 1, diaryTALastRow, 9).getValues();// 日誌を二次元配列化
    var targetRowDiaryTA = findRow(getStudentId, diaryTARangeValue, 0,diaryTALastRow);// 日誌行数を取得

    // プロフィールデータを辞書型に変換
    var profile = {

        studentId: getStudentId,
        nameFurigana: getSearchName,

        nameKanji: dbRangeValues[targetRowDb][colNumNameKanji],
        gender: dbRangeValues[targetRowDb][colNumGender],
        birthday: dbRangeValues[targetRowDb][colNumBirthday],
        age: dbRangeValues[targetRowDb][colNumAge],
        teacherCampus: dbRangeValues[targetRowDb][colNumTeacherCampus],
        teacherHighschool: dbRangeValues[targetRowDb][colNumTeacherHighschool],
        grade: dbRangeValues[targetRowDb][colNumGrade],
        courseNow: dbRangeValues[targetRowDb][colNumCourseNow],
        courseOld: dbRangeValues[targetRowDb][colNumCourseOld],
        enrollmentDate: dbRangeValues[targetRowDb][colNumEnrollmentDate],
        slackFullName: dbRangeValues[targetRowDb][colNumSlackFullName],
        slackDisplayName: dbRangeValues[targetRowDb][colNumSlackDisplayName],
        pjsUrl: dbRangeValues[targetRowDb][colNumPjsUrl],
        picPermission: dbRangeValues[targetRowDb][colNumPicPermission],
        mail: dbRangeValues[targetRowDb][colMail],

        englishLevel: dbRangeValues[targetRowDb][colNumEnglishLevel],
        chinese: 'null',

        classType: 'null', //特進,AL,nomal
        proNLevel: 'null', // α,β
        diaryTA1: diaryTARangeValue[targetRowDiaryTA][7],
        diaryTA2: diaryTARangeValue[targetRowDiaryTA][8],
        attendPctNoLate: 'null', // 出席率
        attendPctConLate: 'null', // 登校率
        facePic: 'null',

    }
    // console.log('profile.nameKanji:%s',profile.nameKanji);

    // Profileをセルに入力
    superSheet.getRange("B4").setValue(profile.studentId);
    superSheet.getRange("B5").setValue(profile.nameKanji);
    superSheet.getRange("B6").setValue(profile.nameFurigana);
    superSheet.getRange("B7").setValue(profile.gender);
    superSheet.getRange("B8").setValue(profile.birthday);
    superSheet.getRange("B9").setValue(profile.age);
    superSheet.getRange("B10").setValue(profile.pjsUrl);
    superSheet.getRange("D4").setValue(profile.grade);
    superSheet.getRange("D5").setValue(profile.courseNow);
    superSheet.getRange("D6").setValue(profile.proNLevel);
    superSheet.getRange("D7").setValue(profile.classType);
    superSheet.getRange("D8").setValue(profile.slackFullName);
    superSheet.getRange("D9").setValue(profile.slackDisplayName);
    superSheet.getRange("D10").setValue(profile.picPermission);
    superSheet.getRange("F4").setValue(profile.teacherCampus);
    superSheet.getRange("F5").setValue(profile.teacherHighschool);
    superSheet.getRange("F6").setValue(profile.diaryTA1);
    superSheet.getRange("F7").setValue(profile.diaryTA2);
    superSheet.getRange("F8").setValue(profile.enrollmentDate);
    superSheet.getRange("F9").setValue(profile.courseOld);
    superSheet.getRange("F10").setValue(profile.mail);
    superSheet.getRange("H4").setValue(profile.attendPctNoLate);
    superSheet.getRange("H5").setValue(profile.attendPctConLate);
    superSheet.getRange("I4").setValue(profile.facePic);

    // 【生徒情報】
    const studentInfo = nskobe2021bms.getSheetByName('生徒情報');
    const siLastRow = studentInfo.getLastRow(); // 生徒情報の最終行を取得
    const siRangeValues = studentInfo.getRange(2, 1, siLastRow, 10).getValues();// 趣味一覧を二次元配列化
    // 生徒情報のうち対象の生徒情報を抜き出し降順でリスト化
    var siListSort = [];
    for (var i=0; i<siRangeValues.length; i++){
        if (siRangeValues[i][3] == getStudentId){
            siListSort.unshift(siRangeValues[i]);
        }
    }
    // 生徒情報をセルに入力
    if (siListSort.length<10){
        for (var i=0;i<siListSort.length;i++){
            superSheet.getRange("A"+(24+i)).setValue(siListSort[i][0]);
            superSheet.getRange("B"+(24+i)).setValue(siListSort[i][1]);
            superSheet.getRange("C"+(24+i)).setValue(siListSort[i][7]);
            superSheet.getRange("D"+(24+i)).setValue(siListSort[i][8]);
            superSheet.getRange("E"+(24+i)).setValue(siListSort[i][9]);
        }
    }else{
        for (var i=0;i<10;i++){
            superSheet.getRange("A"+(24+i)).setValue(siListSort[i][0]);
            superSheet.getRange("B"+(24+i)).setValue(siListSort[i][1]);
            superSheet.getRange("C"+(24+i)).setValue(siListSort[i][7]);
            superSheet.getRange("D"+(24+i)).setValue(siListSort[i][8]);
            superSheet.getRange("E"+(24+i)).setValue(siListSort[i][9]);
        }
    }


    // 【趣味一覧】
    const hobbies = nskobe2021bms.getSheetByName('趣味一覧');
    const hobiesLastRow = hobbies.getLastRow(); // DBの最終行を取得
    const hobbiesRangeValues = hobbies.getRange(1, 1, hobiesLastRow, 7).getValues();// 趣味一覧を二次元配列化
    var targetRowHobbies = findRow(getStudentId, hobbiesRangeValues, 0,hobiesLastRow );// 趣味一覧行数を取得
    // 趣味一覧をセル入力
    superSheet.getRange("A20").setValue(hobbiesRangeValues[targetRowHobbies][2]);
    superSheet.getRange("C20").setValue(hobbiesRangeValues[targetRowHobbies][3]);
    superSheet.getRange("F20").setValue(hobbiesRangeValues[targetRowHobbies][4]);
    superSheet.getRange("I20").setValue(hobbiesRangeValues[targetRowHobbies][5]);


    // 【ものパス】
    const prgManual = SpreadsheetApp.openById('ENTER YOUR ID');
    const mPass = prgManual.getSheetByName('名簿');
    const mPassLastRow = mPass.getLastRow();  // ものパス取得状況の最終行を取得
    const mPassRangeValues = mPass.getRange(3,1,mPassLastRow,9).getValues();// ものパス取得状況を二次元配列化
    var targetRowMPass = findRow(profile.mail,mPassRangeValues,1,mPassLastRow);// ものパス取得状況行数を取得
    // ものパス取得状況をセル入力
    superSheet.getRange("A37").setValue(mPassRangeValues[targetRowMPass][2]);
    superSheet.getRange("B37").setValue(mPassRangeValues[targetRowMPass][3]);
    superSheet.getRange("C37").setValue(mPassRangeValues[targetRowMPass][4]);
    superSheet.getRange("D37").setValue(mPassRangeValues[targetRowMPass][5]);
    superSheet.getRange("E37").setValue(mPassRangeValues[targetRowMPass][6]);
    superSheet.getRange("F37").setValue(mPassRangeValues[targetRowMPass][7]);


    // 【ものづくりプラン】
    const monoPlanSheet = SpreadsheetApp.openByUrl('ENTER YOUR URL');
    const mPlan = monoPlanSheet.getSheetByName('フォームの回答 1');
    const mPlanLastRow = mPlan.getLastRow();  // ものづくりプランの最終行を取得
    const mPlanRangeValues = mPlan.getRange(3,1,mPlanLastRow,9).getValues();// ものづくりプランを二次元配列化
    // ものづくりプランのうち対象の情報を抜き出し降順でリスト化
    var mPlanList = [];
    for (var i=0; i<mPlanRangeValues.length; i++){
        if (mPlanRangeValues[i][1] == profile.mail){
            mPlanList.unshift(mPlanRangeValues[i]);
        }
    }
    // ものづくりプランをセル入力
    if (mPlanList.length<10){
        for (var i=0;i<mPlanList.length;i++){
            superSheet.getRange("A"+(41+i)).setValue(mPlanList[i][0]);
            superSheet.getRange("B"+(41+i)).setValue(mPlanList[i][2]);
            superSheet.getRange("C"+(41+i)).setValue(mPlanList[i][3]);
            superSheet.getRange("D"+(41+i)).setValue(mPlanList[i][4]);
            superSheet.getRange("E"+(41+i)).setValue(mPlanList[i][5]);
            superSheet.getRange("F"+(41+i)).setValue(mPlanList[i][6]);
            superSheet.getRange("G"+(41+i)).setValue(mPlanList[i][7]);
            superSheet.getRange("H"+(41+i)).setValue(mPlanList[i][8]);
            superSheet.getRange("I"+(41+i)).setValue(mPlanList[i][9]);
        }
    }else{
        for (var i=0;i<10;i++){
            superSheet.getRange("A"+(41+i)).setValue(mPlanList[i][0]);
            superSheet.getRange("B"+(41+i)).setValue(mPlanList[i][2]);
            superSheet.getRange("C"+(41+i)).setValue(mPlanList[i][3]);
            superSheet.getRange("D"+(41+i)).setValue(mPlanList[i][4]);
            superSheet.getRange("E"+(41+i)).setValue(mPlanList[i][5]);
            superSheet.getRange("F"+(41+i)).setValue(mPlanList[i][6]);
            superSheet.getRange("G"+(41+i)).setValue(mPlanList[i][7]);
            superSheet.getRange("H"+(41+i)).setValue(mPlanList[i][8]);
            superSheet.getRange("I"+(41+i)).setValue(mPlanList[i][9]);
        }
    }

    // 【基礎科目学習プラン】
    const basicStudyPlanSheet = SpreadsheetApp.openByUrl('ENTER YOUR URL');
    const bsPlan = basicStudyPlanSheet.getSheetByName('フォームの回答 1');
    const bsPlanLastRow = bsPlan.getLastRow();  // 基礎科目学習プランの最終行を取得
    const bsPlanRangeValues = bsPlan.getRange(2,1,bsPlanLastRow,7).getValues();// 基礎科目学習プランを二次元配列化
    // 基礎科目学習プランのうち対象の情報を抜き出し降順でリスト化
    var bsPlanList = [];
    for (var i=0; i<bsPlanRangeValues.length; i++){
        if (bsPlanRangeValues[i][1] == profile.mail){
            bsPlanList.unshift(bsPlanRangeValues[i]);
        }
    }
    // 基礎科目学習プランをセル入力
    if (bsPlanList.length<10){
        for (var i=0;i<bsPlanList.length;i++){
            superSheet.getRange("A"+(54+i)).setValue(bsPlanList[i][0]);
            superSheet.getRange("B"+(54+i)).setValue(bsPlanList[i][3]);
            superSheet.getRange("D"+(54+i)).setValue(bsPlanList[i][4]);
            superSheet.getRange("F"+(54+i)).setValue(bsPlanList[i][5]);
            superSheet.getRange("H"+(54+i)).setValue(bsPlanList[i][6]);
        }
    }else{
        for (var i=0;i<10;i++){
            superSheet.getRange("A"+(54+i)).setValue(bsPlanList[i][0]);
            superSheet.getRange("B"+(54+i)).setValue(bsPlanList[i][3]);
            superSheet.getRange("D"+(54+i)).setValue(bsPlanList[i][4]);
            superSheet.getRange("F"+(54+i)).setValue(bsPlanList[i][5]);
            superSheet.getRange("H"+(54+i)).setValue(bsPlanList[i][6]);
        }
    }

    // 【PJS】
    const pjs = SpreadsheetApp.openByUrl(profile.pjsUrl);
    // 【目標】
    const pjsNumeric = pjs.getSheetByName('②目標');
    const pjsNumericGraduate = pjsNumeric.getRange("C5").getValue();
    const pjsNumeric1 = pjsNumeric.getRange("I11").getValue();
    const pjsNumeric2 = pjsNumeric.getRange("L11").getValue();
    const pjsNumeric3 = pjsNumeric.getRange("O11").getValue();
    superSheet.getRange("B13").setValue(pjsNumericGraduate);
    superSheet.getRange("B14").setValue(pjsNumeric1);
    superSheet.getRange("B15").setValue(pjsNumeric2);
    superSheet.getRange("B16").setValue(pjsNumeric3);


    // 【面談記録】
    const pjsInterview = pjs.getSheetByName('④面談記録');
    const pjsInterviewRangeValues = pjsInterview.getRange(9,2,50,3).getValues();// 二次元配列化
    // 降順ソート
    let pjsInterviewRangeValuesSort = []
    for(var i=0;i<50;i++){
        if (pjsInterviewRangeValues[i][0]!=''){
            pjsInterviewRangeValuesSort.unshift(pjsInterviewRangeValues[i]);
        }
    }
    // 面談記録をセルに入力
    if (pjsInterviewRangeValuesSort.length<3){
        for (var i=0;i<pjsInterviewRangeValuesSort.length;i++){
            superSheet.getRange("A"+(67+i)).setValue(pjsInterviewRangeValuesSort[i][0]);
            superSheet.getRange("B"+(67+i)).setValue(pjsInterviewRangeValuesSort[i][1]);
            superSheet.getRange("C"+(67+i)).setValue(pjsInterviewRangeValuesSort[i][2]);
        }
    }else{
        for (var i=0;i<3;i++){
            superSheet.getRange("A"+(67+i)).setValue(pjsInterviewRangeValuesSort[i][0]);
            superSheet.getRange("B"+(67+i)).setValue(pjsInterviewRangeValuesSort[i][1]);
            superSheet.getRange("C"+(67+i)).setValue(pjsInterviewRangeValuesSort[i][2]);
        }
    }




}
