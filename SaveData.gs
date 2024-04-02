function saveMission() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("WTInput");
  var recordSheet = ss.getSheetByName("WTRecord");

  var date = inputSheet.getRange("B3").getValue();
  var zone = inputSheet.getRange("A6").getValue();

  var rowRange = recordSheet.getDataRange();
  var values = rowRange.getValues();
  var numRows = values.length;

  var criteriaColumnIndex = [0, 1];
  var criteria = [date, zone];

  var response = Browser.msgBox(
    "ยืนยันการบันทึกข้อมูลเวลาทำงานใช่หรือไม่?",
    "คุณต้องการบันทึกข้อมูลเวลาทำงานของ Mission Zone ของวันที่: " + Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    Browser.Buttons.YES_NO
  );

  if (response == "yes") {
    // Delete existing rows with the same date and zone
    for (var i = numRows - 1; i >= 0; i--) {
      if (
        values[i][criteriaColumnIndex[0]].toString() == criteria[0].toString() &&
        values[i][criteriaColumnIndex[1]].toString() == criteria[1].toString()
      ) {
        recordSheet.deleteRow(i + 1); // Rows are 1-indexed
      }
    }

    // Append new data
    var week = inputSheet.getRange("E3").getValue();
    var startTime = inputSheet.getRange("B6").getValue();
    var endTime = inputSheet.getRange("C6").getValue();
    var _4S = inputSheet.getRange("D6").getValue();
    var ky4 = inputSheet.getRange("E6").getValue();
    var wm = inputSheet.getRange("F6").getValue();
    var sgps = inputSheet.getRange("G6").getValue();
    var otp = inputSheet.getRange("H6").getValue();
    var otr = inputSheet.getRange("I6").getValue();
    var trmm = inputSheet.getRange("J6").getValue();
    var safetyMeeting = inputSheet.getRange("K6").getValue();
    var quarterlyMeeting = inputSheet.getRange("L6").getValue();

    var mission = ["PK1", "PK2", "PK3", "MSG1", "MSG2", "SFE", "MSC", "TMC", "CH", "MAM1", "MAM2", "FA"];

    for (var i = 0; i < mission.length; i++) {
      recordSheet.appendRow([date,zone, week, mission[i], startTime, endTime, _4S, ky4, wm, sgps, otp, otr, trmm, safetyMeeting, quarterlyMeeting]);
    }

    // Inform user that data has been saved
    Browser.msgBox("เสร็จสมบูรณ์", "บันทึกข้อมูลเรียบร้อย!", Browser.Buttons.OK);
  }
}

function saveFinal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("WTInput");
  var recordSheet = ss.getSheetByName("WTRecord");

  var date = inputSheet.getRange("B3").getValue();
  var zone = inputSheet.getRange("A8").getValue();

  var rowRange = recordSheet.getDataRange();
  var values = rowRange.getValues();
  var numRows = values.length;

  var criteriaColumnIndex = [0, 1];
  var criteria = [date, zone];

  var response = Browser.msgBox(
    "ยืนยันการบันทึกข้อมูลเวลาทำงานใช่หรือไม่?",
    "คุณต้องการบันทึกข้อมูลเวลาทำงานของ Final Zone ของวันที่: " + Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    Browser.Buttons.YES_NO
  );

  if (response == "yes") {
    // Delete existing rows with the same date and zone
    for (var i = numRows - 1; i >= 0; i--) {
      if (
        values[i][criteriaColumnIndex[0]].toString() == criteria[0].toString() &&
        values[i][criteriaColumnIndex[1]].toString() == criteria[1].toString()
      ) {
        recordSheet.deleteRow(i + 1); // Rows are 1-indexed
      }
    }

    // Append new data
    var week = inputSheet.getRange("E3").getValue();
    var startTime = inputSheet.getRange("B8").getValue();
    var endTime = inputSheet.getRange("C8").getValue();
    var _4S = inputSheet.getRange("D8").getValue();
    var ky4 = inputSheet.getRange("E8").getValue();
    var wm = inputSheet.getRange("F8").getValue();
    var sgps = inputSheet.getRange("G8").getValue();
    var otp = inputSheet.getRange("H8").getValue();
    var otr = inputSheet.getRange("I8").getValue();
    var trmm = inputSheet.getRange("J8").getValue();
    var safetyMeeting = inputSheet.getRange("K8").getValue();
    var quarterlyMeeting = inputSheet.getRange("L8").getValue();

    var final = ["PAP", "PAT", "FBP", "FAF", "FAR1", "FAR2", "FSP", "FSA", "FIF", "FIR"];

    for (var i = 0; i < final.length; i++) {
      recordSheet.appendRow([date, zone, week, final[i], startTime, endTime, _4S, ky4, wm, sgps, otp, otr, trmm, safetyMeeting, quarterlyMeeting]);
    }

    // Inform user that data has been saved
    Browser.msgBox("เสร็จสมบูรณ์", "บันทึกข้อมูลเรียบร้อย!", Browser.Buttons.OK);
  }
}

function saveTransmission() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("WTInput");
  var recordSheet = ss.getSheetByName("WTRecord");

  var date = inputSheet.getRange("B3").getValue();
  var zone = inputSheet.getRange("A7").getValue();

  var rowRange = recordSheet.getDataRange();
  var values = rowRange.getValues();
  var numRows = values.length;

  var criteriaColumnIndex = [0, 1];
  var criteria = [date, zone];

  var response = Browser.msgBox(
    "ยืนยันการบันทึกข้อมูลเวลาทำงานใช่หรือไม่?",
    "คุณต้องการบันทึกข้อมูลเวลาทำงานของ Transmission Zone ของวันที่: " + Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    Browser.Buttons.YES_NO
  );

  if (response == "yes") {
    // Delete existing rows with the same date and zone
    for (var i = numRows - 1; i >= 0; i--) {
      if (
        values[i][criteriaColumnIndex[0]].toString() == criteria[0].toString() &&
        values[i][criteriaColumnIndex[1]].toString() == criteria[1].toString()
      ) {
        recordSheet.deleteRow(i + 1); // Rows are 1-indexed
      }
    }

    // Append new data
    var week = inputSheet.getRange("E3").getValue();
    var startTime = inputSheet.getRange("B7").getValue();
    var endTime = inputSheet.getRange("C7").getValue();
    var _4S = inputSheet.getRange("D7").getValue();
    var ky4 = inputSheet.getRange("E7").getValue();
    var wm = inputSheet.getRange("F7").getValue();
    var sgps = inputSheet.getRange("G7").getValue();
    var otp = inputSheet.getRange("H7").getValue();
    var otr = inputSheet.getRange("I7").getValue();
    var trmm = inputSheet.getRange("J7").getValue();
    var safetyMeeting = inputSheet.getRange("K7").getValue();
    var quarterlyMeeting = inputSheet.getRange("L7").getValue();

    var transmission = ["TMF", "TMR", "TMP"];

    for (var i = 0; i < transmission.length; i++) {
      recordSheet.appendRow([date, zone, week, transmission[i], startTime, endTime, _4S, ky4, wm, sgps, otp, otr, trmm, safetyMeeting, quarterlyMeeting]);
    }

    // Inform user that data has been saved
    Browser.msgBox("เสร็จสมบูรณ์", "บันทึกข้อมูลเรียบร้อย!", Browser.Buttons.OK);
  }
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// function saveMission() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var inputSheet = ss.getSheetByName("WTInput");
//   var recordSheet = ss.getSheetByName("WTRecord");

//   var date = inputSheet.getRange("B3").getValue();

//   var rowRange = recordSheet.getDataRange();
//   var values = rowRange.getValues();
//   var numRows = values.length;

//   var criteriaColumnIndex = 0;
//   var criteria = date;

//   var response = Browser.msgBox(
//     "Confirm Data Save",
//     "Do you want to save data for the date: " + Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy"),
//     Browser.Buttons.YES_NO
//   );

//   if (response == "yes") {
//     // Delete existing rows with the same date
//     for (var i = numRows - 1; i >= 0; i--) {
//       if (values[i][criteriaColumnIndex].toString() == criteria.toString()) {
//         recordSheet.deleteRow(i + 1); // Rows are 1-indexed
//       }
//     }

//     // Append new data
//     var week = inputSheet.getRange("E3").getValue();
//     var startTime = [inputSheet.getRange("B6").getValue(), inputSheet.getRange("B7").getValue(), inputSheet.getRange("B8").getValue()];
//     var endTime = [inputSheet.getRange("C6").getValue(), inputSheet.getRange("C7").getValue(), inputSheet.getRange("C8").getValue()];
//     var taktTime = [inputSheet.getRange("D6").getValue(), inputSheet.getRange("D7").getValue(), inputSheet.getRange("D8").getValue()];
//     var safetyMeeting = [inputSheet.getRange("E6").getValue(), inputSheet.getRange("E7").getValue(), inputSheet.getRange("E8").getValue()];
//     var trmm = [inputSheet.getRange("F6").getValue(), inputSheet.getRange("F7").getValue(), inputSheet.getRange("F8").getValue()];
//     var _4S = [inputSheet.getRange("G6").getValue(), inputSheet.getRange("G7").getValue(), inputSheet.getRange("G8").getValue()];
//     var ky4 = [inputSheet.getRange("H6").getValue(), inputSheet.getRange("H7").getValue(), inputSheet.getRange("H8").getValue()];
//     var wm = [inputSheet.getRange("I6").getValue(), inputSheet.getRange("I7").getValue(), inputSheet.getRange("I8").getValue()];
//     var sgps = [inputSheet.getRange("J6").getValue(), inputSheet.getRange("J7").getValue(), inputSheet.getRange("J8").getValue()];
//     var otp = [inputSheet.getRange("K6").getValue(), inputSheet.getRange("K7").getValue(), inputSheet.getRange("K8").getValue()];
//     var otr = [inputSheet.getRange("L6").getValue(), inputSheet.getRange("L7").getValue(), inputSheet.getRange("L8").getValue()];

//     var mission = ["PK1", "PK2", "PK3", "MSG1", "MSG2", "SFE", "MSC", "TMC", "CH", "MAM1", "MAM2", "FA"];
//     var transmission = ["TMF", "TMR", "TMP"];
//     var final = ["PAP", "PAT", "FBP", "FAF", "FAR1", "FAR2", "FSP", "FSA", "FIF", "FIR"];

//     for (var i = 0; i < mission.length; i++) {
//       recordSheet.appendRow([date, week, mission[i], startTime[0], endTime[0], taktTime[0], safetyMeeting[0], trmm[0], _4S[0], ky4[0], wm[0], sgps[0], otp[0], otr[0]]);
//     }

//     for (var i = 0; i < transmission.length; i++) {
//       recordSheet.appendRow([date, week, transmission[i], startTime[1], endTime[1], taktTime[1], safetyMeeting[1], trmm[1], _4S[1], ky4[1], wm[1], sgps[1], otp[1], otr[1]]);
//     }

//     for (var i = 0; i < final.length; i++) {
//       recordSheet.appendRow([date, week, final[i], startTime[2], endTime[2], taktTime[2], safetyMeeting[2], trmm[2], _4S[2], ky4[2], wm[2], sgps[2], otp[2], otr[2]]);
//     }

//     // Inform user that data has been saved
//     Browser.msgBox("Finished", "Data has been saved successfully!", Browser.Buttons.OK);
//   }
// }
//------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function saveProduction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // destination sheet
  var recordSheet = ss.getSheetByName("ProdRecord");
  var inputSheet = ss.getSheetByName("ProdInput");

  // collect the data
  var inputRange = [ss.getRangeByName("Prod_PK1"), ss.getRangeByName("Prod_PK2"), ss.getRangeByName("Prod_PK3"), ss.getRangeByName("Prod_MSG1"), ss.getRangeByName("Prod_MSG2"),
                    ss.getRangeByName("Prod_SFE"), ss.getRangeByName("Prod_MSC"), ss.getRangeByName("Prod_TMC"), ss.getRangeByName("Prod_CH"), ss.getRangeByName("Prod_MAM1"),
                    ss.getRangeByName("Prod_FA"), ss.getRangeByName("Prod_TMF"), ss.getRangeByName("Prod_TMR"), ss.getRangeByName("Prod_TMP"), ss.getRangeByName("Prod_PAP"),
                    ss.getRangeByName("Prod_PAT"), ss.getRangeByName("Prod_FBP"), ss.getRangeByName("Prod_FAF"), ss.getRangeByName("Prod_FAR1"), ss.getRangeByName("Prod_FSP"),
                    ss.getRangeByName("Prod_FSA"), ss.getRangeByName("Prod_FIR"), ss.getRangeByName("Prod_BM"), ss.getRangeByName("Prod_BF"), ss.getRangeByName("Prod_PB"),
                    ss.getRangeByName("Prod_WD"), ss.getRangeByName("Prod_MC"),];
  var inputVals = [inputRange[0].getValues().flat(), inputRange[1].getValues().flat(), inputRange[2].getValues().flat(), inputRange[3].getValues().flat(),inputRange[4].getValues().flat(), 
                  inputRange[5].getValues().flat(), inputRange[6].getValues().flat(), inputRange[7].getValues().flat(), inputRange[8].getValues().flat(), inputRange[9].getValues().flat(), 
                  inputRange[10].getValues().flat(), inputRange[11].getValues().flat(), inputRange[12].getValues().flat(), inputRange[13].getValues().flat(), inputRange[14].getValues().flat(), inputRange[15].getValues().flat(),inputRange[16].getValues().flat(), inputRange[17].getValues().flat(), inputRange[18].getValues().flat(), inputRange[19].getValues().flat(),
                  inputRange[20].getValues().flat(), inputRange[21].getValues().flat(), inputRange[22].getValues().flat(), inputRange[23].getValues().flat(), inputRange[24].getValues().flat(),
                  inputRange[25].getValues().flat(), inputRange[26].getValues().flat()];

  //console.log(inputVals[0][0]);
  var rowRange = recordSheet.getDataRange();
  var values = rowRange.getValues();
  var numRows = values.length;

  var criteriaColumnIndex = [0, 1];
  var criteriaLine = [inputSheet.getRange("A5").getValue(),inputSheet.getRange("A6").getValue(),inputSheet.getRange("A7").getValue(),inputSheet.getRange("A8").getValue(),
                      inputSheet.getRange("A9").getValue(),inputSheet.getRange("A10").getValue(),inputSheet.getRange("A11").getValue(),inputSheet.getRange("A12").getValue(),
                      inputSheet.getRange("A13").getValue(),inputSheet.getRange("A14").getValue(),inputSheet.getRange("E5").getValue(),inputSheet.getRange("E6").getValue(),
                      inputSheet.getRange("E7").getValue(),inputSheet.getRange("E8").getValue(),inputSheet.getRange("E9").getValue(),inputSheet.getRange("E10").getValue(),
                      inputSheet.getRange("E11").getValue(),inputSheet.getRange("E12").getValue(),inputSheet.getRange("E13").getValue(),inputSheet.getRange("E14").getValue(),
                      inputSheet.getRange("I5").getValue(),inputSheet.getRange("I6").getValue(),inputSheet.getRange("I7").getValue(),inputSheet.getRange("I8").getValue(),
                      inputSheet.getRange("I9").getValue(),inputSheet.getRange("I10").getValue(),inputSheet.getRange("I11").getValue()];

  var criteriaDate = [inputSheet.getRange("B5").getValue(),inputSheet.getRange("B6").getValue(),inputSheet.getRange("B7").getValue(),inputSheet.getRange("B8").getValue(),
                      inputSheet.getRange("B9").getValue(),inputSheet.getRange("B10").getValue(),inputSheet.getRange("B11").getValue(),inputSheet.getRange("B12").getValue(),
                      inputSheet.getRange("B13").getValue(),inputSheet.getRange("B14").getValue(),inputSheet.getRange("F5").getValue(),inputSheet.getRange("F6").getValue(),
                      inputSheet.getRange("F7").getValue(),inputSheet.getRange("F8").getValue(),inputSheet.getRange("F9").getValue(),inputSheet.getRange("F10").getValue(),
                      inputSheet.getRange("F11").getValue(),inputSheet.getRange("F12").getValue(),inputSheet.getRange("F13").getValue(),inputSheet.getRange("F14").getValue(),
                      inputSheet.getRange("J5").getValue(),inputSheet.getRange("J6").getValue(),inputSheet.getRange("J7").getValue(),inputSheet.getRange("J8").getValue(),
                      inputSheet.getRange("J9").getValue(),inputSheet.getRange("J10").getValue(),inputSheet.getRange("J11").getValue()];

  // var inputClear = [inputSheet.getRange("B5:D5"), inputSheet.getRange("B6:D6"), inputSheet.getRange("B7:D7"), inputSheet.getRange("B8:D8"), inputSheet.getRange("B9:D9"),
  //                   inputSheet.getRange("B10:D10"), inputSheet.getRange("B11:D11"), inputSheet.getRange("B12:D12"), inputSheet.getRange("B13:D13"), inputSheet.getRange("B14:D14"),
  //                   inputSheet.getRange("F5:H5"), inputSheet.getRange("F6:H6"), inputSheet.getRange("F7:H7"), inputSheet.getRange("F8:H8"), inputSheet.getRange("F9:H9"),
  //                   inputSheet.getRange("F10:H10"), inputSheet.getRange("F11:H11"), inputSheet.getRange("F12:H12"), inputSheet.getRange("F13:H13"), inputSheet.getRange("F14:H14"),
  //                   inputSheet.getRange("J5:L5"), inputSheet.getRange("J6:L6"), inputSheet.getRange("J7:L7"), inputSheet.getRange("J8:L8"), inputSheet.getRange("J9:L9"),
  //                   inputSheet.getRange("J10:L10"), inputSheet.getRange("J11:L11"),]

  var inputClear = [inputSheet.getRange("C5:D5"), inputSheet.getRange("C6:D6"), inputSheet.getRange("C7:D7"), inputSheet.getRange("C8:D8"), inputSheet.getRange("C9:D9"),
                    inputSheet.getRange("C10:D10"), inputSheet.getRange("C11:D11"), inputSheet.getRange("C12:D12"), inputSheet.getRange("C13:D13"), inputSheet.getRange("C14:D14"),
                    inputSheet.getRange("G5:H5"), inputSheet.getRange("G6:H6"), inputSheet.getRange("G7:H7"), inputSheet.getRange("G8:H8"), inputSheet.getRange("G9:H9"),
                    inputSheet.getRange("G10:H10"), inputSheet.getRange("G11:H11"), inputSheet.getRange("G12:H12"), inputSheet.getRange("G13:H13"), inputSheet.getRange("G14:H14"),
                    inputSheet.getRange("K5:L5"), inputSheet.getRange("K6:L6"), inputSheet.getRange("K7:L7"), inputSheet.getRange("K8:L8"), inputSheet.getRange("K9:L9"),
                    inputSheet.getRange("K10:L10"), inputSheet.getRange("K11:L11")]
  var dateRange = [inputSheet.getRange("B5"), inputSheet.getRange("B6"), inputSheet.getRange("B7"), inputSheet.getRange("B8"), inputSheet.getRange("B9"),
                  inputSheet.getRange("B10"), inputSheet.getRange("B11"), inputSheet.getRange("B12"), inputSheet.getRange("B13"), inputSheet.getRange("B14"),
                  inputSheet.getRange("F5"), inputSheet.getRange("F6"), inputSheet.getRange("F7"), inputSheet.getRange("F8"), inputSheet.getRange("F9"),
                  inputSheet.getRange("F10"), inputSheet.getRange("F11"), inputSheet.getRange("F12"), inputSheet.getRange("F13"), inputSheet.getRange("F14"),
                  inputSheet.getRange("J5"), inputSheet.getRange("J6"), inputSheet.getRange("J7"), inputSheet.getRange("J8"), inputSheet.getRange("J9"),
                  inputSheet.getRange("J10"), inputSheet.getRange("J11"),]
  var dateRangeValue = []
  for (var i=0; i<dateRange.length;i++) {
    dateRangeValue.push(dateRange[i].getValue())
  }

  var response = Browser.msgBox(
    "ยืนยันการบันทึกข้อมูลยอดผลิตใช่หรือไม่?",
    Browser.Buttons.YES_NO
  );

  if (response == "yes") {
    // Delete existing rows with the same line and date
    for (var i = numRows - 1; i >= 0; i--) {
      for (var j = 0; j < criteriaLine.length; j++) {
        if (
          values[i][criteriaColumnIndex[0]].toString() == criteriaLine[j].toString() &&
          values[i][criteriaColumnIndex[1]].toString() == criteriaDate[j].toString()
        ) {
          recordSheet.deleteRow(i + 1); // Rows are 1-indexed
        }
      }
    }

    for (var i = 0; i < inputVals.length; i++) {
      recordSheet.appendRow(inputVals[i])
      dateRangeValue[i].setDate(dateRangeValue[i].getDate() + 1)
      dateRange[i].setValue(dateRangeValue[i])
      inputClear[i].clearContent()
    }
    

    // Inform user that data has been saved
    Browser.msgBox("Finished", "Data has been saved successfully!", Browser.Buttons.OK);  
  }

}

function saveProductionSpecial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // destination sheet
  var recordSheet = ss.getSheetByName("ProdRecordS");
  var inputSheet = ss.getSheetByName("ProdInput");

  var lineNames = [inputSheet.getRange("A17").getValue(), inputSheet.getRange("A17").getValue(), inputSheet.getRange("A17").getValue(),
                  inputSheet.getRange("A21").getValue(),inputSheet.getRange("A21").getValue(),inputSheet.getRange("A21").getValue(),inputSheet.getRange("A21").getValue(),
                  inputSheet.getRange("E21").getValue(),inputSheet.getRange("E21").getValue(),inputSheet.getRange("E21").getValue(),inputSheet.getRange("E21").getValue(),
                  inputSheet.getRange("I21").getValue(),inputSheet.getRange("I21").getValue(),inputSheet.getRange("I21").getValue(),inputSheet.getRange("I21").getValue()];
  var weeks = [inputSheet.getRange("A19").getValue(), inputSheet.getRange("A19").getValue(), inputSheet.getRange("A19").getValue(),
                  inputSheet.getRange("A24").getValue(),inputSheet.getRange("A24").getValue(),inputSheet.getRange("A24").getValue(),inputSheet.getRange("A24").getValue(),
                  inputSheet.getRange("E24").getValue(),inputSheet.getRange("E24").getValue(),inputSheet.getRange("E24").getValue(),inputSheet.getRange("E24").getValue(),
                  inputSheet.getRange("I24").getValue(),inputSheet.getRange("I24").getValue(),inputSheet.getRange("I24").getValue(),inputSheet.getRange("I24").getValue()];                
  var dates = [inputSheet.getRange("B17:B19").getValue(), inputSheet.getRange("B17:B19").getValue(), inputSheet.getRange("B17:B19").getValue(),
              inputSheet.getRange("B21:B24").getValue(), inputSheet.getRange("B21:B24").getValue(), inputSheet.getRange("B21:B24").getValue(), inputSheet.getRange("B21:B24").getValue(),
              inputSheet.getRange("F21:F24").getValue(), inputSheet.getRange("F21:F24").getValue(), inputSheet.getRange("F21:F24").getValue(), inputSheet.getRange("F21:F24").getValue(),
              inputSheet.getRange("J21:J24").getValue(), inputSheet.getRange("J21:J24").getValue(), inputSheet.getRange("J21:J24").getValue(), inputSheet.getRange("J21:J24").getValue()];

  var inputRange = [ss.getRangeByName("Prod_MAM2_L"), ss.getRangeByName("Prod_MAM2_MU"), ss.getRangeByName("Prod_MAM2_M"),
                    ss.getRangeByName("Prod_FAR2_L"),ss.getRangeByName("Prod_FAR2_MU"),ss.getRangeByName("Prod_FAR2_CABIN"),ss.getRangeByName("Prod_FAR2_M"),
                    ss.getRangeByName("Prod_CABIN_L"),ss.getRangeByName("Prod_CABIN_MU"),ss.getRangeByName("Prod_CABIN_CABIN"),ss.getRangeByName("Prod_CABIN_M"),
                    ss.getRangeByName("Prod_FIF_L"),ss.getRangeByName("Prod_FIF_MU"),ss.getRangeByName("Prod_FIF_CABIN"),ss.getRangeByName("Prod_FIF_M")];
  var inputVals = [inputRange[0].getValues().flat(), inputRange[1].getValues().flat(), inputRange[2].getValues().flat(),
                  inputRange[3].getValues().flat(), inputRange[4].getValues().flat(), inputRange[5].getValues().flat(), inputRange[6].getValues().flat(),
                  inputRange[7].getValues().flat(), inputRange[8].getValues().flat(), inputRange[9].getValues().flat(), inputRange[10].getValues().flat(),
                  inputRange[11].getValues().flat(), inputRange[12].getValues().flat(), inputRange[13].getValues().flat(), inputRange[14].getValues().flat()];
  
  var inputClearAct = [inputSheet.getRange("D17"), inputSheet.getRange("D18"), inputSheet.getRange("D19"),
                      inputSheet.getRange("D21"), inputSheet.getRange("D22"), inputSheet.getRange("D23"), inputSheet.getRange("D24"),
                      inputSheet.getRange("H21"), inputSheet.getRange("H22"), inputSheet.getRange("H23"), inputSheet.getRange("H24"),
                      inputSheet.getRange("L21"), inputSheet.getRange("L22"), inputSheet.getRange("L23"), inputSheet.getRange("L24")];

  var inputClearWeek = [inputSheet.getRange("A19"), inputSheet.getRange("A24"), inputSheet.getRange("E24"), inputSheet.getRange("I24")]

  var dateRange = [inputSheet.getRange("B17:B19"), inputSheet.getRange("B21:B24"), inputSheet.getRange("F21:F24"), inputSheet.getRange("J21:J24")]
  var dateRangeValue = []
  for (var i=0; i<dateRange.length;i++) {
    dateRangeValue.push(dateRange[i].getValue())
  }

  var rowRange = recordSheet.getDataRange();
  var values = rowRange.getValues();
  var numRows = values.length;

  var criteriaColumnIndex = [0, 2, 3];
  var criteriaModel = [inputSheet.getRange("C17").getValue(), inputSheet.getRange("C18").getValue(), inputSheet.getRange("C19").getValue(),
                      inputSheet.getRange("C21").getValue(), inputSheet.getRange("C22").getValue(), inputSheet.getRange("C23").getValue(), inputSheet.getRange("C24").getValue(),
                      inputSheet.getRange("G21").getValue(), inputSheet.getRange("G22").getValue(), inputSheet.getRange("G23").getValue(), inputSheet.getRange("G24").getValue(),
                      inputSheet.getRange("K21").getValue(), inputSheet.getRange("K22").getValue(), inputSheet.getRange("K23").getValue(), inputSheet.getRange("K24").getValue()];



  dataAll = [];
  //data = [dates[0], weeks[0], lineNames[0], ...inputVals[0]]
  for (var i = 0; i < inputVals.length; i++) {
    data = [dates[i], weeks[i], lineNames[i], ...inputVals[i]]
    dataAll.push(data)
  }

  for (var i = numRows - 1; i >= 0; i--) {
    for (var j = 0; j < criteriaModel.length; j++) {
      if (
        values[i][criteriaColumnIndex[0]].toString() == dates[j].toString() &&
        values[i][criteriaColumnIndex[1]].toString() == lineNames[j].toString() &&
        values[i][criteriaColumnIndex[2]].toString() == criteriaModel[j].toString()
      ) {
        recordSheet.deleteRow(i + 1); // Rows are 1-indexed
      }
    }
  }

  for (var i = 0; i < inputVals.length; i++) {
      recordSheet.appendRow(dataAll[i])
      inputClearAct[i].clearContent()
    }

  for (var i = 0; i < inputClearWeek.length; i++) {
    inputClearWeek[i].clearContent()
  }

  for (var i = 0; i < dateRangeValue.length; i++) {
    dateRangeValue[i].setDate(dateRangeValue[i].getDate() + 1)
    dateRange[i].setValue(dateRangeValue[i])
  }

  //Logger.log(values[1][criteriaColumnIndex[0]].toString()==dates[1].toString() && values[1][criteriaColumnIndex[2]].toString() == lineNames[1].toString())
}

//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function savePlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // destination sheet
  var recordSheet = ss.getSheetByName("recordPlan");
  var inputSheet = ss.getSheetByName("InputPlan");

  // Create day list to record
  var dayRange = ss.getRangeByName("dayList").getValues();
  var dayRangeList = [];
  for(var i=0; i< dayRange.length; i++) {
    for (var j = 0; j < dayRange[i].length; j++) {
      dayRangeList.push(dayRange[i][j]);
    }
  }

  // Create date list to record
  var dateList = [];
  var month = inputSheet.getRange("G2").getValue()-1;
  var year = inputSheet.getRange("I2").getValue();
  for (var i = 0; i<dayRangeList.length; i++) {
    dateValue = new Date(year,month,dayRangeList[i]);
    dateList.push(dateValue);
  }

  // Create model list to record
  var modelRange = ss.getRangeByName("modelList").getValues();
  var modelRangeList = [];
  for(var i=0; i< modelRange.length; i++) {
    for (var j = 0; j < modelRange[i].length; j++) {
      if (modelRange[i][j] !== null && modelRange[i][j] !== "") {
        modelRangeList.push(modelRange[i][j]);
      }
    }
  }

  // Create Production list to record
  var prodDayRange = [];
  var prodPlanByDayList = [[],[],[],[],[],[],[],[],[],[],
                            [],[],[],[],[],[],[],[],[],[],
                            [],[],[],[],[],[],[],[],[],[],
                            []];

  // Create List of data to record [[date, model, prod]]
  var dataList = [];

  // Create criteria variable for delete duplicated value
  var dateCriteria = dateList;
  var modelGroupCriteria = modelRangeList;

  var rowRange = recordSheet.getDataRange();
  var values = rowRange.getValues();
  var numRows = values.length;

  var criteriaColumnIndex = [0, 1];

  var date = new Date(year,month,1)
  var response = Browser.msgBox(
    "ยืนยันการบันทึกแผนผลิตใช่หรือไม่?",
    "คุณต้องการบันทึกแผนผลิตของเดือน: " + Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM-yyyy"),
    Browser.Buttons.YES_NO
  );

  if (response == "yes") {
    // Delete existing rows with the same date and zone
    for (var i = numRows - 1; i >= 0; i--) {
      for (var j = 0; j < dateCriteria.length; j++) {
        for (var k = 0; k < modelGroupCriteria.length; k++) {
          if (
            values[i][criteriaColumnIndex[0]].toString() == dateCriteria[j].toString() &&
            values[i][criteriaColumnIndex[1]].toString() == modelGroupCriteria[k].toString()
          ) {
            recordSheet.deleteRow(i + 1); // Rows are 1-indexed
          }
        }
      }
    }

    // Create Production list to record
    for (i=1; i<=dateList.length; i++) {
      var range_name = "day_" + i;
      prodDayRange.push(ss.getRangeByName(range_name).getValues().flat())
    }
    for (i=0; i<prodDayRange.length; i++) {
      for (j=0; j<modelRangeList.length; j++) {
        if (prodDayRange[i][j] !== null && prodDayRange[i][j] !== "") {
          prodPlanByDayList[i].push(prodDayRange[i][j]);
        } else {
          prodPlanByDayList[i].push(0);
        }
      }
    }

    // Create List of data to record [[date, model, prod]]
    for (var i=0; i<dateList.length; i++) {
      for (var j=0; j<modelRangeList.length; j++) {
          data = [dateList[i], modelRangeList[j], prodPlanByDayList[i][j]];
          //Logger.log(data)
          dataList.push(data)
      }
    }
    
    // Append to recordSheet
    for (var i=0; i<dataList.length; i++) {
      recordSheet.appendRow(dataList[i])
    }

    // data = dateList[30] + " " + modelRangeList[75] + " " + prodPlanByDayList[30][75];

    // Logger.log(dataList.length);

    // Inform user that data has been saved
    Browser.msgBox("Finished", "Data has been saved successfully!", Browser.Buttons.OK);  
  }
}
