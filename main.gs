const configEnum = {
  mapSheetName: "對照名單", // 要比對的表單名稱
  staffSheetName: "SlackID清單(勿動)", // staff sheet name
  channelId: "", // channel id
  slackBotToken: "", // slack bot token
  slackBotName: "Robot", // slack bot name
  configSheetName: "SlackBot設定檔" // config sheet name
};

// 手動檢查名單
function ManualCheckList() {
  const btnUI = SpreadsheetApp.getUi();
  // 避免誤觸，先詢問確認
  const confirmRes = btnUI.alert(
     "警告",
     "即將發訊息到正式群組，確定繼續執行此操作嗎？",
      btnUI.ButtonSet.YES_NO);
  if (confirmRes == btnUI.Button.NO) return;
  CheckList();
}

// 檢查未填體溫名單
function CheckList() {
  const today = Utilities.formatDate(new Date(), "GMT+8", "MM/dd"); // get today
  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.configSheetName); // 設定檔表單
  const lastRowNum = settingConfigSheet.getLastRow(); // 取得最後一列num
  const completeMsgTextList = settingConfigSheet.getRange(4, 2, (lastRowNum - 3), 1).getValues().filter(e => e[0] !== "").map(e => e[0]); // 取得全數完成提示訊息
  const unDoneMsgTextList = settingConfigSheet.getRange(4, 3, (lastRowNum - 3), 1).getValues().filter(e => e[0] !== "").map(e => e[0]); // 取得未完成提示訊息
  
  // 取得忘記填名單
  const forgotList = GetForgotList();
  const completed = (forgotList.length === 0);
  let text = "";
  switch (completed) {
    case true: {
      /* 全數完成 */
      text = `${today} ${completeMsgTextList.join("\n")}`;
      break;
    }
    case false: {
      /* 需通報 */
      // 取得員工名單
      const staffList = GetStaffList();
      if (staffList.length === 0) return;

      // map提醒名單
      let needTagUIdList = [];
      for (const [index, item] of forgotList.entries()) {
        const _memberObj = staffList.find(e => e.mail === item.mail);
        if (!_memberObj) continue;
        needTagUIdList.push(`<@${_memberObj.uId}>`);
        // SendMsgToMember(_memberObj.uId); // 順道發訊息提醒個人
      }
      text = `${today} ${unDoneMsgTextList.join("\n").replace("{needTagUIdList}", needTagUIdList.join(" "))}`;
    }
  }

  if (!text) return;
  PostMsgToSlackGroup(text); // send msg
}

// 手動提醒未填人員
function ManualRemind() {
  const btnUI = SpreadsheetApp.getUi();
  // 避免誤觸，先詢問確認
  const confirmRes = btnUI.alert(
     "警告",
     "即將檢查並發送訊息給未填寫人員",
      btnUI.ButtonSet.YES_NO);
  if (confirmRes == btnUI.Button.NO) return;
  RemindHandler();
}

// 提醒程序
function RemindHandler () {
  // 取得忘記填名單
  const forgotList = GetForgotList();
  const completed = (forgotList.length === 0);
  if (completed) return; // 已全數完成無需提醒人員

  /* 需要提醒人員處理程序 */
  // 取得員工名單
  const staffList = GetStaffList();
  if (staffList.length === 0) return;
  // map提醒名單
  for (const [index, item] of forgotList.entries()) {
    const _memberObj = staffList.find(e => e.mail === item.mail);
    if (!_memberObj) continue;
    SendMsgToMember(_memberObj.uId); // 發訊息提醒人員
  }
}


// 發送訊息給個人
function SendMsgToMember (uId) {
  // validate params
  if (!uId) return;

  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.configSheetName); // 設定檔表單
  const lastRowNum = settingConfigSheet.getLastRow(); // 取得最後一列num
  const remindMsgTextList = settingConfigSheet.getRange(4, 7, (lastRowNum - 3), 1).getValues().filter(e => e[0] !== "").map(e => e[0]); // 取得手動點擊提醒未填人員提示訊息

  const slackBotName = settingConfigSheet.getSheetValues(4, 1, 1, 1)[0][0]; // 取得機器人顯示名稱
  const text = `${remindMsgTextList.join("\n").replace("{needTagUId}", `<@${uId}>`)}`;
  const formData = {
    'token': configEnum.slackBotToken, // slack bot token
    'channel': uId, // slack member uId
    "username": slackBotName.trim() || configEnum.slackBotName,
    "text": text
  };
  const options = {
    'method': 'post',
    'payload': formData
  };
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
}


// 發佈訊息至群組
function PostMsgToSlackGroup (msgText) {
  // validate params
  if (!msgText) return;

  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.configSheetName); // 設定檔表單
  const slackBotName = settingConfigSheet.getSheetValues(4, 1, 1, 1)[0][0]; // 取得機器人顯示名稱

  const formData = {
    'token': configEnum.slackBotToken, // slack bot token
    'channel': configEnum.channelId, // slack channel id
    "username": slackBotName.trim() || configEnum.slackBotName,
    "text": msgText
  };
  const options = {
    'method': 'post',
    'payload': formData
  };
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
}


// 取得忘記填名單
function GetForgotList() {
  const mapSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.mapSheetName);
  const lastRow = mapSheet.getLastRow(); // 取得最後一列num
  const contentList = mapSheet.getRange(2, 1, (lastRow - 1), 6).getValues();

  // 取得員工名單
  const staffList = GetStaffList();
  if (staffList.length === 0) return [];

  // 忘記名單
  const forgotList = contentList.filter(e => {
    const isStaff = staffList.find(staff => staff.mail === e[1]);
    if (e[5] !== "V" && isStaff) return e; // 過濾非Staff人員
  }).map(e => {
    // e = [部門, 信箱, 姓名]
    return {
      mail: e[1],
      name: e[2]
    };
  });

  return forgotList || [];
}


// 取的Staff list
function GetStaffList() {
  const staffSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.staffSheetName); // Staff名單 Sheet 名稱
  const rows = staffSheet.getDataRange().getValues();
  if (rows.length === 0) return [];
  return rows.map(e => {
    return {
      name: e[0],
      mail: e[1],
      uId: e[2]
    };
  });
}


// 匯入Staff名單
function ImportStaffList() {
  const btnUI = SpreadsheetApp.getUi();
  // 避免誤觸，先詢問確認
  const confirmRes = btnUI.alert(
     "確定要匯入Staff名單？",
     "",
      btnUI.ButtonSet.YES_NO);
  if (confirmRes == btnUI.Button.NO) return;

  const formData = {
    'token': configEnum.slackBotToken, // slack bot token
    'channel': configEnum.channelId, // slack channel id
    'limit': 500
  };
  const options = {
    'method' : 'post',
    'payload' : formData
  };
  const response = UrlFetchApp.fetch('https://slack.com/api/conversations.members', options);
  const data = JSON.parse(response.getContentText());
  
  const staffSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.staffSheetName); // Staff名單 Sheet 名稱
  // 如果有member資料，就清空Staff名單sheet資料
  if (data["members"].length > 0) {
    staffSheet.clear(); // 清除員工資料
  }

  for (index in data["members"]) {
    const uid = data["members"][index];
    const formData = {
      'token': configEnum.slackBotToken, // slack bot token
      'user': uid // user id
    };
    const options = {
      'method' : 'post',
      'payload' : formData
    };
    const response = UrlFetchApp.fetch('https://slack.com/api/users.info', options);
    const user = JSON.parse(response.getContentText());
    const email = user["user"]["profile"]["email"]; // staff mail
    const real_name = user["user"]["profile"]["real_name_normalized"]; // staff name
    if (user["user"]["is_bot"]) continue; // 過濾機器人

    const row = [real_name, email, uid]; // 依照 Column 填入
    staffSheet.appendRow(row); // append data
  }
}

// 建立提前提醒填寫體溫排程
function CreateRemindTrigger() {
  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.configSheetName); // 設定檔表單
  const hour = settingConfigSheet.getSheetValues(4, 8, 1, 1)[0][0]; // 取得小時
  const minute = settingConfigSheet.getSheetValues(4, 9, 1, 1)[0][0]; // 取得分鐘
  const day = Number(settingConfigSheet.getSheetValues(4, 10, 1, 1)[0][0]); // 取得循環天數

  CreateTrigger("RemindHandler", hour, minute, day); // 建立檢查體溫登記排程
}

// 建立檢查體溫登記排程
function CreateCheckListTrigger() {
  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configEnum.configSheetName); // 設定檔表單
  const hour = settingConfigSheet.getSheetValues(4, 4, 1, 1)[0][0]; // 取得小時
  const minute = settingConfigSheet.getSheetValues(4, 5, 1, 1)[0][0]; // 取得分鐘
  const day = Number(settingConfigSheet.getSheetValues(4, 6, 1, 1)[0][0]); // 取得循環天數

  CreateTrigger("CheckList", hour, minute, day); // 建立檢查體溫登記排程
}

// 建立觸發條件
function CreateTrigger(fn, hour, minute, day) {
  if (!fn) return;
  const btnUI = SpreadsheetApp.getUi();
  
  // 判斷參數如果設定不符格式就擋掉
  if (hour === "" || minute === "" || day <= 0) {
    btnUI.alert("參數設定錯誤", "請重新審視參數設定！", btnUI.ButtonSet.OK);
    return;
  }

  // 避免誤觸，先詢問確認
  const confirmRes = btnUI.alert(
     "確定要建立新的排程設定？",
     "",
      btnUI.ButtonSet.YES_NO);
  if (confirmRes == btnUI.Button.NO) return;
  
  ScriptApp.newTrigger(fn)
  .timeBased()
  .atHour(Number(hour))
  .nearMinute(Number(minute))
  .everyDays(Number(day))
  .inTimezone("Asia/Taipei")
  .create();
}