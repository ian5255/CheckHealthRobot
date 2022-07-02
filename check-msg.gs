const checkMsgConfigEnum = {
  mapSheetName: "對照名單", // 要比對的表單名稱
  staffSheetName: "SlackID清單(勿動)", // staff sheet name
  channelId: "", // channel id
  slackBotToken: "", // slack bot token
  slackBotName: "測試訊息機器人", // slack bot name
  configSheetName: "SlackBot設定檔" // config sheet name
};

// 測試檢查名單訊息
function TestCheckListMsg() {
  const today = Utilities.formatDate(new Date(), "GMT+8", "MM/dd"); // get today
  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(checkMsgConfigEnum.configSheetName); // 設定檔表單
  const lastRowNum = settingConfigSheet.getLastRow(); // 取得最後一列num
  const completeMsgTextList = settingConfigSheet.getRange(4, 2, (lastRowNum - 3), 1).getValues().filter(e => e[0] !== "").map(e => e[0]); // 取得全數完成提示訊息
  const unDoneMsgTextList = settingConfigSheet.getRange(4, 3, (lastRowNum - 3), 1).getValues().filter(e => e[0] !== "").map(e => e[0]); // 取得未完成提示訊息

  const testMsgList = [`${today} ${completeMsgTextList.join("\n")}`, `============分隔線============`, `${today} ${unDoneMsgTextList.join("\n").replace("{needTagUIdList}", checkMsgConfigEnum.channelId)}`];
  TestPostMsgToSlackGroup(testMsgList); // send msg to group
}

// 測試提醒程序
function TestRemindMsg () {
  TestSendMsgToMember(checkMsgConfigEnum.channelId);
}

// 發送訊息給個人
function TestSendMsgToMember (uId) {
  // validate params
  if (!uId) return;

  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(checkMsgConfigEnum.configSheetName); // 設定檔表單
  const lastRowNum = settingConfigSheet.getLastRow(); // 取得最後一列num
  const remindMsgTextList = settingConfigSheet.getRange(4, 7, (lastRowNum - 3), 1).getValues().filter(e => e[0] !== "").map(e => e[0]); // 取得手動點擊提醒未填人員提示訊息

  const slackBotName = settingConfigSheet.getSheetValues(4, 1, 1, 1)[0][0]; // 取得機器人顯示名稱
  const text = `${remindMsgTextList.join("\n").replace("{needTagUId}", uId)}`;
  const formData = {
    'token': checkMsgConfigEnum.slackBotToken, // slack bot token
    'channel': uId, // slack member uId
    "username": slackBotName.trim() || checkMsgConfigEnum.slackBotName,
    "text": text
  };
  const options = {
    'method': 'post',
    'payload': formData
  };
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
}


// 發佈訊息至群組
function TestPostMsgToSlackGroup (testMsgList) {
  // validate params
  if (testMsgList.length === 0) return;

  const settingConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(checkMsgConfigEnum.configSheetName); // 設定檔表單
  const slackBotName = settingConfigSheet.getSheetValues(4, 1, 1, 1)[0][0]; // 取得機器人顯示名稱

  for (const text of testMsgList) {
    const formData = {
      'token': checkMsgConfigEnum.slackBotToken, // slack bot token
      'channel': checkMsgConfigEnum.channelId, // slack channel id
      "username": slackBotName.trim() || checkMsgConfigEnum.slackBotName,
      "text": text
    };
    const options = {
      'method': 'post',
      // "contentType": "application/x-www-form-urlencoded,application/json",
      'payload': formData
    };
    UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
  }
}