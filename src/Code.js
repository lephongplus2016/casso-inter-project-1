//hàm tạo Menu trên thanh công cụ
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
  .addItem('Intern 1.0.0', 'run')
  .addSeparator()
  .addItem('Form Sheet', 'formSheet').addToUi();
}

//hàm tạo mẫu gg sheet
function formSheet(){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var mySheet = activeSpreadsheet.getSheetByName("Values of API");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  activeSpreadsheet.insertSheet().setName('Values of API');
  mySheet = activeSpreadsheet.getSheetByName("Values of API");
  var cell = mySheet.getRange('A1');
  cell.setValue('API key');
  mySheet = activeSpreadsheet.getSheetByName("UserID");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  activeSpreadsheet.insertSheet().setName('UserID');
  mySheet = activeSpreadsheet.getSheetByName("UserID");
  var range = mySheet.getRange('A1:B1').merge();
  range = mySheet.getRange('C1:D1').merge();
  mySheet.getRange('A1').setValue('User');
  mySheet.getRange('C1').setValue('Business');
  mySheet.getRange('E1').setValue('Bank Account');
  var values = ["id", "name"];
  mySheet.getRange('A2').setValue('ID');
  mySheet.getRange('B2').setValue('Email');
  mySheet.getRange('C2').setValue('ID');
  mySheet.getRange('D2').setValue('Name');
  mySheet.getRange('A1:E2').setHorizontalAlignment('center');
  mySheet = activeSpreadsheet.getSheetByName("Sheet1");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
}

// hàm lấy access token
function postApiKeyToToken() {
  var myFile = SpreadsheetApp.getActiveSpreadsheet();
  var apiSheet = myFile.getSheetByName('Values of API');
  var api_key = apiSheet.getRange('B1').getValue();
  var data = {
      code: api_key,
  };
  //Logger.log(api_key);
  var options = {
      method: "post",
      contentType: "application/json",
      // Convert the JavaScript object to a JSON string.
      payload: JSON.stringify(data),
  };
  var response = UrlFetchApp.fetch(
      "http://dev.casso.vn:3338/v1/token",
      options
  );

    // là các giá trị mà  The HTTP response về
  //Logger.log(response);

  // convert về json object để sử dụng
  var res = JSON.parse(response.getContentText());
  //Logger.log(res.access_token);

  if(response != null){
    apiSheet.getRange('A2').setValue('Refresh Token');
    apiSheet.getRange('B2').setValue(res.refresh_token);
    apiSheet.getRange('A3').setValue('Access Token');
    apiSheet.getRange('B3').setValue(res.access_token);
  }
  else{
    alert('Cannot get reponse from API');
  }

  return res.access_token;
}

// Hàm get api lấy user info
function getUserInfo(token) {
  var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UserID');
  var options = {
      method: "get",
      contentType: "application/json",
      headers: {
          Authorization: token,
      },
  };
  var response = UrlFetchApp.fetch(
      "http://dev.casso.vn:3338/v1/userInfo",
      options
  );
  var res = JSON.parse(response.getContentText());
  //Logger.log(res.data);
  var users = res.data;
  Logger.log(users.bankAccs);
  addUser(users, userSheet);
}

function addUser(value, userSheet){
  Logger.log(value.user);
  var userID = value.user.id;
  var userEmail = value.user.email;
  var bussID = value.business.id;
  var bussName = value.business.name;
  var bankAcc = value.bankAccs;
  if (bankAcc.length == 0) bankAcc = '';
  var row = [userID, userEmail, bussID, bussName, bankAcc];
  userSheet.appendRow(row);
}

function run() {
    var token = postApiKeyToToken();
    getUserInfo(token);
}
