//hàm tạo Menu trên thanh công cụ
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Menu")
        .addItem("Intern 1.0.0", "run")
        .addSeparator()
        .addItem("Form Sheet", "formSheet")
        .addToUi();
}

//hàm tạo mẫu gg sheet
function formSheet() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var mySheet = activeSpreadsheet.getSheetByName("Values of API");
    if (mySheet != null) {
        activeSpreadsheet.deleteSheet(mySheet);
    }
    activeSpreadsheet.insertSheet().setName("Values of API");
    mySheet = activeSpreadsheet.getSheetByName("Values of API");
    var cell = mySheet.getRange("A1");
    cell.setValue("API key");
    mySheet.getRange("B1").setValue("Fill API key here");
    mySheet = activeSpreadsheet.getSheetByName("UserID");
    if (mySheet != null) {
        activeSpreadsheet.deleteSheet(mySheet);
    }
    activeSpreadsheet.insertSheet().setName("UserID");
    mySheet = activeSpreadsheet.getSheetByName("UserID");
    var range = mySheet.getRange("A1:B1").merge();
    range = mySheet.getRange("C1:D1").merge();
    range = mySheet.getRange("E1:F1").merge();

    mySheet.getRange("A1").setValue("User");
    mySheet.getRange("C1").setValue("Business");
    mySheet.getRange("E1").setValue("Bank Account");
    var values = ["id", "name"];
    mySheet.getRange("A2").setValue("ID");
    mySheet.getRange("B2").setValue("Email");
    mySheet.getRange("C2").setValue("ID");
    mySheet.getRange("D2").setValue("Name");
    mySheet.getRange("E2").setValue("Bank Name");
    mySheet.getRange("F2").setValue("Bank Account Id");

    mySheet.setColumnWidth(2, 200);
    mySheet.setColumnWidth(5, 200);
    mySheet.setColumnWidth(6, 150);

    mySheet.getRange("A1:F2").setHorizontalAlignment("center");
    mySheet.getRange("A1:F2").setFontWeight("bold");

    mySheet = activeSpreadsheet.getSheetByName("Sheet1");
    if (mySheet != null) {
        activeSpreadsheet.deleteSheet(mySheet);
    }
    SpreadsheetApp.getUi().alert("Please fill API key in Values of API Sheet");
}

// hàm lấy access token
function postApiKeyToToken() {
    var myFile = SpreadsheetApp.getActiveSpreadsheet();
    var apiSheet = myFile.getSheetByName("Values of API");
    var api_key = apiSheet.getRange("B1").getValue();
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

    if (response != null) {
        apiSheet.getRange("A2").setValue("Refresh Token");
        apiSheet.getRange("B2").setValue(res.refresh_token);
        apiSheet.getRange("A3").setValue("Access Token");
        apiSheet.getRange("B3").setValue(res.access_token);
    } else {
        SpreadsheetApp.getUi().alert("Cannot get reponse from API");
    }

    return res.access_token;
}

// Hàm get api lấy user info
function getUserInfo(token) {
    var userSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserID");
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

function addUser(value, userSheet) {
    Logger.log(value.user);
    var userID = value.user.id;
    var userEmail = value.user.email;
    var bussID = value.business.id;
    var bussName = value.business.name;
    var bankAcc = value.bankAccs[0].bank.codeName;
    var bankId = value.bankAccs[0].bankSubAccId;

    var row = [userID, userEmail, bussID, bussName, bankAcc, bankId];
    userSheet.appendRow(row);

    // list bank account
    for (var i = 1; i < value.bankAccs.length; i++) {
        let tempAcc = value.bankAccs[i].bank.codeName;
        let tempId = value.bankAccs[i].bankSubAccId;
        let row = ["", "", "", "", tempAcc, tempId];
        userSheet.appendRow(row);
    }
}

function run() {
    var myFile = SpreadsheetApp.getActiveSpreadsheet();
    var apiSheet = myFile.getSheetByName("Values of API");
    var api_key = apiSheet.getRange("B1").getValue();
    if (api_key == "Fill API key here") {
        SpreadsheetApp.getUi().alert("Please fill API Key!");
    } else {
        var token = postApiKeyToToken();
        getUserInfo(token);
    }
}
