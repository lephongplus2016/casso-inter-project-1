//hàm tạo Menu trên thanh công cụ ============================================================================================
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Menu")
        .addItem("Start App", "formSheet")
        .addSeparator()
        .addItem("Input API key", "showWelcome")
        .addSeparator()
        .addItem("DashBoard", "showIndex")
        .addSeparator()
        .addItem("Get Token", "getTokenAgain")
        .addSeparator()
        .addItem("Get User Info", "runUserInfo")
        .addSeparator()
        .addItem("Get Transactions", "runTransactions")
        .addSeparator()
        .addItem("Draw Income Chart", "draw_chart")
        .addToUi();
}
//============================================================================================================================

function convertDate(date) {
    //xử lý ngày cho Transaction
    var year = +date.substring(6, 10);
    var month = +date.substring(3, 5);
    var day = +date.substring(0, 2);
    let date_trans = new Date(year, month - 1, day + 1);
    date_trans = Utilities.formatDate(date_trans, "GTM", "yyyy-MM-dd");
    return date_trans;
}

function removeAPIKey() {
    var myFile = SpreadsheetApp.getActiveSpreadsheet();
    var apiSheet = myFile.getSheetByName("Values of API");
    var rangRemove = apiSheet.getRange("B1:B4");
    rangRemove.clearContent();
}

function checkAPIKeyIsAvailable() {
    var myFile = SpreadsheetApp.getActiveSpreadsheet();
    var apiSheet = myFile.getSheetByName("Values of API");
    if (apiSheet == null) {
        SpreadsheetApp.getUi().alert("You have to run start app first!");
        return false;
    }
    var api_key = apiSheet.getRange("B1").getValue();
    if (api_key == "") {
        SpreadsheetApp.getUi().alert(
            "You deleted the api key, please login again to use the service!"
        );
        return false;
    } else {
        return true;
    }
}

// check cho mo dashboard
function checkTokenIsAvailable() {
    var myFile = SpreadsheetApp.getActiveSpreadsheet();

    var apiSheet = myFile.getSheetByName("Values of API");
    if (apiSheet == null) {
        SpreadsheetApp.getUi().alert("You have to run start app first!");
        return false;
    }
    var token = apiSheet.getRange("B3").getValue();
    if (token == "") {
        SpreadsheetApp.getUi().alert(
            "You deleted the api key, please login again to use the service!"
        );
        return false;
    }
    try {
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
        return true;
    } catch (e) {
        SpreadsheetApp.getUi().alert("Access Token Is Expired ");
        showGetToken();
        return false;
    }
}
// cac ham trong menu ========================================================================================================
function run_input_api() {
    // check is user run start app
    var myFile = SpreadsheetApp.getActiveSpreadsheet();
    var apiSheet = myFile.getSheetByName("Values of API");
    if (apiSheet == null) {
        SpreadsheetApp.getUi().alert("You have to run start app first!");
        return false;
    }
    var html = HtmlService.createHtmlOutputFromFile("input_api_key");
    SpreadsheetApp.getUi()
        .showModalDialog(html, "Log In to Casso")
        .setHeight(800);
}

function runUserInfo() {
    // check api is available
    if (checkAPIKeyIsAvailable()) {
        var myFile = SpreadsheetApp.getActiveSpreadsheet();
        var apiSheet = myFile.getSheetByName("Values of API");
        var api_key = apiSheet.getRange("B1").getValue();
        if (
            api_key ==
                "Please choose Input API key in the Menu before active other functions" &&
            api_key == null
        ) {
            SpreadsheetApp.getUi().alert(
                "Please choose Input API key in the Menu before active other functions!"
            );
        } else {
            showLoadingDialog();
            var token = apiSheet.getRange("B3").getValue();
            getUserInfo(token);
        }
    }
}

function runTransactions() {
    // check api is available
    if (checkAPIKeyIsAvailable()) {
        var myFile = SpreadsheetApp.getActiveSpreadsheet();
        var apiSheet = myFile.getSheetByName("Values of API");
        var api_key = apiSheet.getRange("B1").getValue();
        if (
            api_key ==
                "Please choose Input API key in the Menu before active other functions" &&
            api_key == null
        ) {
            SpreadsheetApp.getUi().alert(
                "Please choose Input API key in the Menu before active other functions!"
            );
        } else {
            // nhap ngay bat dau giao dich
            var ui = SpreadsheetApp.getUi();
            var res = ui.prompt(`Bạn muốn lấy danh sách giao dịch từ ngày nào?
  ví dụ: 01-12-2021`);
            var fromDate = res.getResponseText();

            fromDate = convertDate(fromDate);
            showLoadingSlowDialog();
            var token = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName("Values of API")
                .getRange("B3")
                .getValue();
            getTransaction(fromDate, token);
        }
    }
}

// ham ui lua chon
function showIndex() {
    // check api is available
    if (checkAPIKeyIsAvailable() && checkTokenIsAvailable()) {
        //lay ten nguoi dung
        var myFile = SpreadsheetApp.getActiveSpreadsheet();
        var apiSheet = myFile.getSheetByName("Values of API");
        var api_key = apiSheet.getRange("B1").getValue();
        if (
            api_key ==
                "Please choose Input API key in the Menu before active other functions" &&
            api_key == null
        ) {
            SpreadsheetApp.getUi().alert(
                "Please choose Input API key in the Menu before active other functions!"
            );
        } else {
            var token = apiSheet.getRange("B3").getValue();
            var nameUser = getNameUser(token);
            var emailUser = getEmailUser(token);
        }

        // chay sidebar
        var tmp = HtmlService.createTemplateFromFile("index");
        tmp.nameUser = nameUser;
        tmp.emailUser = emailUser;
        var html = tmp.evaluate().setTitle("Casso API");

        SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
            .showSidebar(html);
    }
}

// dialog welcome
function showWelcome() {
    var html = HtmlService.createHtmlOutputFromFile("welcome");

    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModelessDialog(html, "Welcome");
}
//==========================================================================================================================

// hieu ung ================================================================================================================
function showLoadingDialog() {
    var html = HtmlService.createHtmlOutputFromFile("loading")
        .setWidth(200)
        .setHeight(100);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(html, "App is loading!");
}

function showLoadingSlowDialog() {
    var html = HtmlService.createHtmlOutputFromFile("loadingSlow")
        .setWidth(200)
        .setHeight(100);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(html, "App is loading!");
}
