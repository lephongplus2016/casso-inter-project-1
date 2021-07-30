//hàm tạo Menu trên thanh công cụ ============================================================================================
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Menu")
        .addItem("Form Sheet", "formSheet")
        .addSeparator()
        .addItem("Get User Info", "runUserInfo")
        .addSeparator()
        .addItem("Get Transactions", "runTransactions")
        .addToUi();
}
//============================================================================================================================

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

function convertDate(date) {
    //xử lý ngày cho Transaction
    var year = +date.substring(6, 10);
    var month = +date.substring(3, 5);
    var day = +date.substring(0, 2);
    let date_trans = new Date(year, month - 1, day + 1);
    date_trans = Utilities.formatDate(date_trans, "GTM", "yyyy-MM-dd");
    return date_trans;
}

// cac ham trong menu ========================================================================================================
function runUserInfo() {
    var myFile = SpreadsheetApp.getActiveSpreadsheet();
    var apiSheet = myFile.getSheetByName("Values of API");
    var api_key = apiSheet.getRange("B1").getValue();
    if (api_key == "Fill API key here") {
        SpreadsheetApp.getUi().alert("Please fill API Key!");
    } else {
        showLoadingDialog();
        var token = postApiKeyToToken();
        getUserInfo(token);
    }
}

function runTransactions() {
    var myFile = SpreadsheetApp.getActiveSpreadsheet();
    var apiSheet = myFile.getSheetByName("Values of API");
    var api_key = apiSheet.getRange("B1").getValue();
    if (api_key == "Fill API key here") {
        SpreadsheetApp.getUi().alert("Please fill API Key!");
    } else {
        // nhap ngay bat dau giao dich
        var ui = SpreadsheetApp.getUi();
        var res = ui.prompt(`Bạn muốn lấy danh sách giao dịch từ ngày nào?
  ví dụ: 01-12-2021`);
        var fromDate = res.getResponseText();

        fromDate = convertDate(fromDate);
        showLoadingSlowDialog();
        var token = postApiKeyToToken();
        getTransaction(fromDate, token);
    }
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
