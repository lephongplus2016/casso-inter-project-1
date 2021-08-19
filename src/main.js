//hàm tạo Menu trên thanh công cụ ============================================================================================
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Menu")
        .addItem("Start App", "formSheet")
        .addSeparator()
        .addItem("Input API key", "showWelcome")
        .addSeparator()
        .addItem("DashBoard", "showIndex")
        // .addSeparator()
        // .addItem("Get Token", "getTokenAgain")
        // .addSeparator()
        // .addItem("Get User Info", "runUserInfo")
        // .addSeparator()
        // .addItem("Get Transactions", "runTransactions")
        // .addSeparator()
        // .addItem("Draw Income Chart", "chart_html")
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
    var language = getLanguage();
    var apiSheet = myFile.getSheetByName("Values of API");
    if (language == ""){
        SpreadsheetApp.getUi().alert("You have to run start app first!");
        return false;
    }
    else if(language == "VN"){
      apiSheet = myFile.getSheetByName("Các giá trị API");
    }
    var api_key = apiSheet.getRange("B1").getValue();
    if (language == "EN") {
      if(api_key == ""){
        SpreadsheetApp.getUi().alert(
            "You deleted the api key, please login again to use the service!"
        );
        return false;
      }
      else if(api_key == "Please choose Input API key in the Menu before active other functions"){
        SpreadsheetApp.getUi().alert(
            "Please choose Input API key in the Menu before active other functions!"
        );
        return false;
      }
    }
    else{
      if(api_key == ""){
        SpreadsheetApp.getUi().alert(
            "Bạn đã xoá API key, vui lòng đăng nhập để sử dụng dịch vụ!"
        );
        return false;
      }
      else if(api_key == "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác"){
        SpreadsheetApp.getUi().alert(
            "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác!"
        );
        return false;
      }
    }
        return true;
    
}

// check cho mo dashboard
function checkTokenIsAvailable() {
    if(!checkAPIKeyIsAvailable()) return false;
    try {
      var apiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Values of API");
      var language = getLanguage();
      if(language == "VN") {
        apiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Các giá trị API");
      }
      token = apiSheet.getRange("B3").getValue();
      if(token == ""){
        if(language == "EN") SpreadsheetApp.getUi().alert("You deleted access token, please get token to use the device");
        else SpreadsheetApp.getUi().alert("Bạn đã xoá Access Token, vui lòng lấy Token để sử dụng dịch vụ");
        showGetToken();
        return false;
      }
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
      if(getLanguage() == "VN"){
        SpreadsheetApp.getUi().alert("Access Token đã hết hạn");
      }
      else SpreadsheetApp.getUi().alert("Access Token Is Expired ");
        showGetToken();
        return false;
    }
}
// cac ham trong menu ========================================================================================================
function run_input_api() {
    // check is user run start app
    var language = getLanguage();
    if (language == "") {
        SpreadsheetApp.getUi().alert("You have to run start app first!");
        return false;
    }
    else if(language == "EN"){
      var html = HtmlService.createHtmlOutputFromFile("input_api_key");
      SpreadsheetApp.getUi()
          .showModalDialog(html, "Log In to Casso")
          .setHeight(800);
    }
    else{
      var html = HtmlService.createHtmlOutputFromFile("input_api_key_VN");
      SpreadsheetApp.getUi()
          .showModalDialog(html, "Đăng nhập vào Casso")
          .setHeight(800);
    }
}

function runUserInfo() {
    // check api is available
    if (checkAPIKeyIsAvailable()) {
      var myFile = SpreadsheetApp.getActiveSpreadsheet();
      var apiSheet = myFile.getSheetByName("Values of API");
      if(apiSheet != null){
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
      else{
        apiSheet = myFile.getSheetByName("Các giá trị API");
        var api_key = apiSheet.getRange("B1").getValue();
        if (
            api_key ==
                "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác" &&
            api_key == null
        ) {
            SpreadsheetApp.getUi().alert(
                "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác!"
            );
        } else {
            showLoadingDialog();
            var token = apiSheet.getRange("B3").getValue();
            getUserInfo(token);
        }
      }
    }
}

function runTransactions() {
    // check api is available
    if (checkAPIKeyIsAvailable()) {
        var myFile = SpreadsheetApp.getActiveSpreadsheet();
      var apiSheet = myFile.getSheetByName("Values of API");
      if(apiSheet != null){
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
            var ui = SpreadsheetApp.getUi();
            var res = ui.prompt(`You want to get Transactions from which date?
  Example: 01-12-2021 (dd-mm-yyyy)`);
            var token = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName("Values of API")
                .getRange("B3")
                .getValue();
        }
      }
      else{
        apiSheet = myFile.getSheetByName("Các giá trị API");
        var api_key = apiSheet.getRange("B1").getValue();
        if (
            api_key ==
                "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác" &&
            api_key == null
        ) {
            SpreadsheetApp.getUi().alert(
                "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác!"
            );
        } else {
            var ui = SpreadsheetApp.getUi();
            var res = ui.prompt(`Bạn muốn lấy danh sách giao dịch từ ngày nào?
  ví dụ: 01-12-2021 (ngày-tháng-năm)`);
            var token = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName("Các giá trị API")
                .getRange("B3")
                .getValue();
        }
      } 
            var fromDate = res.getResponseText();

            fromDate = convertDate(fromDate);
            showLoadingSlowDialog();
            getTransaction(fromDate, token);
    }
}

// ham ui lua chon
function showIndex() {
    // check api is available
    if (checkAPIKeyIsAvailable()) {
        var myFile = SpreadsheetApp.getActiveSpreadsheet();
      var apiSheet = myFile.getSheetByName("Values of API");
      if(apiSheet != null){
        var api_key = apiSheet.getRange("B1").getValue();
        if (
            api_key ==
                "Please choose Input API key in the Menu before active other functions" &&
            api_key == null
        ) {
            SpreadsheetApp.getUi().alert(
                "Please choose Input API key in the Menu before active other functions!"
            );
        }
      }
      else {
        apiSheet = myFile.getSheetByName("Các giá trị API");
        var api_key = apiSheet.getRange("B1").getValue();
        if (
            api_key ==
                "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác" &&
            api_key == null
        ) {
            SpreadsheetApp.getUi().alert(
                "Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác!"
            );
        }
      }
      if(checkTokenIsAvailable()){
            var token = apiSheet.getRange("B3").getValue();
            var nameUser = getNameUser(token);
            var emailUser = getEmailUser(token);
        var language = getLanguage();
        // chay sidebar
        if(language == "EN"){
          var tmp = HtmlService.createTemplateFromFile("index");
          tmp.nameUser = nameUser;
          tmp.emailUser = emailUser;
          var html = tmp.evaluate().setTitle("Casso API");

          SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
              .showSidebar(html);
        }
        else{
          var tmp = HtmlService.createTemplateFromFile("index_VN");
          tmp.nameUser = nameUser;
          tmp.emailUser = emailUser;
          var html = tmp.evaluate().setTitle("Casso API");

          SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
              .showSidebar(html);
        }

      }
    }
}

// dialog welcome
function showWelcome() {
  var language = getLanguage();
  if(language == "EN"){
    var html = HtmlService.createHtmlOutputFromFile("welcome");
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModelessDialog(html, "Before start");
  }
  else{
    var html = HtmlService.createHtmlOutputFromFile("welcome_VN");
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModelessDialog(html, "Trước khi bắt đầu");
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

// get Language
function getLanguage(){
  if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Values of API"))
    return "EN";
  else if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Các giá trị API"))
    return "VN";
  return "";
}