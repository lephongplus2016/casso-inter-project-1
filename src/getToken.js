function getToken(email, key) {
    try {
        showLoadingDialog();
        var myFile = SpreadsheetApp.getActiveSpreadsheet();
        var apiSheet = myFile.getSheetByName("Values of API");
        if(getLanguage() == "VN") apiSheet = myFile.getSheetByName("Các giá trị API");
        var data = {
            code: key,
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
        var res1 = JSON.parse(response.getContentText());

        options = {
            method: "get",
            contentType: "application/json",
            headers: {
                Authorization: res1.access_token,
            },
        };
        response = UrlFetchApp.fetch(
            "http://dev.casso.vn:3338/v1/userInfo",
            options
        );
        var res2 = JSON.parse(response.getContentText());
        if (res2.data.user.email != email) {
            apiSheet.getRange("B1").setValue("");
            apiSheet.getRange("B2").setValue("");
            apiSheet.getRange("B3").setValue("");
            apiSheet.getRange("B4").setValue("");
            if(getLanguage() == "EN")
              SpreadsheetApp.getUi().alert(
                  "This API key is not belong to this email. \n Please try again"
              );
            else
              SpreadsheetApp.getUi().alert(
                  "API key này không thuộc về email đã nhập. \n Vui lòng thử lại"
              );
            input_api_html();
        } else {
            var time_expire = res1.expires_in;
            var time = new Date();
            // đổi time sang đơn vị ms
            time = time.getTime() + time_expire * 1000;
            var newTime = new Date(time);
            newTime = Utilities.formatDate(
                newTime,
                Session.getScriptTimeZone(),
                "dd-MM-yyyy hh:mm:ss a"
            );
            apiSheet.getRange("B1").setValue(key);
            apiSheet.getRange("A2").setValue("Refresh Token");
            apiSheet.getRange("B2").setValue(res1.refresh_token);
            apiSheet.getRange("A3").setValue("Access Token");
            apiSheet.getRange("B3").setValue(res1.access_token);
            if(getLanguage() == "EN"){
              apiSheet.getRange("A4").setValue("Expire Time");
              apiSheet.getRange("B4").setValue(newTime);
              SpreadsheetApp.getUi().alert("Get Token successfully");
            }
            else{
              apiSheet.getRange("A4").setValue("Thời gian hết hạn");
              apiSheet.getRange("B4").setValue(newTime);
              SpreadsheetApp.getUi().alert("Lấy Token thành công");
            }
            showIndex();
        }
    } catch (e) {
        apiSheet.getRange("B1").setValue("");
        apiSheet.getRange("B2").setValue("");
        apiSheet.getRange("B3").setValue("");
        apiSheet.getRange("B4").setValue("");
        if(getLanguage() == "EN")
          SpreadsheetApp.getUi().alert("Wrong API key. \n Please try again");
        else
          SpreadsheetApp.getUi().alert("Sai API key. \n Vui lòng thử lại");
        input_api_html();
    }
}

function getTokenAgain() {
    try {
        showLoadingDialog();
        var language = "EN";
        var apiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Values of API");
        if(apiSheet == null){
          apiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Các giá trị API");
          language = "VN";
        }
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
        var res1 = JSON.parse(response.getContentText());
        var time_expire = res1.expires_in;
        var time = new Date();
        // đổi time sang đơn vị ms
        time = time.getTime() + time_expire * 1000;
        var newTime = new Date(time);
        newTime = Utilities.formatDate(
            newTime,
            Session.getScriptTimeZone(),
            "dd-MM-yyyy hh:mm:ss"
        );
        apiSheet.getRange("A2").setValue("Refresh Token");
        apiSheet.getRange("B2").setValue(res1.refresh_token);
        apiSheet.getRange("A3").setValue("Access Token");
        apiSheet.getRange("B3").setValue(res1.access_token);
        if(language == "EN"){
          apiSheet.getRange("A4").setValue("Expire Time");
          apiSheet.getRange("B4").setValue(newTime);
          SpreadsheetApp.getUi().alert("Get Token successfully");
        }
        else{
          apiSheet.getRange("A4").setValue("Thời gian hết hạn");
          apiSheet.getRange("B4").setValue(newTime);
          SpreadsheetApp.getUi().alert("Lấy Token thành công");
        }
        showIndex();
    } catch (e) {
        apiSheet.getRange("B1").setValue("");
        apiSheet.getRange("B2").setValue("");
        apiSheet.getRange("B3").setValue("");
        apiSheet.getRange("B4").setValue("");
        SpreadsheetApp.getUi().alert(
            "You deleted the api key, please login again to use the service!"
        );
        input_api_html();
    }
}

// hàm getTokenAgain sẽ không thể được gọi ở menu nữa, chỉ xuất hiện khi gọi showGetToken
function showGetToken() {
    var ui = SpreadsheetApp.getUi();
    var apiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Values of API");
    if(apiSheet != null){
      var result = ui.alert(
          "Please confirm",
          "Your token has expired, do you want to get a new token?",
          ui.ButtonSet.YES_NO
      );
    }
    else{
      var result = ui.alert(
          "Vui lòng xác nhận",
          "Token của bạn đã hết bạn, bạn có muốn lấy token mới không?",
          ui.ButtonSet.YES_NO
      );
    }

    if (result == ui.Button.YES) {
        // User clicked "Yes".
        getTokenAgain();
    }
}
