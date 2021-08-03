function getToken(email, key) {
    try{
      var myFile = SpreadsheetApp.getActiveSpreadsheet();
      var apiSheet = myFile.getSheetByName("Values of API");
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
      if(res2.data.user.email != email){
        apiSheet.getRange("B1").setValue("")
        apiSheet.getRange("B2").setValue("");
        apiSheet.getRange("B3").setValue("");
        apiSheet.getRange("B4").setValue("");
        SpreadsheetApp.getUi().alert("This API key is not belong to this email. \n Please try again");
      }
      else{
        var time_expire = res1.expires_in;
        var time = new Date();
        // đổi time sang đơn vị ms
        time = time.getTime() + time_expire * 1000; 
        var newTime = new Date(time);
        newTime = Utilities.formatDate(newTime, Session.getScriptTimeZone(), "dd-MM-yyyy hh:mm:ss a");
        apiSheet.getRange("B1").setValue(key);
        apiSheet.getRange("A2").setValue("Refresh Token");
        apiSheet.getRange("B2").setValue(res1.refresh_token);
        apiSheet.getRange("A3").setValue("Access Token");
        apiSheet.getRange("B3").setValue(res1.access_token);
        apiSheet.getRange("A4").setValue("Expire Time");
        apiSheet.getRange("B4").setValue(newTime);
        SpreadsheetApp.getUi().alert("Get Token successfully");
      }
    }
    catch(e){
      apiSheet.getRange("B1").setValue("")
      apiSheet.getRange("B2").setValue("");
      apiSheet.getRange("B3").setValue("");
      apiSheet.getRange("B4").setValue("");
      SpreadsheetApp.getUi().alert("Wrong API key. \n Please try again");
    }
  }