function updateTrans() {
  var language = getLanguage();
  if(language == "EN"){
    var trans_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
    var fromDate = trans_sheet.getRange("B2").getDisplayValue();
    if(fromDate != ""){
      fromDate = convertDate(fromDate.toString());

      // User lựa chọn cập nhật giao dịch ngân hàng
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('Do you want to get new Transactions?', ui.ButtonSet.YES_NO);

      //xử lý lựa chọn YES
      if (response == ui.Button.YES) {
        
        var token = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Values of API").getRange("B3").getValue();
      }
    }
  }
  else{
    var trans_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Giao dịch ngân hàng");
    var fromDate = trans_sheet.getRange("B2").getDisplayValue();
    if(fromDate != ""){
      fromDate = convertDate(fromDate.toString());

      // User lựa chọn cập nhật giao dịch ngân hàng
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('Bạn có muốn cập nhật giao dịch mới không?', ui.ButtonSet.YES_NO);

      //xử lý lựa chọn YES
      if (response == ui.Button.YES) {
        
        var token = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Các giá trị API").getRange("B3").getValue();
      }
    }
  }
        getTransaction(fromDate, token);
}
