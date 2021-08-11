function updateTrans() {
  var trans_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  var fromDate = trans_sheet.getRange("B2").getDisplayValue();
  if(fromDate != null){
    fromDate = convertDate(fromDate.toString());

    // User lựa chọn cập nhật giao dịch ngân hàng
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Do you want to get new Transactions?', ui.ButtonSet.YES_NO);

    //xử lý lựa chọn YES
    if (response == ui.Button.YES) {
      
      var token = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Values of API").getRange("B3").getValue();
      getTransaction(fromDate, token);
    }
  }
}
