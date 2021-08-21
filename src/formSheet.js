//hàm tạo mẫu gg sheet
function formSheet(){
  var ui = SpreadsheetApp.getUi();
  var language = "EN"
  var response = ui.alert('Bạn có muốn thiết lập Tiếng Việt làm ngôn ngữ mặc định? \n If you want to set English as your default language, please click on No or turn off this form', ui.ButtonSet.YES_NO);
  if(response == ui.Button.YES){
    language = "VN";
  }
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var mySheet = activeSpreadsheet.getSheetByName("Values of API");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  mySheet = activeSpreadsheet.getSheetByName("Các giá trị API");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  if(language == "EN"){
    activeSpreadsheet.insertSheet().setName('Values of API');
    mySheet = activeSpreadsheet.getSheetByName("Values of API");
    var cell = mySheet.getRange('A1');
    cell.setValue('API key');
    mySheet.getRange('B1').setValue('Please choose Input API key in the Menu before active other functions');
  }
  else{
    activeSpreadsheet.insertSheet().setName('Các giá trị API');
    mySheet = activeSpreadsheet.getSheetByName("Các giá trị API");
    var cell = mySheet.getRange('A1');
    cell.setValue('API key');
    mySheet.getRange('B1').setValue('Vui lòng chọn Input API key trong Menu trước khi chạy các chức năng khác');
  }
  mySheet.setColumnWidth(2,200);
  //-------------------------------------------------------------- Sheet UserID------------------------------------------------
  mySheet = activeSpreadsheet.getSheetByName("User Info");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  mySheet = activeSpreadsheet.getSheetByName("Thông tin người dùng");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  if(language == "EN"){
    activeSpreadsheet.insertSheet().setName('User Info');
    mySheet = activeSpreadsheet.getSheetByName("User Info");
    var range = mySheet.getRange('A1:B1').merge();
    range = mySheet.getRange('C1:D1').merge();
      range = mySheet.getRange('E1:F1').merge();

    mySheet.getRange('A1').setValue('User');
    mySheet.getRange('C1').setValue('Business');
    mySheet.getRange('E1').setValue('Bank Account');
    var values = ["id", "name"];
    mySheet.getRange('A2').setValue('ID');
    mySheet.getRange('B2').setValue('Email');
    mySheet.getRange('C2').setValue('ID');
    mySheet.getRange('D2').setValue('Name');
    mySheet.getRange('E2').setValue('Bank Name');
    mySheet.getRange('F2').setValue('Bank Account Id');
  }
  else{
    activeSpreadsheet.insertSheet().setName('Thông tin người dùng');
    mySheet = activeSpreadsheet.getSheetByName("Thông tin người dùng");
    var range = mySheet.getRange('A1:B1').merge();
    range = mySheet.getRange('C1:D1').merge();
      range = mySheet.getRange('E1:F1').merge();

    mySheet.getRange('A1').setValue('Người dùng');
    mySheet.getRange('C1').setValue('Doanh nghiệp');
    mySheet.getRange('E1').setValue('Tài khoản ngân hàng');
    var values = ["id", "name"];
    mySheet.getRange('A2').setValue('ID');
    mySheet.getRange('B2').setValue('Email');
    mySheet.getRange('C2').setValue('ID');
    mySheet.getRange('D2').setValue('Tên');
    mySheet.getRange('E2').setValue('Tên ngân hàng');
    mySheet.getRange('F2').setValue('ID tài khoản ngân hàng');
  }
  mySheet.setColumnWidth(2,200) ;
  mySheet.setColumnWidth(5,200) ;
  mySheet.setColumnWidth(6,150) ;


  mySheet.getRange('A1:F2').setHorizontalAlignment('center');
  mySheet.getRange('A1:F2').setFontWeight('bold');
  
  mySheet = activeSpreadsheet.getSheetByName("Sheet1");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  //--------------------------------------------------------------------------Sheet Transactions----------------------
  mySheet = activeSpreadsheet.getSheetByName("Transactions");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  mySheet = activeSpreadsheet.getSheetByName("Giao dịch ngân hàng");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  if(language == "EN"){
    activeSpreadsheet.insertSheet().setName('Transactions');
    mySheet = activeSpreadsheet.getSheetByName("Transactions");

    mySheet.getRange('A1').setValue('ID');
    mySheet.getRange('B1').setValue('Date');
    mySheet.getRange('B:B').setNumberFormat("dd-MM-yyyy");
    mySheet.getRange('C1').setValue('Description');
    mySheet.getRange('D1').setValue('Bank Account Id');
    mySheet.getRange('E1').setValue('Amount');
  }
  else{
    activeSpreadsheet.insertSheet().setName('Giao dịch ngân hàng');
    mySheet = activeSpreadsheet.getSheetByName("Giao dịch ngân hàng");

    mySheet.getRange('A1').setValue('ID');
    mySheet.getRange('B1').setValue('Ngày diễn ra');
    mySheet.getRange('B:B').setNumberFormat("dd-MM-yyyy");
    mySheet.getRange('C1').setValue('Mô tả');
    mySheet.getRange('D1').setValue('ID tài khoản ngân hàng');
    mySheet.getRange('E1').setValue('Giá trị');
  }

  mySheet.setColumnWidth(3,400);
  mySheet.setColumnWidth(4,150);


  mySheet.getRange('A1:E1').setHorizontalAlignment('center');
  mySheet.getRange('A:B').setHorizontalAlignment('center');
  mySheet.getRange('A1:E1').setFontWeight('bold');
  
  var chartSheet = activeSpreadsheet.getSheetByName("DailyChart");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    chartSheet = activeSpreadsheet.getSheetByName("Biểu đồ ngày");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
  chartSheet = activeSpreadsheet.getSheetByName("MonthlyChart");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    chartSheet = activeSpreadsheet.getSheetByName("Biểu đồ tháng");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
  chartSheet = activeSpreadsheet.getSheetByName("QuarterlyChart");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    chartSheet = activeSpreadsheet.getSheetByName("Biểu đồ quý");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
  
  if(language == "EN")
    SpreadsheetApp.getUi().alert('Please choose Input API key in the Menu before active other functions');
  else SpreadsheetApp.getUi().alert('Vui lòng chọn Input API key trong Menu trước khi kích hoạt các chức năng khác');
}