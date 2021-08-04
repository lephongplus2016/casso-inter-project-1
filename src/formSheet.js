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
  mySheet.getRange('B1').setValue('Please choose Input API key in the Menu before active other functions');
  mySheet.setColumnWidth(2,200);
  //-------------------------------------------------------------- Sheet UserID------------------------------------------------
  mySheet = activeSpreadsheet.getSheetByName("UserID");
  if (mySheet != null) {
    activeSpreadsheet.deleteSheet(mySheet);
  }
  activeSpreadsheet.insertSheet().setName('UserID');
  mySheet = activeSpreadsheet.getSheetByName("UserID");
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
  activeSpreadsheet.insertSheet().setName('Transactions');
  mySheet = activeSpreadsheet.getSheetByName("Transactions");

  mySheet.getRange('A1').setValue('ID');
  mySheet.getRange('B1').setValue('Date');
  mySheet.getRange('B:B').setNumberFormat("dd-MM-yyyy");
  mySheet.getRange('C1').setValue('Description');
  mySheet.getRange('D1').setValue('Bank Account Id');
  mySheet.getRange('E1').setValue('Amount');

  mySheet.setColumnWidth(3,400);
 mySheet.setColumnWidth(4,150);


  mySheet.getRange('A1:E1').setHorizontalAlignment('center');
  mySheet.getRange('A:B').setHorizontalAlignment('center');
  mySheet.getRange('A1:E1').setFontWeight('bold');
  
  SpreadsheetApp.getUi().alert('Please choose Input API key in the Menu before active other functions');

}