function create_onOpen(){
    ScriptApp.newTrigger('updateTrans')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
  
    SpreadsheetApp
  }