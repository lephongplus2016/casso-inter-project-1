// Hàm api lấy transactions
function getTransaction(token){
  var options = {
      method: "get",
      headers: {
          Authorization: token,
      },
  };
  var response = UrlFetchApp.fetch(
      "http://dev.casso.vn:3338/v1/transactions?fromDate=2021-07-01&sort=DESC&pageSize=100",
      options
  );
  var res = JSON.parse(response.getContentText());
  Logger.log(res.data.records);
  addTransactions(res.data.records);
}

// them transactions vao sheet 'Transactions'
function addTransactions(data){
  
  transactionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  data.forEach( function(i) {
    let id = i.id;
    let date = i.when;
    let des = i.description;
    let bankId = i.bankSubAccId;
    let amount = i.amount;
    var row = [id, date, des, bankId, amount];
    transactionSheet.appendRow(row);
  } )
}
