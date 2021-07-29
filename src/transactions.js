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

    //xử lý ngày
    var year = +date.substring(0, 4);
    var month = +date.substring(5, 7);
    var day = +date.substring(8, 10);
    let date_trans = new Date(year, month - 1, day);
    date_trans = Utilities.formatDate(date_trans, "GTM", "dd-MM-yyyy");

    //xử lý description
    let des = i.description;
    if(des.length > 50){
      var str1 = des.substr(0, des.length/2);
      var str2 = des.substr(des.length/2);
      str1 = str1 + str2.substr(0, str2.indexOf(' '));
      str2 = str2.substr(str2.indexOf(' '));
      des = str1 + "\n" + str2;
    }
    
    let bankId = i.bankSubAccId;
    let amount = i.amount;
    var row = [id, date_trans, des, bankId, amount];
    transactionSheet.appendRow(row);
  } )
}
