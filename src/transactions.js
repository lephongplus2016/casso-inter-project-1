// Hàm api lấy transactions
function getTransaction(fromDate, token) {
    var options = {
        method: "get",
        headers: {
            Authorization: token,
        },
    };
    var response = UrlFetchApp.fetch(
        "http://dev.casso.vn:3338/v1/transactions?fromDate=" +
            fromDate +
            "&sort=DESC&pageSize=100",
        options
    );
    var res = JSON.parse(response.getContentText());
    Logger.log(res.data.records);
    addTransactions(res.data.records);
}

// them transactions vao sheet 'Transactions'
function addTransactions(data){
  
  transactionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Giao dịch ngân hàng');
  var add_new = 0;
  data.forEach( function(i) {
    let id = i.id;
    let date = i.when;
    let des = i.description;
    let bankId = i.bankSubAccId;
    let amount = i.amount;

    // xử lý id
    var numRow = transactionSheet.getLastRow();
    if(numRow >= 2){
      var id_list = transactionSheet.getRange(2, 1, numRow - 1).getValues();
      var have = 0;
      for(var i = 0; i < numRow - 1; i++){
        if(id == id_list[i]) have = 1;
      }
    }

    //xử lý ngày
    var year = +date.substring(0, 4);
    var month = +date.substring(5, 7);
    var day = +date.substring(8, 10);
    let date_trans = new Date(year, month - 1, day);
    date = Utilities.formatDate(date_trans, "GTM", "dd-MM-yyyy");

    //xử lý description
    if(des.length > 50){
      var str1 = des.substr(0, des.length/2);
      var str2 = des.substr(des.length/2);
      str1 = str1 + str2.substr(0, str2.indexOf(' '));
      str2 = str2.substr(str2.indexOf(' '));
      des = str1 + "\n" + str2;
    }
    
    var row = [id, date, des, bankId, amount];
    if(!have){
      transactionSheet.appendRow(row);
      var cell_colored = 'E' + (numRow+1).toString();
      if(amount >= 0) transactionSheet.getRange(cell_colored).setFontColor("green");
      else transactionSheet.getRange(cell_colored).setFontColor("red");
    }
  } )
  transactionSheet.getRange('A1:E').sort({column: 2, ascending: false});
}
