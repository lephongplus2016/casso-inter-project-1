// Hàm api lấy transactions
function getTransaction(fromDate, token) {
    try{
      var options = {
          method: "get",
          headers: {
              Authorization: token,
          },
      };
      var response = UrlFetchApp.fetch(
          "http://dev.casso.vn:3338/v1/transactions?fromDate=" +
              fromDate +
              "&sort=DESC&pageSize=1000",
          options
      );
          // SpreadsheetApp.getUi().alert(response);
  
    }
    catch(e){
      if(language == "EN") SpreadsheetApp.getUi().alert("Access Token Is Expired ");
      else SpreadsheetApp.getUi().alert("Access Token đã hết hạn");
      showGetToken();
    }
  
      var res = JSON.parse(response.getContentText());
      //Logger.log(res.data.records);
      // SpreadsheetApp.getUi().alert(res.data);
      addTransactions(res.data.records);
    
  }
  
  // them transactions vao sheet 'Transactions'
  function addTransactions(data) {
      var transactionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          "Transactions"
      );
      if(getLanguage() == "VN")
        transactionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          "Giao dịch ngân hàng"
        );
  
      // mang cac giao dich moi nhat
      var new_tran = [];
  
      data.forEach(function (i) {
          let id = i.id;
          let date = i.when;
          let des = i.description;
          let bankId = i.bankSubAccId;
          let amount = i.amount;
  
          // xử lý id
          var numRow = transactionSheet.getLastRow();
          if (numRow >= 2) {
              var id_list = transactionSheet
                  .getRange(2, 1, numRow - 1)
                  .getValues();
              var have = 0;
              for (var i = 0; i < numRow - 1; i++) {
                  if (id == id_list[i]) have = 1;
              }
          }
  
          //xử lý ngày
          var year = +date.substring(0, 4);
          var month = +date.substring(5, 7);
          var day = +date.substring(8, 10);
          let date_trans = new Date(year, month - 1, day + 1);
          date = Utilities.formatDate(date_trans, "GTM", "dd-MM-yyyy");
  
          //xử lý description
          if (des.length > 50) {
              var str1 = des.substr(0, des.length / 2);
              var str2 = des.substr(des.length / 2);
              str1 = str1 + str2.substr(0, str2.indexOf(" "));
              str2 = str2.substr(str2.indexOf(" "));
              des = str1 + "\n" + str2;
          }
  
          var row = [id, date, des, bankId, amount];
          if (!have) {
              transactionSheet.appendRow(row);
              var cell_colored = "E" + (numRow + 1).toString();
              if (amount >= 0)
                  transactionSheet.getRange(cell_colored).setFontColor("green");
              else transactionSheet.getRange(cell_colored).setFontColor("red");
  
              // xử lý amount format
              var formats = [["#,###"]];
              transactionSheet.getRange(cell_colored).setNumberFormat(formats);
  
              //log nhung ngay lon hon ngay dau tien
              var firstDate = transactionSheet.getRange("B2").getDisplayValue();
              if (compareDate(date, firstDate)) {
                  //Logger.log("Ngay nay moi ne :" + date + " tại row :" + numRow);
                  new_tran.push(numRow);
              }
          }
      });
  
      // sắp xếp theo ngày mới nhất
      var numRow = new_tran.length;
      if(numRow > 0){
        transactionSheet.insertRows(2, numRow);
        transactionSheet
            .getRange(new_tran[0] + numRow + 1, 1, numRow, 5)
            .copyTo(transactionSheet.getRange(2, 1, numRow, 5));
        transactionSheet.deleteRows(new_tran[0] + numRow + 1, numRow);
      }
  }
  
  function compareDate(date01, date02) {
      date01 = date01.toString();
     
      var thisdate1 = date01.substring(0, 2);
      var thismonth1 = date01.substring(3, 5);
      var thisyear1 = date01.substring(6, 10);
      let new1 = new Date(thisyear1, thismonth1 - 1, thisdate1);
  
      var thisdate2 = date02.substring(0, 2);
      var thismonth2 = date02.substring(3, 5);
      var thisyear2 = date02.substring(6, 10);
      let new2 = new Date(thisyear2, thismonth2 - 1, thisdate2);
  
      return new1.valueOf() >= new2.valueOf();
  }
  