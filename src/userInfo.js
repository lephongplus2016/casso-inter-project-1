// Hàm get api lấy user info
function getUserInfo(token) {
  try{
    var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UserID');
    var options = {
        method: "get",
        contentType: "application/json",
        headers: {
            Authorization: token,
        },
    };
    var response = UrlFetchApp.fetch(
        "http://dev.casso.vn:3338/v1/userInfo",
        options
    );
    var res = JSON.parse(response.getContentText());
    var users = res.data;
    //Logger.log(users.bankAccs);
    addUser(users, userSheet);
  }
  catch(e){
    SpreadsheetApp.getUi().alert("Access Token Is Expired \n Please choose Get Token in the Menu and then replay the function you want");
  }
}

// them user trong sheet userinfo
function addUser(value, userSheet){
  //Logger.log(value.user);
  var userID = value.user.id;
  var userEmail = value.user.email;
  var bussID = value.business.id;
  var bussName = value.business.name;
  var bankAcc = value.bankAccs[0].bank.codeName;
  var bankId = value.bankAccs[0].bankSubAccId;    

  var row = [userID, userEmail, bussID, bussName, bankAcc, bankId];
  userSheet.appendRow(row);

// list bank account
  for( var i=1; i<value.bankAccs.length; i++){
    let tempAcc = value.bankAccs[i].bank.codeName;
    let tempId = value.bankAccs[i].bankSubAccId;
    let row = ["","","","",tempAcc, tempId];
    userSheet.appendRow(row);
  }

}

// Lay thong tin ten nguoi dung
function getNameUser(token){
    var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UserID');
  var options = {
      method: "get",
      contentType: "application/json",
      headers: {
          Authorization: token,
      },
  };
  var response = UrlFetchApp.fetch(
      "http://dev.casso.vn:3338/v1/userInfo",
      options
  );
  var res = JSON.parse(response.getContentText());
  var users = res.data.business.name;
  return users;
}