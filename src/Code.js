// hàm lấy access token
function postApiKeyToToken() {
    var data = {
        code: "baf72bc0-ea0c-11eb-a0ac-6d74b34a8413",
    };
    var options = {
        method: "post",
        contentType: "application/json",
        // Convert the JavaScript object to a JSON string.
        payload: JSON.stringify(data),
    };
    var response = UrlFetchApp.fetch(
        "http://dev.casso.vn:3338/v1/token",
        options
    );

    // là các giá trị mà  The HTTP response về
    Logger.log(response);

    // convert về json object để sử dụng
    var res = JSON.parse(response);
    Logger.log(res.access_token);
    return res.access_token;
}

// Hàm get api lấy user info
function getUserInfo(token) {
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

    Logger.log(response);
}

function run() {
    var token = postApiKeyToToken();
    getUserInfo(token);
}
