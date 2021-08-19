function chart_html(){
  var language = getLanguage();
  if(language == "EN"){
    var html = HtmlService.createHtmlOutputFromFile("chart_req");
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModelessDialog(html, "Chart requirements");
  }
  else{
    var html = HtmlService.createHtmlOutputFromFile("chart_req_VN");
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModelessDialog(html, "Các yêu cầu vẽ biểu đồ");
  }
}

function convertDateFromHTML(date){
  var day = date.substring(8,10);
  var month = date.substring(5,7);
  var year = date.substring(0,4);
  return day + "-" + month + "-" + year;
}

function chart_req(chart_sort, fromDate, toDate){
  fromDate = convertDateFromHTML(fromDate);
  toDate = convertDateFromHTML(toDate);
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var chartSheet;
  if(chart_sort == "Day"){
    chartSheet = activeSpreadsheet.getSheetByName("DailyChart");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    chartSheet = activeSpreadsheet.getSheetByName("Biểu đồ ngày");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    draw_chart_day(fromDate, toDate);
  }
  else if(chart_sort == "Month"){
    chartSheet = activeSpreadsheet.getSheetByName("MonthlyChart");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    chartSheet = activeSpreadsheet.getSheetByName("Biểu đồ tháng");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    draw_chart_month(fromDate, toDate);
  }
  else if(chart_sort == "Quarter"){
    chartSheet = activeSpreadsheet.getSheetByName("QuarterlyChart");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    chartSheet = activeSpreadsheet.getSheetByName("Biểu đồ quý");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    draw_chart_quarter(fromDate, toDate);
  }
}

//lay data tu transaction
function getData(fromDate, toDate, nameSheet) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    activeSpreadsheet.insertSheet().setName(nameSheet);
    chartSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
  if(getLanguage() == "EN"){
    chartSheet.getRange("A1").setValue("Date (MM-dd-yyyy)");
    chartSheet.getRange("B1").setValue("Revenue");
    chartSheet.getRange("C1").setValue("Expense");

    var data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  }
  else{
    chartSheet.getRange("A1").setValue("Ngày (Tháng-Ngày-Năm)");
    chartSheet.getRange("B1").setValue("Doanh thu");
    chartSheet.getRange("C1").setValue("Chi phí");

    var data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Giao dịch ngân hàng");
  }
    chartSheet.getRange("A1:C").setHorizontalAlignment("center");
    chartSheet.getRange("A1:C1").setFontWeight("bold");
    chartSheet.setColumnWidth(1, 200);
    var numRow = data_Sheet.getLastRow();
    var date_cur;
    var amount_cur;
    var new_row;
    var numRow_chart;
    for (var i = 2; i <= numRow; i = i + 1) {
        date_cur = data_Sheet.getRange(i, 2).getDisplayValue();
        var year = +date_cur.substring(6, 10);
        var month = +date_cur.substring(3, 5);
        var day = +date_cur.substring(0, 2);
        amount_cur = data_Sheet.getRange(i, 5).getDisplayValue();
        amount_value = data_Sheet.getRange(i, 5).getValue();

        // lay theo dieu kien
        if (compareDate(date_cur, fromDate) && compareDate(toDate, date_cur)) {
            if (amount_value > 0) {
                //Logger.log(amount_cur);
                date_cur = new Date(year, month - 1, day + 1);
                date_cur = Utilities.formatDate(date_cur, "GTM", "MM-dd-yyyy");
                new_row = [date_cur, amount_cur];
                chartSheet.appendRow(new_row);
                numRow_chart = chartSheet.getLastRow();
                chartSheet.getRange(numRow_chart, 2).setFontColor("green");
            } else {
                //Logger.log("expense: "+amount_cur);
                date_cur = new Date(year, month - 1, day + 1);
                date_cur = Utilities.formatDate(date_cur, "GTM", "MM-dd-yyyy");
                amount_cur = amount_cur.substring(1, amount_cur.length);
                new_row = [date_cur, "", amount_cur];
                chartSheet.appendRow(new_row);
                numRow_chart = chartSheet.getLastRow();

                chartSheet.getRange(numRow_chart, 3).setFontColor("red");
            }
        }
    }
}

// xu ly du lieu truoc
function handleDataForDay() {
    //them chart theo ngay
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var dailyChartSheet = activeSpreadsheet.getSheetByName("DailyChart");
  if(getLanguage() == "VN") dailyChartSheet = activeSpreadsheet.getSheetByName("Biểu đồ ngày");

    dailyChartSheet
        .getRange(2, 2, dailyChartSheet.getLastRow(), 1)
        .setFontColor("green");
    dailyChartSheet
        .getRange(2, 3, dailyChartSheet.getLastRow(), 1)
        .setFontColor("red");
    // xử lý amount format
    var formats = [["#,###"]];
    dailyChartSheet
        .getRange(2, 2, dailyChartSheet.getLastRow(), 2)
        .setNumberFormat(formats);

    //combine
    var last_row = dailyChartSheet.getLastRow();
    for (var i = 2; i <= last_row; i++) {
        // i la gia tri duoc cong gop
        var dateCurr = dailyChartSheet.getRange(i, 1).getValue();
        // dieu kien dung
        if (dateCurr == "") {
            break;
        }
        //vong lap xu ly combine theo ngày
        while (true) {
            // j sau khi cong se bi xoa
            var dateNext = dailyChartSheet.getRange(i + 1, 1).getValue();

            if (checkTheSameDay(dateCurr, dateNext)) {
                //Logger.log('Combine at :'+dateCurr);

                //lay du lieu gop
                let new_revenue =
                    dailyChartSheet.getRange(i, 2).getValue() +
                    dailyChartSheet.getRange(i + 1, 2).getValue();
                let new_expense =
                    dailyChartSheet.getRange(i, 3).getValue() +
                    dailyChartSheet.getRange(i + 1, 3).getValue();

                //sua du lieu
                dailyChartSheet.getRange(i, 2).setValue(new_revenue);
                dailyChartSheet.getRange(i, 3).setValue(new_expense);

                // xoa cot cu
                dailyChartSheet.deleteRow(i + 1);
            } else {
                break;
            }
        }
    }
}

// ham ve chart
function chartForDay(fromDate, toDate) {
  var language = getLanguage();
    var data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DailyChart");
  if(language == "VN") data_Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Biểu đồ ngày");
    var data_chart = data_Sheet.getRange("A1:C");
    var hAxisOptions = {
        slantedText: true,
        slantedTextAngle: 60,
        gridlines: {
            count: 12,
        },
    };
    var lineChartBuilder = data_Sheet.newChart().asColumnChart();
    var chart = lineChartBuilder
        .addRange(data_chart)
        .setPosition(2, 6, 0, 0)
        .setTitle("User's Income from " + fromDate + " to " + toDate)
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.RIGHT)
        .setOption("hAxis", hAxisOptions)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("hAxis", { title: "Date (mm-dd-yyyy)" })
        .setOption("series", {
            0: { color: "blue", labelInLegend: "Revenue" },
            1: { color: "red", labelInLegend: "Expense" },
        })
        .build();
  if(language == "VN")
    chart = lineChartBuilder
        .addRange(data_chart)
        .setPosition(2, 6, 0, 0)
        .setTitle("Thu chi người dùng từ ngày " + fromDate + " đến ngày " + toDate)
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.RIGHT)
        .setOption("hAxis", hAxisOptions)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("hAxis", { title: "Ngày (tháng-ngày-năm)" })
        .setOption("series", {
            0: { color: "blue", labelInLegend: "Doanh thu" },
            1: { color: "red", labelInLegend: "Chi phí" },
        })
        .build();

    data_Sheet.insertChart(chart);
}

// MM-dd-yyyy
function checkTheSameDay(date01, date02) {
    return date01.valueOf() == date02.valueOf();
}

// ham main
function draw_chart_day(fromDate, toDate) {
    // var fromDate = '15-07-2021';
    // var toDate ='23-07-2021';
    if (!compareDate(toDate, fromDate)) {
      if(getLanguage() == "VN")
        SpreadsheetApp.getUi().alert(
            "Bạn vừa nhập ngày bắt đầu sau ngày kết thúc \n Vui lòng nhập lại"
        );
      else
        SpreadsheetApp.getUi().alert(
            "You just enter from Date after to Date \n Please re-enter"
        );
        chart_html();
    } else{
      if(getLanguage() == "VN")
        getData(fromDate, toDate, "Biểu đồ ngày");
      else getData(fromDate, toDate, "DailyChart");
        handleDataForDay();
        chartForDay(fromDate, toDate);
    }
}

//==============================================================================================================================
// ve chart theo thang========================================================================================================
function draw_chart_month(fromDate, toDate) {
  if(getLanguage() == "EN")
    getData(fromDate, toDate, "MonthlyChart");
  else getData(fromDate, toDate, "Biểu đồ tháng");
    handleDataForMonth();
    chart_month(fromDate, toDate);
}

function handleDataForMonth() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var monthlyChartSheet = activeSpreadsheet.getSheetByName("MonthlyChart");
  if(getLanguage() == "VN")
    monthlyChartSheet = activeSpreadsheet.getSheetByName("Biểu đồ tháng");

    monthlyChartSheet
        .getRange(2, 2, monthlyChartSheet.getLastRow(), 1)
        .setFontColor("green");
    monthlyChartSheet
        .getRange(2, 3, monthlyChartSheet.getLastRow(), 1)
        .setFontColor("red");
    // xử lý amount format
    var formats = [["#,###"]];
    monthlyChartSheet
        .getRange(2, 2, monthlyChartSheet.getLastRow(), 2)
        .setNumberFormat(formats);

    //combine
    var last_row = monthlyChartSheet.getLastRow();
    for (var i = 2; i <= last_row; i++) {
        // i la gia tri duoc cong gop
        var dateCurr = monthlyChartSheet.getRange(i, 1).getValue();

        // dieu kien dung
        if (dateCurr == "") {
            break;
        }

        //vong lap xu ly combine theo ngày
        while (true) {
            // j sau khi cong se bi xoa
            var dateNext = monthlyChartSheet.getRange(i + 1, 1).getValue();

            if (checkTheSameMonth(dateCurr, dateNext)) {
                //Logger.log('Combine at :'+dateCurr);

                //lay du lieu gop
                let new_revenue =
                    monthlyChartSheet.getRange(i, 2).getValue() +
                    monthlyChartSheet.getRange(i + 1, 2).getValue();
                let new_expense =
                    monthlyChartSheet.getRange(i, 3).getValue() +
                    monthlyChartSheet.getRange(i + 1, 3).getValue();

                //sua du lieu
                monthlyChartSheet.getRange(i, 2).setValue(new_revenue);
                monthlyChartSheet.getRange(i, 3).setValue(new_expense);

                // xoa cot cu
                monthlyChartSheet.deleteRow(i + 1);
            } else {
                // sua thong tin date
                var date01 = monthlyChartSheet.getRange(i, 1).getDisplayValue();
                var year = date01.substring(6, 10);
                var month = date01.substring(0, 2);
                let date_trans = month + " - " + year;
                Logger.log(date_trans);
                monthlyChartSheet.getRange(i, 1).setValue(date_trans);
                break;
            }
        }
    }
}

function chart_month(fromMonth, toMonth) {
    fromMonth = getMonthYear(fromMonth);
    toMonth = getMonthYear(toMonth);
    var data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MonthlyChart");
  if(getLanguage() == "VN") data_Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Biểu đồ tháng");
    var data_chart = data_Sheet.getRange("A1:C");
    var hAxisOptions = {
        slantedText: true,
        slantedTextAngle: 60,
        gridlines: {
            count: 12,
        },
    };
    var lineChartBuilder = data_Sheet.newChart().asColumnChart();
    var chart = lineChartBuilder
        .addRange(data_chart)
        .setPosition(2, 6, 0, 0)
        .setTitle("User's Income from " + fromMonth + " to " + toMonth)
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.RIGHT)
        .setOption("hAxis", hAxisOptions)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("hAxis", { title: "Month (MM-yyyy)" })
        .setOption("series", {
            0: { color: "blue", labelInLegend: "Revenue" },
            1: { color: "red", labelInLegend: "Expense" },
        })
        .build();
  if(getLanguage() == "VN")
    chart = lineChartBuilder
        .addRange(data_chart)
        .setPosition(2, 6, 0, 0)
        .setTitle("Thu chi của người dùng từ " + fromMonth + " đến " + toMonth)
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.RIGHT)
        .setOption("hAxis", hAxisOptions)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("hAxis", { title: "Tháng (tháng-năm)" })
        .setOption("series", {
            0: { color: "blue", labelInLegend: "Doanh thu" },
            1: { color: "red", labelInLegend: "Chi phí" },
        })
        .build();
  
    data_Sheet.insertChart(chart);
}

// MM-dd-yyyy
function checkTheSameMonth(date01, date02) {
    //kiểm tra format của date khi checkmonth
    if (date01 == date01.valueOf()) {
        date01 = convertDate2(date01);
    }
    if (date02 == date02.valueOf()) {
        date02 = convertDate2(date02);
    }
    return (
        date01.getMonth() == date02.getMonth() &&
        date01.getYear() == date02.getYear()
    );
}

// MM-dd-yyyy
function convertDate2(date01) {
    var year = date01.substring(6, 10);
    var day = date01.substring(3, 5);
    var month = date01.substring(0, 2);
    let date_trans = new Date(year, month - 1, day);
    //date_trans = Utilities.formatDate(date_trans, "GTM", "MM-dd-yyyy");
    return date_trans;
}

// dd-MM-yyyy
function getMonthYear(date01) {
    var year = date01.substring(6, 10);
    var month = date01.substring(3, 5);
    let date_trans = month + " - " + year;
    return date_trans;
}

//==============================================================================================================================
// ve chart theo quý========================================================================================================

function draw_chart_quarter(fromDate, toDate) {
  if(getLanguage() == "EN")
    getData(fromDate, toDate, "QuarterlyChart");
  else getData(fromDate, toDate, "Biểu đồ quý"); 
    handleDataForQuarter();
    chart_quarter(fromDate, toDate);
}

function handleDataForQuarter() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var quarterlyChartSheet =
        activeSpreadsheet.getSheetByName("QuarterlyChart");
  if(getLanguage() == "VN") quarterlyChartSheet =
        activeSpreadsheet.getSheetByName("Biểu đồ quý");

    quarterlyChartSheet
        .getRange(2, 2, quarterlyChartSheet.getLastRow(), 1)
        .setFontColor("green");
    quarterlyChartSheet
        .getRange(2, 3, quarterlyChartSheet.getLastRow(), 1)
        .setFontColor("red");
    // xử lý amount format
    var formats = [["#,###"]];
    quarterlyChartSheet
        .getRange(2, 2, quarterlyChartSheet.getLastRow(), 2)
        .setNumberFormat(formats);

    //combine
    var last_row = quarterlyChartSheet.getLastRow();
    for (var i = 2; i <= last_row; i++) {
        // i la gia tri duoc cong gop
        var dateCurr = quarterlyChartSheet.getRange(i, 1).getValue();

        // dieu kien dung
        if (dateCurr == "") {
            break;
        }

        //vong lap xu ly combine theo ngày
        while (true) {
            // j sau khi cong se bi xoa
            var dateNext = quarterlyChartSheet.getRange(i + 1, 1).getValue();

            if (checkTheSameQuarter(dateCurr, dateNext)) {
                //Logger.log('Combine at :'+dateCurr);

                //lay du lieu gop
                let new_revenue =
                    quarterlyChartSheet.getRange(i, 2).getValue() +
                    quarterlyChartSheet.getRange(i + 1, 2).getValue();
                let new_expense =
                    quarterlyChartSheet.getRange(i, 3).getValue() +
                    quarterlyChartSheet.getRange(i + 1, 3).getValue();

                //sua du lieu
                quarterlyChartSheet.getRange(i, 2).setValue(new_revenue);
                quarterlyChartSheet.getRange(i, 3).setValue(new_expense);

                // xoa cot cu
                quarterlyChartSheet.deleteRow(i + 1);
            } else {
                // sua thong tin date
                var date01 = quarterlyChartSheet
                    .getRange(i, 1)
                    .getDisplayValue();
                var year = date01.substring(6, 10);
                var month = date01.substring(0, 2);
              if(getLanguage() == "EN"){
                if (month == "01" || month == "02" || month == "03") {
                    var date_trans = "Quarter I - " + year;
                } else if (month == "04" || month == "05" || month == "06") {
                    var date_trans = "Quarter II - " + year;
                } else if (month == "08" || month == "09" || month == "07") {
                    var date_trans = "Quarter III - " + year;
                } else {
                    var date_trans = "Quarter IV - " + year;
                }
              }
              else{
                if (month == "01" || month == "02" || month == "03") {
                    var date_trans = "Quý I - " + year;
                } else if (month == "04" || month == "05" || month == "06") {
                    var date_trans = "Quý II - " + year;
                } else if (month == "08" || month == "09" || month == "07") {
                    var date_trans = "Quý III - " + year;
                } else {
                    var date_trans = "Quý IV - " + year;
                }
              }
                quarterlyChartSheet.getRange(i, 1).setValue(date_trans);
                break;
            }
        }
    }
}

function chart_quarter(fromMonth, toMonth) {
    fromMonth = getMonthYearToQuarter(fromMonth);
    toMonth = getMonthYearToQuarter(toMonth);
    var data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QuarterlyChart");
  if(getLanguage() == "VN")
    data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Biểu đồ quý");
    var data_chart = data_Sheet.getRange("A1:C");
    var hAxisOptions = {
        slantedText: true,
        slantedTextAngle: 20,
        gridlines: {
            count: 12,
        },
    };
    var lineChartBuilder = data_Sheet.newChart().asColumnChart();
    var chart = lineChartBuilder
        .addRange(data_chart)
        .setPosition(2, 6, 0, 0)
        .setTitle("User's Income from " + fromMonth + " to " + toMonth)
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.RIGHT)
        .setOption("hAxis", hAxisOptions)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("hAxis", { title: "Quarter (Quarter-yyyy)" })
        .setOption("series", {
            0: { color: "blue", labelInLegend: "Revenue" },
            1: { color: "red", labelInLegend: "Expense" },
        })
        .build();
  if(getLanguage() == "VN")
    chart = lineChartBuilder
        .addRange(data_chart)
        .setPosition(2, 6, 0, 0)
        .setTitle("Chi phí người dùng từ " + fromMonth + " đến " + toMonth)
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.RIGHT)
        .setOption("hAxis", hAxisOptions)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("hAxis", { title: "Quý (quý-năm)" })
        .setOption("series", {
            0: { color: "blue", labelInLegend: "Doanh thu" },
            1: { color: "red", labelInLegend: "Chi phí" },
        })
        .build();

    data_Sheet.insertChart(chart);
}

// MM-dd-yyyy
function checkTheSameQuarter(date01, date02) {
    //kiểm tra format của date khi checkmonth
    if (date01 == date01.valueOf()) {
        date01 = convertDate2(date01);
    }
    if (date02 == date02.valueOf()) {
        date02 = convertDate2(date02);
    }
    if (date01.getYear() == date02.getYear()) {
        if (
            (date01.getMonth() == "00" ||
                date01.getMonth() == "01" ||
                date01.getMonth() == "02") &&
            (date02.getMonth() == "00" ||
                date02.getMonth() == "01" ||
                date02.getMonth() == "02")
        ) {
            return true;
        } else if (
            (date01.getMonth() == "03" ||
                date01.getMonth() == "05" ||
                date01.getMonth() == "04") &&
            (date02.getMonth() == "04" ||
                date02.getMonth() == "05" ||
                date02.getMonth() == "03")
        ) {
            return true;
        } else if (
            (date01.getMonth() == "07" ||
                date01.getMonth() == "08" ||
                date01.getMonth() == "06") &&
            (date02.getMonth() == "07" ||
                date02.getMonth() == "08" ||
                date02.getMonth() == "06")
        ) {
            return true;
        } else if (
            (date01.getMonth() == "10" ||
                date01.getMonth() == "11" ||
                date01.getMonth() == "09") &&
            (date02.getMonth() == "10" ||
                date02.getMonth() == "11" ||
                date02.getMonth() == "09")
        ) {
            return true;
        } else return false;
    } else {
        return false;
    }
}

// dd-MM-yyyy
function getMonthYearToQuarter(date01) {
    var year = date01.substring(6, 10);
    var month = date01.substring(3, 5);
  if(getLanguage() == "EN"){
    if (month == "01" || month == "02" || month == "03") {
        var date_trans = "Quarter I - " + year;
    } else if (month == "04" || month == "05" || month == "06") {
        var date_trans = "Quarter II - " + year;
    } else if (month == "08" || month == "09" || month == "07") {
        var date_trans = "Quarter III - " + year;
    } else {
        var date_trans = "Quarter IV - " + year;
    }
  }
  else{
    if (month == "01" || month == "02" || month == "03") {
        var date_trans = "Quý I - " + year;
    } else if (month == "04" || month == "05" || month == "06") {
        var date_trans = "Quý II - " + year;
    } else if (month == "08" || month == "09" || month == "07") {
        var date_trans = "Quý III - " + year;
    } else {
        var date_trans = "Quý IV - " + year;
    }
  }
    return date_trans;
}