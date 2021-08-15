function getData(fromDate, toDate) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var chartSheet = activeSpreadsheet.getSheetByName("Chart");
    if (chartSheet != null) {
        activeSpreadsheet.deleteSheet(chartSheet);
    }
    activeSpreadsheet.insertSheet().setName("Chart");
    chartSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chart");
    chartSheet.getRange("A1").setValue("Date (MM-dd-yyyy)");
    chartSheet.getRange("B1").setValue("Revenue");
    chartSheet.getRange("C1").setValue("Expense");
    chartSheet.getRange("A1:C").setHorizontalAlignment("center");
    chartSheet.getRange("A1:C1").setFontWeight("bold");
    chartSheet.setColumnWidth(1, 200);

    var data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
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
                Logger.log(amount_cur);
                date_cur = new Date(year, month - 1, day + 1);
                date_cur = Utilities.formatDate(date_cur, "GTM", "MM-dd-yyyy");
                new_row = [date_cur, amount_cur];
                chartSheet.appendRow(new_row);
                numRow_chart = chartSheet.getLastRow();
                chartSheet.getRange(numRow_chart, 2).setFontColor("green");
            } else {
                Logger.log("expense: " + amount_cur);
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

function handleData() {
    //them chart theo ngay
    // copy data
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var dailyChartSheet = activeSpreadsheet.getSheetByName("DailyChart");
    if (dailyChartSheet != null) {
        activeSpreadsheet.deleteSheet(dailyChartSheet);
    }
    activeSpreadsheet.insertSheet().setName("DailyChart");
    dailyChartSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DailyChart");

    var dataSheet = activeSpreadsheet.getSheetByName("Chart");

    var data = dataSheet.getRange(1, 1, dataSheet.getLastRow(), 3);
    var dest = dailyChartSheet.getRange(1, 1, dataSheet.getLastRow(), 3);
    data.copyTo(dest);

    dailyChartSheet
        .getRange(2, 2, dataSheet.getLastRow(), 1)
        .setFontColor("green");
    dailyChartSheet
        .getRange(2, 3, dataSheet.getLastRow(), 1)
        .setFontColor("red");
    // xử lý amount format
    var formats = [["#,###"]];
    dailyChartSheet
        .getRange(2, 2, dataSheet.getLastRow(), 2)
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
                Logger.log("Combine at :" + dateCurr);

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

function draw_chart() {
    // var fromDate = '15-07-2021';
    // var toDate ='23-07-2021';

    var res = SpreadsheetApp.getUi()
        .prompt(`Bạn muốn vẽ biểu đồ thu chi trong khoảng thời gian nào?
    ví dụ: 31-07-2021 to 31-08-2021
    Lưu ý: biểu đồ chỉ có thể vẽ dựa trên dữ liệu ở Sheet Transactions`);
    var res = res.getResponseText();
    var fromDate = res.substring(0, 10);
    var toDate = res.substring(14, 24);
    if (!compareDate(toDate, fromDate)) {
        SpreadsheetApp.getUi().alert(
            "Bạn vừa nhập ngày bắt đầu sau ngày kết thúc \n Vui lòng nhập lại"
        );
        draw_chart();
    } else if (res.length != 24 && res.length > 0) {
        SpreadsheetApp.getUi().alert(
            "Bạn vừa nhập sai cú pháp. \n Vui lòng nhập lại"
        );
        draw_chart();
    } else if (res.length == 24) {
        getData(fromDate, toDate);
        handleData();
        chart(fromDate, toDate);
    }
}

function chart(fromDate, toDate) {
    var data_Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DailyChart");
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
            0: { color: "green", labelInLegend: "Revenue" },
            1: { color: "red", labelInLegend: "Expense" },
        })
        .build();

    data_Sheet.insertChart(chart);
}

// MM-dd-yyyy
function checkTheSameDay(date01, date02) {
    return date01.valueOf() == date02.valueOf();
}
