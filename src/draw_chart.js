function getData(fromDate, toDate){
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var chartSheet = activeSpreadsheet.getSheetByName("Chart");
    if (chartSheet != null) {
      activeSpreadsheet.deleteSheet(chartSheet);
    }
    activeSpreadsheet.insertSheet().setName('Chart');
    chartSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chart");
    chartSheet.getRange("A1").setValue("Date (MM-dd-yyyy)");
    chartSheet.getRange("B1").setValue("Amount");
    chartSheet.getRange('A1:B').setHorizontalAlignment('center');
    chartSheet.getRange('A1:B1').setFontWeight('bold');
    chartSheet.setColumnWidth(1,200);
    var data_Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
    var numRow = data_Sheet.getLastRow();
    var date_cur;
    var amount_cur;
    var new_row;
    var numRow_chart;
    for(var i = 2; i <= numRow; i = i+1){
      date_cur = data_Sheet.getRange(i, 2).getDisplayValue();
      var year = +date_cur.substring(6, 10);
      var month = +date_cur.substring(3, 5);
      var day = +date_cur.substring(0, 2);
      amount_cur = data_Sheet.getRange(i, 5).getDisplayValue();
      if(compareDate(date_cur, fromDate) && compareDate(toDate, date_cur)){
        date_cur = new Date(year, month - 1, day + 1);
        date_cur = Utilities.formatDate(date_cur, "GTM", "MM-dd-yyyy");
        new_row = [date_cur, amount_cur];
        chartSheet.appendRow(new_row);
        numRow_chart = chartSheet.getLastRow();
        if(data_Sheet.getRange(i, 5).getValue() > 0) chartSheet.getRange(numRow_chart,2).setFontColor("green");
        else chartSheet.getRange(numRow_chart,2).setFontColor("red");
      }
    }
  }
  
  function draw_chart(){
    var res = SpreadsheetApp.getUi().prompt(`Bạn muốn vẽ biểu đồ thu chi trong khoảng thời gian nào?
    ví dụ: 31-07-2021 to 31-08-2021
    Lưu ý: biểu đồ chỉ có thể vẽ dựa trên dữ liệu ở Sheet Transactions`);
    var res = res.getResponseText();
    var fromDate = res.substring(0, 10);
    var toDate = res.substring(14, 24);
    if(!compareDate(toDate, fromDate)) {
      SpreadsheetApp.getUi().alert("Bạn vừa nhập ngày bắt đầu sau ngày kết thúc \n Vui lòng nhập lại");
      draw_chart();
    }
    else if(res.length != 24 && res.length > 0){
      SpreadsheetApp.getUi().alert("Bạn vừa nhập sai cú pháp. \n Vui lòng nhập lại");
      draw_chart();
    }
    else if(res.length == 24){
      getData(fromDate, toDate);
      chart(fromDate, toDate);
    }
  }
  
  function chart(fromDate, toDate){
    var data_Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chart");
    var data_chart = data_Sheet.getRange("A1:B");
    var hAxisOptions = {
      slantedText: true,
      slantedTextAngle: 60,
      gridlines: {
        count: 12
      },
    };
    var lineChartBuilder = data_Sheet.newChart().asLineChart();
    var chart = lineChartBuilder
      .addRange(data_chart)
      .setPosition(2, 4, 0, 0)
      .setTitle("User's Income from " + fromDate + " to " + toDate)
      .setNumHeaders(2)
      .setLegendPosition(Charts.Position.RIGHT)
      .setOption('hAxis', hAxisOptions)
      .setOption("useFirstColumnAsDomain", true)
      .setOption("hAxis", {title: "Date (mm-dd-yyyy)"})
      .setOption("series", {
        0: {color: "green", labelInLegend: "Income"},
      })
      .build();
      
    data_Sheet.insertChart(chart);  
    
  }