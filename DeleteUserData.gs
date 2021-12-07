function myFunction() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function(st){
      if(st.getName() == "説明書" || st.getName() == "QR" || st.getName() == "template_data" || st.getName() == "LOG"){
      }else{
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(st);
      }
    });
}
