  let status = "";
  let name = "";
  let total = 0;
  let current = 0;
  let UserInfo = "";
  let location = "";
  let stampName = "";
  let getDate = "";
  let stat = "";
  let found = false;
  var LogSheet = SpreadsheetApp.getActive().getSheetByName('LOG');

  function doGet(e) {
    var no = e.parameter.no;
    if(no == "GET"){
      return getList();
    }
    var datetime = new Date();
    UserInfo = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.getActive().getSheetByName('QR');
    LogSheet.appendRow(['QR読み取り', datetime, no, UserInfo]);
    var res = "";
    var userData = makeData(UserInfo);
    var lastRow = userData.getLastRow();
    total = userData.getRange(2, 7).getValue();
    current = userData.getRange(2, 8).getValue();

    for (var i = 2; i <= lastRow; i++) {
      if (userData.getRange(i, 1).getValue() == no) {
        location = userData.getRange(i, 3).getValue();
        name = userData.getRange(i, 4).getValue();
        time = userData.getRange(i, 2).getValue();
        found = true;
        if (userData.getRange(i, 5).getValue() == "YES") {
          userData.getRange(i, 5).setValue("YES");
          status = "すでにスタンプをGETしています!";
          stat = "このスタンプは" + location + "にある" + name + "です!";
        } else {
          userData.getRange(i, 5).setValue("YES");
          userData.getRange(i, 2).setValue(datetime);
          status = "スタンプGET!";
          stat = location + "にある" + name + "をGETしました!";
          current = current + 1;
          LogSheet.appendRow(['スタンプGET', datetime, no, UserInfo]);
        }
        break;
      }
    }
    // 実行日時＆メールアドレスセット
    if (!found) {
      var t = HtmlService.createTemplateFromFile("nodata");
      t.status = "エラーです";
      t.total = total;
      t.current = current;
      t.UserInfo = UserInfo;
      t.stat = "そのスタンプは登録されていません。";
      LogSheet.appendRow(['スタンプ登録なし', datetime, no, UserInfo]);
      return t.evaluate().setTitle("エラーです" + " | 梅田キャンパススタンプラリー").setFaviconUrl("https://lh5.googleusercontent.com/iwiSXn6lrfBDyQTHpti1ndSJhnSSBexm4qHMDCSx1xIgk4YoT4dQUj7yjgGTbifn8qVokXwlLtB7wf7rw9Yz=w2880-h1578?.png").addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    var t = HtmlService.createTemplateFromFile("getstamp");
    t.status = status;
    t.total = total;
    t.current = current;
    t.UserInfo = UserInfo;
    t.name = name;
    t.location = location;
    t.time = time;
    t.stat = stat;
    return t.evaluate().setTitle(status + " | 梅田キャンパススタンプラリー").setFaviconUrl("https://lh5.googleusercontent.com/iwiSXn6lrfBDyQTHpti1ndSJhnSSBexm4qHMDCSx1xIgk4YoT4dQUj7yjgGTbifn8qVokXwlLtB7wf7rw9Yz=w2880-h1578?.png").addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  function getList(){
    var datetime = new Date();
    UserInfo = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.getActive().getSheetByName('QR');
    LogSheet.appendRow(['リスト表示', datetime, "", UserInfo]);
    var res = "";
    var userData = makeData(UserInfo);
    var lastRow = userData.getLastRow();
    total = userData.getRange(2, 7).getValue();
    current = userData.getRange(2, 8).getValue();
    let text = "";
    for (var i = 2; i <= lastRow; i++) {
      if (userData.getRange(i, 1).getValue() == "") {
        break;
      }
        location = userData.getRange(i, 3).getValue();
        name = userData.getRange(i, 4).getValue();
        time = userData.getRange(i, 2).getValue();
        if (userData.getRange(i, 5).getValue() == "YES") {
          text = text+name+" | "+location+"%nnn%";
        }
    }
    status = "取得したスタンプ一覧";
    stat = total+"個中"+current+"個獲得!!";
    var t = HtmlService.createTemplateFromFile("list");
    t.status = status;
    t.UserInfo = UserInfo;
    t.stat = stat;
    t.text = text;
    return t.evaluate().setTitle("GETしたスタンプ一覧" + " | 梅田キャンパススタンプラリー").setFaviconUrl("https://lh5.googleusercontent.com/iwiSXn6lrfBDyQTHpti1ndSJhnSSBexm4qHMDCSx1xIgk4YoT4dQUj7yjgGTbifn8qVokXwlLtB7wf7rw9Yz=w2880-h1578?.png").addMetaTag('viewport', 'width=device-width, initial-scale=1');

  }

  function getStatus() {
    return this.status;
  }

  function getTotal() {
    return this.total;
  }

  function getCurrent() {
    return this.current;
  }

  function getUserInfo() {
    return this.UserInfo;
  }

  function makeData(name) {
    var datetime = new Date();
    //同じ名前のシートがなければ作成
    var sheet = SpreadsheetApp.getActive().getSheetByName(name)
    if (sheet) return sheet
    var sheet_1 = SpreadsheetApp.getActive().getSheetByName("template_data"); //temp_sheetというシートがある前提
    var sheet = sheet_1.copyTo(SpreadsheetApp.getActive());
    LogSheet.appendRow(['データ作成', datetime, "NULL", UserInfo]);
    sheet.setName(name);
    return sheet;
  }
