  let status = "";
  let name = "";
  let total = 0;
  let current = 0;
  let UserInfo = "";
  let location = "";
  let stampName = "";
  let getDate = "";
  let stat = "";
  let time = "";
  let found = false;
  var LogSheet = SpreadsheetApp.getActive().getSheetByName('LOG');

  function doGet(e) {
    var no = e.parameter.no; //?no=の部分を取得
    /* 取得済みスタンプリストを表示 */
    if (no == "GET") {
      return getList();
    }
    //ログ用現在時刻を取得
    var datetime = new Date();
    //ユーザーを識別するためのトークンを取得
    UserInfo = Session.getTemporaryActiveUserKey();
    var ss = SpreadsheetApp.getActive().getSheetByName('QR');
    //Log: Read QR Code
    LogSheet.appendRow(['QR読み取り', datetime, no, UserInfo]);
    //出力
    var res = "";
    //トークンからユーザーデータを取得
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
          let date = new Date(time);
          time = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy.MM.dd. HH:mm');
        } else {
          userData.getRange(i, 5).setValue("YES");
          userData.getRange(i, 2).setValue(datetime);
          status = "スタンプGET!";
          stat = location + "にある" + name + "をGETしました!";
          current = current + 1;
          let date = new Date(datetime);
          time = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy.MM.dd. HH:mm');
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
      return t.evaluate().setTitle("エラーです" + " | 梅田キャンパススタンプラリー").setFaviconUrl("https://drive.google.com/uc?id=10aiMpGWRiuimSM_FJcTUJKoVfLbdbSd_&.png").addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
    return t.evaluate().setTitle(status + " | 梅田キャンパススタンプラリー").setFaviconUrl("https://drive.google.com/uc?id=10aiMpGWRiuimSM_FJcTUJKoVfLbdbSd_&.png").addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  function getList() {
    var datetime = new Date();
    UserInfo = Session.getTemporaryActiveUserKey();
    var ss = SpreadsheetApp.getActive().getSheetByName('QR');
    LogSheet.appendRow(['リスト表示', datetime, "", UserInfo]);
    var res = "";
    var userData = makeData(UserInfo);
    var lastRow = userData.getLastRow();
    total = userData.getRange(2, 7).getValue();
    current = userData.getRange(2, 8).getValue();
    let text = "";
    let html = "";
    for (var i = 2; i <= lastRow; i++) {
      if (userData.getRange(i, 1).getValue() == "") {
        break;
      }
      location = userData.getRange(i, 3).getValue();
      name = userData.getRange(i, 4).getValue();
      time = userData.getRange(i, 2).getValue();
      if (userData.getRange(i, 5).getValue() == "YES") {
        /* text = text+name+" | "+location+"%nnn%"; */
        let date = new Date(time);
        time = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy.MM.dd. HH:mm');
        html = html + '<style>.' + name + ':after{content: \'in ' + location + '\';}</style><div class="col-12 col-md-4"><div class="card"><div class="card-body"><div class="stamp ' + name + '">' + time + '<br>' + name + '</div></div></div></div>';
      } else {
        html = html + '<div class="col-12 col-md-4"><div class="card"><div class="card-body"><div style="width:200px; height:200px;"></div></div></div></div>';
      }
    }
    status = "取得したスタンプ一覧";
    stat = total + "個中" + current + "個獲得!!";
    var t = HtmlService.createTemplateFromFile("list");
    t.status = status;
    t.UserInfo = UserInfo;
    t.stat = stat;
    t.text = text;
    t.name = name;
    t.html = html;
    return t.evaluate().setTitle("GETしたスタンプ一覧" + " | 梅田キャンパススタンプラリー").setFaviconUrl("https://drive.google.com/uc?id=10aiMpGWRiuimSM_FJcTUJKoVfLbdbSd_&.png").addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

  function countUsers(){
    var sheet = SpreadsheetApp.getActive().getSheetByName("説明書");
    sheet.getRange(4,1).setValue(SpreadsheetApp.getActive().getNumSheets() - 4+"人");
    SpreadsheetApp.getActiveSpreadsheet().toast('更新しました',"ユーザーカウント",8);
  }

  function countClears(){
    SpreadsheetApp.getActiveSpreadsheet().toast('更新しています',"クリア数",8);
    var clear = 0;
    var sheet = SpreadsheetApp.getActive().getSheetByName("説明書");
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function(st){
      if(st.getRange(2,8).getValue() > 6){
        clear++;
      }
    });
    sheet.getRange(8,1).setValue(clear+"人");
    SpreadsheetApp.getActiveSpreadsheet().toast('更新しました',"クリア数",8);
  }
  function countAll(){
    countUsers();
    countClears();
    var sheet = SpreadsheetApp.getActive().getSheetByName("説明書");
    let date = new Date();
    time = Utilities.formatDate(date, 'Asia/Tokyo', 'MM/dd HH:mm:ss');
    sheet.getRange(13,1).setValue(time);
  }
