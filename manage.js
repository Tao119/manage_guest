var channel_token = "line_access_token"　//LINEアクセストークン
var url = "https://api.line.me/v2/bot/message/reply"

function doPost(e) {

  let maximum = 5; //ゲスト上限数を設定

  let spreadsheet = SpreadsheetApp.openById("sheet_id");
  let sheet = spreadsheet.getActiveSheet(); //シートの取得

  var json = e.postData.contents
  var events = JSON.parse(json).events;　//イベントの取得

  events.forEach(function (event) {　//イベント毎に実行

    if (event.type == "message") {　//メッセージイベントなら続行

      let userMessage = event.message.text;
      let userMessage2 = userMessage.split("\n");　//メッセージを行ごとに整理

      let replyMessage = "コマンドを実行できませんでした。\n表記や日程を再確認してください。\n使い方を忘れた場合はhelpと発言してください！";　//デフォルト応答メッセージ
      let replyMessage1 = "日付:" + userMessage2[0] + "\n名前:" + userMessage2[1] + "\n所属:" + userMessage2[2] + "\nで登録しました";　//登録成功時
      let replyMessage2 = "その日程のゲストは規定人数に達しております";　//登録失敗時

      if (userMessage2[0].indexOf('help') != -1) {　//helpに応答しないように
        replyMessage = "";
      }

      for (i = 0; i < 30; i++) {

        let cell = sheet.getRange(1, i + 2); //日付探索用セル
        if(userMessage2[1]!=null){

        if (userMessage2[0] == "キャンセル") {　//キャンセルの時

          if (userMessage2[1] == cell.getValue()) { //日付探索

            for (j = 0; j < 5; j++) {　//ゲスト探索

              if (sheet.getRange(j + 2, i + 2).getValue() == userMessage2[2] + "(" + userMessage2[3] + ")") {

                sheet.getRange(j + 2, i + 2).clearContent();
                replyMessage = "キャンセルしました"　//キャンセル処理
              }
            }
          }
        }
        else if (userMessage2[0] == cell.getValue()) {　//登録の時

          let row = 1;
          while (sheet.getRange(row, i + 2).getValue().length > 0) {　//空白でない最初の行を探索

            row++;
          }
          if (row >= maximum + 2) { //制限人数枠+1行(日付欄)まで埋まっていたら予約いっぱい

            replyMessage = replyMessage2; //🙇‍♀️
          }
          else { //登録可能

            sheet.getRange(row, i + 2).setValue(userMessage2[1] + "(" + userMessage2[2] + ")");
            replyMessage = replyMessage1;　//登録処理
          }
        }
      }
      }

      var message = {　//メッセージ設定
        "replyToken": event.replyToken,
        "messages": [{ "type": "text", "text": replyMessage }]
      };

      var options = { //必要情報
        "method": "post",
        "headers": {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + channel_token
        },
        "payload": JSON.stringify(message)
      };

      UrlFetchApp.fetch(url, options);　//メッセージ送信
    }
  });
}
