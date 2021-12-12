/**
 * 独自メニューの追加
 */
function onOpen(){
  var myMenu = [ //メニュー配列
    {name: "更新開始", functionName: "update_videoLists"},
  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("【動画一覧更新】",myMenu); 
}

/**
 * チャンネル動画アップデート
 */
function update_videoLists() {
  recordVideoLists("（YoutubeチャンネルのID）※準備②", "（動画一覧シートの名前）※準備④");
}

/**
 * 動画一覧アップデート
 */
var recordVideoLists = function(channelId, sheetName) {

  // 既に登録されている一覧を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  sheet.activate();
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
  const rangeValues = range.getValues();

  var result = [];
  var lists = getVideoLists(channelId);

  // Youtubeから取得した動画一覧
  for(var i = 0; i < lists.length; i++) {
    var r = lists[i];

    // 既に登録済みの動画はinsertせずcontinue
    let alreadyFlg = false;
    for (let i = 0; i < rangeValues.length; i++) {
      if (r.videoId === rangeValues[i][4]) {
        alreadyFlg = true;
      }
    }
    if (alreadyFlg === true) {
      continue;
    }

    // 未登録の動画はinsert
    var date = new Date();
    var pub_date = new Date(r.publishedAt);
    var record = ['=ROW()-2',date, pub_date, r.title, r.videoId, 'https://www.youtube.com/watch?v=' + r.videoId, '0'];
    result.push(record);
  }
  insertRecords(result, sheetName, 3); // 3列目からデータを挿入する
}

/**
 * チャンネル情報取得して返す
 */
var getChannelInfo = function(id) {
  var key = '（Google Cloud Platform のAPIKEY）※準備①'; 
  var url = "https://www.googleapis.com/youtube/v3/channels?part=statistics,snippet&id=" + id +"&key=" + key;
  var response = UrlFetchApp.fetch(url);
  var json = JSON.parse(response.getContentText());
  var item = json.items[0];
  var channel_id = item.id;
  var statistics = item.statistics;
  var viewCount = statistics.viewCount;
  var subscriberCount = statistics.subscriberCount;
  var videoCount = statistics.videoCount;
  var snippet = item.snippet;
  var title = snippet.title;
  var description = snippet.description;
  var customUrl = snippet.customUrl;
  var publishedAt = snippet.publishedAt;

  var res = {channel_id: channel_id, viewCount: viewCount, subscriberCount: subscriberCount, videoCount: videoCount, 
    title: title, description:description, customUrl:customUrl, publishedAt: publishedAt};
  return res
}

/**
 * シートを作成して動画一覧を記載
 */
var insertRecords = function(arrData, sheetName, startRow){
  let rows = arrData.length;
  if (rows > 0) {
    let cols = arrData[0].length;
    let ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシートを取得
    let sheet = ss.getSheetByName(sheetName); // sheetName という名前のシートを取得
    if (sheet == null) {
      sheet = mySS.insertSheet(sheetName); // 新シートの追加
      Logger.log("シートを追加しました:" + sheetName);
    }
    sheet.insertRows(startRow, rows); // 空行の作成
    sheet.getRange(startRow, 1, rows, cols).setValues(arrData); // データの書き込み
  }
}

/**
 * 動画一覧を取得
 */
var getVideoLists = function(channelId) {
  let lists = [];
  let pageToken = "";
  for(let page = 0; page < 4; page++) {
    let url = "https://www.googleapis.com/youtube/v3/search";
    url += "?part=snippet&maxResults=50&order=date&type=video";
    url += "&pageToken=" + pageToken + "&channelId=" + channelId +"&key=" + 'AIzaSyBwq7ieCfIQp5E72H5k8mJhpYOEUJ_-JcM';
    let response = UrlFetchApp.fetch(url);
    let json = JSON.parse(response.getContentText());
    if (json.kind == "youtube#searchListResponse") {
      let items = json.items;
      for(let i = 0; i < items.length; i++) {
        let item = items[i];
        if (item.kind == "youtube#searchResult") {
          let id = item.id;
          if (id.kind == "youtube#video") {
            let snippet = item.snippet;
            let video = {videoId: id.videoId, publishedAt: snippet.publishedAt, title: snippet.title};
            lists.push(video);
          }
        }
      }    
    }
    if ("nextPageToken" in json) {
      pageToken = json.nextPageToken;
    } else {
      break;
    }
  }
  return lists
}

/**
 * 動画を１つSlackにポストして投稿
 */
function todayMovie() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('（動画一覧のシート名）※準備④');
  sheet.activate();
  const range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
  const rangeValues = range.getValues();

  let alreadyFlg = false;
  for (let i = 0; i < rangeValues.length; i++) {
    if (alreadyFlg === true) {
      continue;
    }
    Logger.log(rangeValues[i][6]);
    if (rangeValues[i][6] === 0.0) {
      alreadyFlg = true;
      sheet.getRange(3 + i, 7).setValue(1);
      postSlack(rangeValues[i]);
    }
  }
}

/**
 * Slack投稿
 */
function postSlack(issue) {
  if (issue.length <= 0) {
    return;
  }
  console.log(issue.length);

  const postUrl  = '（SlackのWebhookURL）※準備③';
  // slackのincoming webhook用
  const username = 'botbot';  // 通知時に表示されるユーザー名
  const icon     = ':hatching_chick:'; // 通知時に表示されるアイコン
  const subject = '【本日の動画ですよ～】'; //この辺はご自由に
  const body    = '' + subject + '\n' + createPostMessage(issue) + '\n';
  const jsonData =　{
    'username'  : username,
    'icon_emoji': icon,
    'text'      : body
  };
  const options =　{
    'method'     : 'post',
    'contentType': 'application/json',
    'payload'    : JSON.stringify(jsonData)
  };

  UrlFetchApp.fetch(postUrl, options);
}

/**
 * 投稿メッセージ作成
 */
function createPostMessage(issues) {
  let message = '';
  message += formatDate(new Date(issues[2])) + '配信\n';;
  message += issues[3] + '\n';
  message += issues[5] + '\n';
  return message;
}

/**
 * 日付フォーマット作成
 */
function formatDate(date) {
    var format = 'YYYY-MM-DD';
    format     = format.replace(/YYYY/g, date.getFullYear());
    format     = format.replace(/MM/g, ('0' + (date.getMonth() + 1)).slice(-2));
    format     = format.replace(/DD/g, ('0' + date.getDate()).slice(-2));
    return format;
}
