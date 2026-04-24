// ==========================================
// 【事前準備と初期設定】
// 1. Slackアプリの設定
//    - https://api.slack.com/apps にアクセスし、「Create New App」>「From scratch」でアプリを作成。
//    - 左メニュー「OAuth & Permissions」を開き、Scopesの「Bot Token Scopes」に「chat:write」を追加。
//    - 「Install to Workspace」をクリックしてワークスペースにインストールし、「Bot User OAuth Token (xoxb-...)」を取得。
//    - 転送先のSlackチャンネルを開き、作成したアプリをチャンネルに追加（Invite）。
//    - チャンネルID（Cから始まる英数字）をチャンネル詳細ページから取得。
// 2. スプレッドシートの準備
//    - 新規作成し、ID（URLの /d/ と /edit の間の文字列）とシート名を控える。
// 3. GASのスクリプトプロパティ設定
//    - GASエディタ左メニューの「プロジェクトの設定（歯車アイコン）」を開く。
//    - 一番下の「スクリプト プロパティ」で「スクリプト プロパティを追加」をクリックし、以下のキーと値を登録する。
//      - キー: SLACK_BOT_TOKEN    値: 取得したBot Token (xoxb-...)
//      - キー: SLACK_CHANNEL_ID   値: 取得したチャンネルID (C...)
//      - キー: SPREADSHEET_ID     値: スプレッドシートのID
//      - キー: SHEET_NAME         値: 使用するシート名 (例: MailLists)
//  4. チャンネルへAppの招待
//      - 投稿先のチャンネルへ作成したアプリを招待する。
// ==========================================

// 設定値の取得 (スクリプトプロパティから)
const props = PropertiesService.getScriptProperties();
const SLACK_BOT_TOKEN = props.getProperty('SLACK_BOT_TOKEN');
const SLACK_CHANNEL_ID = props.getProperty('SLACK_CHANNEL_ID');
const SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID');
const SHEET_NAME = props.getProperty('SHEET_NAME');

// 検索条件
// ※自分が新規送信したメールも、すべてSlackに転送したい場合は、送信済みも含めるように以下のように変更してください
const SEARCH_QUERY = '{in:inbox in:sent} newer_than:1h';

// 1時間以内に受信したメールのみの場合
// const SEARCH_QUERY = 'in:inbox newer_than:1h';

function main() {
  // スクリプトプロパティの設定漏れチェック
  if (!SLACK_BOT_TOKEN || !SLACK_CHANNEL_ID || !SPREADSHEET_ID || !SHEET_NAME) {
    console.error('スクリプトプロパティが正しく設定されていません。プロジェクトの設定から各値を登録してください。');
    return;
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    console.error('指定されたシートが見つかりません: ' + SHEET_NAME);
    return;
  }

  const lastRow = sheet.getLastRow();
  const processedIds = new Set();
  const threadTsMap = {}; // GmailのスレッドIDと、Slackの親投稿のタイムスタンプ(ts)を紐付ける用

  // 1. スプレッドシートから処理済みのデータを取得
  // A列: Message ID, B列: Gmail Thread ID, C列: Slack ts (タイムスタンプ)
  if (lastRow > 0) {
    const data = sheet.getRange(1, 1, lastRow, 3).getValues();
    data.forEach(row => {
      const msgId = row[0];
      const gmailThreadId = row[1];
      const slackTs = row[2];
      
      if (msgId) processedIds.add(msgId);
      if (gmailThreadId && slackTs && !threadTsMap[gmailThreadId]) {
        threadTsMap[gmailThreadId] = slackTs; // そのスレッドの親となるSlack投稿のtsを記憶
      }
    });
  }

  // 2. Gmailから条件に合致するスレッドを取得
  const threads = GmailApp.search(SEARCH_QUERY);
  const newRows = [];

  // 3. スレッド内のメッセージを古い順に確認
  for (let i = threads.length - 1; i >= 0; i--) {
    const thread = threads[i];
    const gmailThreadId = thread.getId();
    const messages = thread.getMessages();

    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      const messageId = message.getId();

      // 未処理のメッセージの場合
      if (!processedIds.has(messageId)) {
        
        // このGmailスレッドに対応するSlackの親投稿tsがあるか確認
        const parentTs = threadTsMap[gmailThreadId];
        
        // Slackへ送信（parentTsがあればスレッド返信になる）
        const postedTs = sendToSlackAPI(message, parentTs);
        
        if (postedTs) {
          // 親tsがまだ登録されていなければ（新規スレッドの1通目）、取得したtsを親として登録
          if (!parentTs) {
            threadTsMap[gmailThreadId] = postedTs;
          }
          
          // シートに記録（親tsを引き継いで記録する）
          const recordTs = parentTs ? parentTs : postedTs;
          newRows.push([messageId, gmailThreadId, recordTs]);
          processedIds.add(messageId);
        }
      }
    }
  }

  // 4. 新しく処理したデータをスプレッドシートに追記
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
  }
}

// Slack APIを使ってメッセージを送信する関数
function sendToSlackAPI(message, parentTs) {
  const subject = message.getSubject();
  const from = message.getFrom();
  const body = message.getPlainBody();
  const date = message.getDate();
  const messageId = message.getId();

  const maxLength = 2500;
  let textToSend = body;
  if (textToSend.length > maxLength) {
    textToSend = textToSend.substring(0, maxLength) + '\n\n...（文字数制限のため省略）';
  }

  // Slack投稿用のペイロード作成
  const payload = {
    channel: SLACK_CHANNEL_ID,
    attachments: [
      {
        fallback: '新規メール: ' + subject,
        color: parentTs ? '#439FE0' : '#36a64f', // 返信時は少し色を変える
        author_name: from,
        title: subject || '(件名なし)',
        title_link: 'https://mail.google.com/mail/u/0/#inbox/' + messageId,
        text: textToSend,
        footer: 'GMメール転送',
        ts: Math.floor(date.getTime() / 1000)
      }
    ]
  };

  // ★ parentTsが存在する場合（＝返信の場合）はスレッドに繋げる設定を追加
  if (parentTs) {
    payload.thread_ts = parentTs;
    payload.reply_broadcast = true; // 「以下にも投稿する」にチェックを入れる
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + SLACK_BOT_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', options);
    const json = JSON.parse(response.getContentText());
    
    if (json.ok) {
      // 成功したら、投稿されたメッセージのタイムスタンプ(ts)を返す
      return json.ts; 
    } else {
      console.error('Slack API エラー: ' + json.error);
      return null;
    }
  } catch (e) {
    console.error('Slackへの通信に失敗しました: ' + e.message);
    return null;
  }
}

// ==========================================
// スクリプトプロパティ初期設定用関数
// ==========================================
function setupScriptProperty() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'SLACK_BOT_TOKEN': 'hoge',
    'SLACK_CHANNEL_ID': 'hoge',
    'SPREADSHEET_ID': 'hoge',
    'SHEET_NAME': 'hoge'
  });
  console.log('スクリプトプロパティに必要なキーを作成し、ダミー値 "hoge" を設定しました。プロジェクトの設定から正しい値に変更してください。');
}

// ==========================================
// スプレッドシート自動作成＆プロパティ設定用関数
// ==========================================
function setupSpreadsheet() {
  try {
    const ssName = "Slack転送_メール管理リスト";
    const ss = SpreadsheetApp.create(ssName);
    const ssId = ss.getId();
    
    const sheetName = "MailLists";
    ss.getSheets()[0].setName(sheetName);

    try {
      const scriptId = ScriptApp.getScriptId();
      const scriptFile = DriveApp.getFileById(scriptId);
      const parents = scriptFile.getParents();
      
      if (parents.hasNext()) {
        const parentFolder = parents.next();
        const ssFile = DriveApp.getFileById(ssId);
        ssFile.moveTo(parentFolder);
        console.log(`スプレッドシートをスクリプトと同じフォルダ「${parentFolder.getName()}」に移動しました。`);
      }
    } catch (e) {
      console.log("フォルダの移動に失敗したため、マイドライブ直下に作成されました。");
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperties({
      'SPREADSHEET_ID': ssId,
      'SHEET_NAME': sheetName
    });

    console.log(`スプレッドシートのセットアップが完了しました！\nシート名: ${sheetName}\nスプレッドシートURL: ${ss.getUrl()}`);

  } catch (error) {
    console.error("スプレッドシートの作成処理中にエラーが発生しました: " + error.message);
  }
}

// ==========================================
// 自動実行（トリガー）設定用関数
// ==========================================
// この関数を選択して実行すると、main関数を「5分おき」に自動実行する設定を行います。
function setupTrigger() {
  const functionName = 'main';
  
  // 既存の同じトリガーがあれば削除（重複登録を防ぐため）
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 新しく5分おきのトリガーを作成
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(5)
    .create();

  console.log(`関数「${functionName}」を5分おきに自動実行するトリガーを設定しました。`);
}