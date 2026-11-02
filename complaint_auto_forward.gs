// =================================================================
// 苦情・問い合わせ自動転送 GAS(エリアマネージャー配布用テンプレート)
// ------------------------------------------------------------
// このスプレッドシートをコピーし、「シート1」(対応表)を自分のエリアの
// 店舗情報に置き換えるだけで使えます。コード側の変更は不要です。
//
// 【初回セットアップ手順】
// 1. このスプレッドシートをコピー(ファイル > コピーを作成)
// 2. 「シート1」(対応表)の中身を自分のエリアの店舗に置き換える
//    A:店舗コード / B:店舗名 / C:店長名 / D:店長メールアドレス / E:Google Chat Webhook URL
//    ※タブ名は「シート1」のままにすること(コードで名前指定しているため変更不可)
// 3. 拡張機能 > Apps Script を開き、関数選択で installTrigger を選んで実行
//    → 初回はGmail・カレンダー(Tasks)の権限承認画面が出るので許可する
// 4. 以降は15分おきに autoForwardBasicRequest が自動実行される
//
// 【前提条件】
// info@monogatari.co.jp 宛のメールが、このスクリプトを実行する
// Googleアカウント自身のGmailにも届いている必要があります
// (Googleグループ配信・転送設定などで未着信の場合は動作しません)。
//
// 【注意】
// 検索は直近3日以内のメールのみが対象(newer_than:3d)。トリガーが
// 3日以上停止した場合(権限失効・障害等)、それより古いメールは
// 自動では拾われないため手動確認が必要。
//
// 【対象メールの件名フォーマット】
// 例: 【苦情】焼肉きんぐ 戸田店 要対応
// 以下がすべて件名に含まれている必要がある:
//   ・カテゴリタグ: 【苦情】【ご質問】【ご要望】【ご意見】のいずれか
//   ・「焼肉きんぐ」の表記(他業態の同名店舗との誤マッチ防止のため必須)
//   ・対応表(B列)に登録した店舗名(「店」の有無は問わない、部分一致)
//   ・「要対応」または「対応不要」
// ==================================================================

function autoForwardBasicRequest() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート1');
  if (!sheet) {
    Logger.log("対応表シートが見つかりません。タブ名が「シート1」のままか確認してください。");
    return;
  }
  var data = sheet.getDataRange().getValues();

  // 返信文の送付先・完了報告先 = このスクリプトを実行しているエリアマネージャー本人。
  // (テンプレートをコピーした各マネージャーのアドレスに自動で切り替わる)
  var reportTo = Session.getEffectiveUser().getEmail();

  var processedLabel = GmailApp.getUserLabelByName('自動転送済み')
    || GmailApp.createLabel('自動転送済み');

  // 「未読かどうか」ではなく「ラベルが付いているかどうか」で処理済みを判定する。
  // is:unread だと、誰かが受信箱を開いただけで未処理のまま検索対象から外れてしまうため。
  // newer_than:3d は、ラベル未整備の古いメールを初回実行時に大量取得しないための安全弁。
  var searchQuery = 'from:info@monogatari.co.jp (subject:"【苦情】" OR subject:"【ご質問】" OR subject:"【ご要望】" OR subject:"【ご意見】") (subject:"要対応" OR subject:"対応不要") newer_than:3d -label:自動転送済み -label:"対象外（他エリア・店舗不明）"';
  var threads = GmailApp.search(searchQuery);

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    var message = messages[messages.length - 1];
    var subject = message.getSubject();
    var body = message.getBody()
      .replace(/\r\n?/g, '\n')
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
    // 本部向けの定型文(「関係者各位」〜「本部のアカウントでお客様へ送信します」)を除き、
    // 「━━━ お問い合わせ内容」以降だけを店長へ転送する。
    body = maskCustomerEmail_(extractInquiryPart_(body));

    var categoryMatch = subject.match(/【(苦情|ご質問|ご要望|ご意見)】/);
    var category = categoryMatch ? categoryMatch[1] : "お問い合わせ";

    var isActionRequired = subject.indexOf("要対応") !== -1;

    var receiveDate = message.getDate();
    var deadlineDate = new Date(receiveDate.getTime());
    deadlineDate.setDate(deadlineDate.getDate() + 2);
    var deadlineStr = (deadlineDate.getMonth() + 1) + "月" + deadlineDate.getDate() + "日";

    var matched = false;

    for (var j = 1; j < data.length; j++) {
      var storeName    = data[j][1];
      var managerName  = data[j][2];
      var managerEmail = data[j][3];
      var webhookUrl   = data[j][4]; // E列

      // 店舗名だけの一致だと、焼肉きんぐ以外の業態に同名店舗がある場合に誤爆するため、
      // 「焼肉きんぐ」表記も件名に含まれていることを合わせて確認する。
      var storeIndex = subject.indexOf(storeName);
      var isKingBrand = subject.indexOf("焼肉きんぐ") !== -1;

      if (storeIndex !== -1 && isKingBrand) {
        matched = true;

        // お客様の苦情本文に店長本人の氏名が名指しで含まれる場合がある(暴言を伴うことも
        // あるため)。本人へそのまま転送すると配慮に欠けるので、氏名部分のみ伏字にする。
        // 元メール(infoshop@受信箱)自体は改変しないので、生の内容は必要に応じてそちらで確認できる。
        var sanitizedBody = (managerName && body.indexOf(managerName) !== -1)
          ? body.split(managerName).join("（氏名は伏せています）")
          : body;

        // --- 1. メール送信 ---
        try {
          var forwardBody;
          var mailSubject;

          if (isActionRequired) {
            forwardBody = managerName + "店長\n\n"
              + "お疲れ様です。\n\n"
              + "要対応の" + category + "をいただきました。\n"
              + deadlineStr + "の営業終了までに、以下の手順でご対応をお願いいたします。\n\n"
              + "【対応手順】\n"
              + "1. まずはお客様へお電話でご対応ください。\n"
              + "2. お電話が繋がらない場合は、お客様への返信文を作成し、\n"
              + "   " + reportTo + " 宛にお送りください。\n"
              + "   こちらで添削のうえ、ブロック長へ送付します。\n"
              + "※店舗からお客様へ直接メールを送ることはできません。\n"
              + "※お電話で対応が完了した場合も、対応内容を " + reportTo + " までご報告ください。\n\n"
              + "-".repeat(30) + "\n\n"
              + sanitizedBody;
            mailSubject = "【" + deadlineStr + "まで要対応】" + subject;
          } else {
            forwardBody = managerName + "店長\n\n"
              + "お疲れ様です。\n\n"
              + "「対応不要」の" + category + "をいただきましたので、参考までに共有します。\n\n"
              + "特段のご対応は不要ですが、内容をご確認ください。\n\n"
              + "-".repeat(30) + "\n\n"
              + sanitizedBody;
            mailSubject = "【対応不要・参考共有】" + subject;
          }

          GmailApp.sendEmail(managerEmail, mailSubject, forwardBody);
        } catch(e) {
          Logger.log("メール送信エラー [" + storeName + "]: " + e.message);
        }

        // --- 2. Google Chat通知（要対応のみ）---
        if (isActionRequired && webhookUrl && webhookUrl.startsWith("http")) {
          var chatText = "【自動通知】\n"
            + managerName + "店長、お疲れ様です。\n\n"
            + "要対応の「" + category + "」を受信したため、店長のメール宛に転送しました。\n"
            + "まずはお客様へお電話でご対応ください。\n"
            + "繋がらない場合は返信文を作成し、" + reportTo + " 宛にお送りください。\n\n"
            + "▼件名\n" + subject + "\n"
            + "▼期限\n" + deadlineStr + "の営業終了まで";
          try {
            UrlFetchApp.fetch(webhookUrl, {
              method: "post",
              contentType: "application/json",
              payload: JSON.stringify({ "text": chatText })
            });
          } catch(e) {
            Logger.log("チャット送信エラー [" + storeName + "]: " + e.message);
          }
        }

        // --- 3. ToDoリスト追加（要対応のみ）---
        if (isActionRequired) {
          try {
            Tasks.Tasks.insert({
              title: "対応状況確認：" + storeName + "店 (" + managerName + "店長)",
              notes: "件名: 【" + deadlineStr + "まで要対応】" + subject + "\n※店舗からの対応完了報告を待つ",
              due: deadlineDate.toISOString()
            }, '@default');
          } catch(e) {
            Logger.log("Tasks APIエラー: " + e.message);
          }
        }

        // --- 4. 処理ログをシートに記録 ---
        logToSheet(receiveDate, storeName, managerName, category, subject, isActionRequired ? deadlineStr : "対応不要");

        break;
      }
    }

    if (matched) {
      threads[i].markRead();
      threads[i].addLabel(processedLabel);
    } else {
      // 自分のエリアに該当店舗がない(他エリア宛、または対応表の店舗名相違)。
      // 全マネージャー共通の受信箱を使う運用のため、他エリア宛は毎回ノイズになる。
      // 既読化してラベルのみ付与し、通知は出さない。
      var otherAreaLabel = GmailApp.getUserLabelByName('対象外（他エリア・店舗不明）')
        || GmailApp.createLabel('対象外（他エリア・店舗不明）');
      threads[i].addLabel(otherAreaLabel);
      threads[i].markRead();
      Logger.log("店舗名マッチなし: " + subject);
    }
  }
}

// ── 本文から「━━━ お問い合わせ内容」以降だけを取り出す ──
// 区切り線が見つからない形式のメールは、取りこぼし防止のため本文全体を返す。
function extractInquiryPart_(body) {
  var idx = body.indexOf('━━━');
  return idx === -1 ? body : body.substring(idx).trim();
}

// ── お客様のメールアドレスを伏せる ──
// 店舗からお客様へ直接メールすることはないため、店長への転送では不要。
// (電話対応に使う電話番号はそのまま残す)
function maskCustomerEmail_(body) {
  return body.replace(/(●メールアドレス[：:]).*/g, '$1（店舗からの直接返信は不可のため省略）');
}

// ── 処理ログ記録 ──
function logToSheet(date, store, manager, category, subject, deadline) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('転送ログ') || ss.insertSheet('転送ログ');
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['受信日時', '店舗名', '店長名', 'カテゴリ', '件名', '期限', '転送日時']);
  }
  logSheet.appendRow([date, store, manager, category, subject, deadline, new Date()]);
}

// ── 初回セットアップ用: トリガー登録(15分おき) ──
// コピー後、Apps Scriptエディタでこの関数を一度だけ実行する。
// 既に登録済みのトリガーがあれば削除してから登録するため、二重実行しても安全。
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'autoForwardBasicRequest'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoForwardBasicRequest').timeBased().everyMinutes(15).create();
}