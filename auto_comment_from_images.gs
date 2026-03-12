/**
 * 【焼肉きんぐ甲府飯田】57期臨店報告書
 * スプレッドシートの問題写真からAIコメントを自動生成するGoogle Apps Script
 *
 * 【列構成】
 *   G: NO / H: 分類 / I: 項目
 *   J: 今回（チェックボックス） / K: 前回
 *   L: 問題写真1 / M: 問題写真2
 *   N: 改善内容 / O: コメント（自動生成先）
 *
 * 【使い方】
 * 1. スプレッドシートの「拡張機能」→「Apps Script」を開く
 * 2. このコードを貼り付けて保存
 * 3. スプレッドシートをリロード
 * 4. メニュー「画像コメント自動生成」→「APIキーを設定」からAPIキーを登録
 * 5. 対象シート（【営業前インスペ】）を開いた状態で実行
 */

// ============================================================
// 設定
// ============================================================
const CONFIG = {
  // 問題写真1の列（L列 = 12）
  IMAGE_COLUMN_1: 12,

  // 問題写真2の列（M列 = 13）
  IMAGE_COLUMN_2: 13,

  // コメントを書き込む列（O列 = 15）
  COMMENT_COLUMN: 15,

  // 今回チェック欄の列（J列 = 10）チェックボックス
  // 未チェック（FALSE） = 指摘あり → 処理対象
  CHECK_COLUMN: 10,

  // 分類の列（H列 = 8）
  CATEGORY_COLUMN: 8,

  // 項目の列（I列 = 9）
  ITEM_COLUMN: 9,

  // 改善内容の列（N列 = 14）
  IMPROVEMENT_COLUMN: 14,

  // データの開始行（ヘッダー2行目の次）
  DATA_START_ROW: 3,

  // ヘッダー行
  HEADER_ROW: 2,

  // 対象シート名のリスト
  TARGET_SHEETS: ["【営業前インスペ】"],

  // 既にコメントがある場合に上書きするか
  OVERWRITE_EXISTING: false,

  // Claude API モデル
  MODEL: "claude-sonnet-4-20250514",

  // AIへのシステムプロンプト
  SYSTEM_PROMPT: `あなたは飲食店（焼肉店）のインスペクション（臨店検査）の専門家です。
提供された問題写真を分析し、指摘事項に対する簡潔で的確なコメントを日本語で記載してください。

コメントには以下を含めてください：
- 画像から確認できる問題点の具体的な説明
- 改善が必要な箇所の指摘
- 改善のための具体的なアドバイス（該当する場合）

コメントは簡潔に2〜3文程度でまとめてください。
専門用語を避け、店舗スタッフが読んですぐ理解できる平易な表現を使ってください。`,
};

// ============================================================
// メニュー追加
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("画像コメント自動生成")
    .addItem("現在のシートのコメントを生成", "generateCommentsForCurrentSheet")
    .addItem("選択範囲のコメントを生成", "generateCommentsForSelection")
    .addItem("全対象シートのコメントを生成", "generateCommentsForAllSheets")
    .addSeparator()
    .addItem("APIキーを設定", "setApiKey")
    .addToUi();
}

// ============================================================
// APIキー管理
// ============================================================
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "APIキー設定",
    "Anthropic APIキーを入力してください:",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty(
        "ANTHROPIC_API_KEY",
        key
      );
      ui.alert("APIキーを保存しました。");
    }
  }
}

function getApiKey() {
  const key =
    PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!key) {
    throw new Error(
      'APIキーが設定されていません。メニュー「画像コメント自動生成」→「APIキーを設定」から設定してください。'
    );
  }
  return key;
}

// ============================================================
// メイン処理：現在のシートのコメント生成
// ============================================================
function generateCommentsForCurrentSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  if (!CONFIG.TARGET_SHEETS.includes(sheetName)) {
    SpreadsheetApp.getUi().alert(
      '対象シートではありません。\n対象: ' + CONFIG.TARGET_SHEETS.join(", ")
    );
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("データが見つかりません。");
    return;
  }

  processRows(sheet, CONFIG.DATA_START_ROW, lastRow);
}

// ============================================================
// メイン処理：選択範囲のコメント生成
// ============================================================
function generateCommentsForSelection() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = SpreadsheetApp.getActiveRange();

  if (!selection) {
    SpreadsheetApp.getUi().alert("範囲を選択してください。");
    return;
  }

  const startRow = selection.getRow();
  const endRow = startRow + selection.getNumRows() - 1;

  processRows(sheet, startRow, endRow);
}

// ============================================================
// メイン処理：全対象シートのコメント生成
// ============================================================
function generateCommentsForAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalProcessed = 0;
  let totalSkipped = 0;
  let totalErrors = 0;

  for (const sheetName of CONFIG.TARGET_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('シート「' + sheetName + '」が見つかりません。スキップします。');
      continue;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_START_ROW) continue;

    const result = processRows(sheet, CONFIG.DATA_START_ROW, lastRow, true);
    totalProcessed += result.processed;
    totalSkipped += result.skipped;
    totalErrors += result.errors;
  }

  SpreadsheetApp.getUi().alert(
    "全シート処理完了\n\n" +
      "生成: " + totalProcessed + " 件\n" +
      "スキップ: " + totalSkipped + " 件\n" +
      "エラー: " + totalErrors + " 件"
  );
}

// ============================================================
// 行を処理してコメント生成
// ============================================================
function processRows(sheet, startRow, endRow, silent) {
  const apiKey = getApiKey();
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  // シート上の全画像を取得
  const images = sheet.getImages();

  for (let row = startRow; row <= endRow; row++) {
    // J列（今回チェック）を確認：チェック済み（TRUE）ならスキップ
    // 指摘あり = 未チェック（FALSE）の行のみ処理
    if (!isRowFlagged(sheet, row)) {
      skippedCount++;
      continue;
    }

    // O列の既存コメントチェック
    const existingComment = sheet
      .getRange(row, CONFIG.COMMENT_COLUMN)
      .getValue();
    if (existingComment && !CONFIG.OVERWRITE_EXISTING) {
      skippedCount++;
      continue;
    }

    // L列・M列の画像を探す（両方取得を試みる）
    const imageBlobs = findImagesForRow(sheet, images, row);
    if (imageBlobs.length === 0) {
      skippedCount++;
      continue;
    }

    // コンテキスト情報を取得
    const context = getRowContext(sheet, row);

    try {
      // Claude APIで画像を分析してコメント生成
      const comment = analyzeImagesWithClaude(apiKey, imageBlobs, context);

      // O列にコメントを書き込み
      sheet.getRange(row, CONFIG.COMMENT_COLUMN).setValue(comment);
      processedCount++;

      // API レート制限対策
      Utilities.sleep(1000);
    } catch (e) {
      Logger.log(sheet.getName() + " 行" + row + " でエラー: " + e.message);
      errorCount++;
    }
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "【" + sheet.getName() + "】処理完了\n\n" +
        "生成: " + processedCount + " 件\n" +
        "スキップ: " + skippedCount + " 件\n" +
        "エラー: " + errorCount + " 件"
    );
  }

  return { processed: processedCount, skipped: skippedCount, errors: errorCount };
}

// ============================================================
// J列のチェック状態を判定（指摘あり = 処理対象かどうか）
// 未チェック（FALSE / 空）→ true（処理する）
// チェック済み（TRUE）→ false（スキップ）
// ============================================================
function isRowFlagged(sheet, row) {
  const value = sheet.getRange(row, CONFIG.CHECK_COLUMN).getValue();

  // チェックボックス: FALSE = 未チェック = 指摘あり
  if (value === false) return true;

  // 空セルも指摘ありとみなす
  if (value === "" || value === null || value === undefined) return true;

  // TRUE（チェック済み）はスキップ
  return false;
}

// ============================================================
// 行に対応する画像を取得（L列・M列の両方）
// ============================================================
function findImagesForRow(sheet, images, targetRow) {
  const blobs = [];

  // L列（問題写真1）の画像を取得
  const blob1 = findImageInColumn(sheet, images, targetRow, CONFIG.IMAGE_COLUMN_1);
  if (blob1) blobs.push(blob1);

  // M列（問題写真2）の画像を取得
  const blob2 = findImageInColumn(sheet, images, targetRow, CONFIG.IMAGE_COLUMN_2);
  if (blob2) blobs.push(blob2);

  return blobs;
}

// ============================================================
// 指定列の画像を取得
// ============================================================
function findImageInColumn(sheet, images, targetRow, targetCol) {
  // OverGridImage（セルの上に配置された画像）を検索
  for (const image of images) {
    const anchor = image.getAnchorCell();
    if (anchor.getRow() === targetRow && anchor.getColumn() === targetCol) {
      return image.getBlob();
    }
  }

  // =IMAGE() 関数をチェック
  try {
    const cell = sheet.getRange(targetRow, targetCol);
    const formula = cell.getFormula();
    if (formula && formula.toUpperCase().startsWith("=IMAGE(")) {
      const match = formula.match(/=IMAGE\("([^"]+)"/i);
      if (match && match[1]) {
        return UrlFetchApp.fetch(match[1]).getBlob();
      }
    }
  } catch (e) {
    Logger.log("IMAGE関数の取得エラー: " + e.message);
  }

  // セルにURLが文字列として入っている場合
  try {
    const cellValue = sheet.getRange(targetRow, targetCol).getValue().toString();
    if (cellValue && isImageUrl(cellValue)) {
      return UrlFetchApp.fetch(cellValue).getBlob();
    }
  } catch (e) {
    Logger.log("画像URL取得エラー: " + e.message);
  }

  return null;
}

// ============================================================
// URLが画像かどうか判定
// ============================================================
function isImageUrl(url) {
  if (!url) return false;
  var lowerUrl = url.toLowerCase();
  return (
    /\.(png|jpg|jpeg|gif|webp|bmp)/.test(lowerUrl) ||
    lowerUrl.includes("googleusercontent.com") ||
    lowerUrl.includes("drive.google.com")
  );
}

// ============================================================
// 行のコンテキスト情報を取得
// ============================================================
function getRowContext(sheet, row) {
  return {
    sheetName: sheet.getName(),
    category: sheet.getRange(row, CONFIG.CATEGORY_COLUMN).getValue().toString().trim(),
    item: sheet.getRange(row, CONFIG.ITEM_COLUMN).getValue().toString().trim(),
    improvement: sheet.getRange(row, CONFIG.IMPROVEMENT_COLUMN).getValue().toString().trim(),
  };
}

// ============================================================
// Claude API で画像を分析（複数画像対応）
// ============================================================
function analyzeImagesWithClaude(apiKey, imageBlobs, context) {
  // メッセージのcontentを組み立て（画像 + テキスト）
  var content = [];

  // 画像を追加
  for (var i = 0; i < imageBlobs.length; i++) {
    var blob = imageBlobs[i];
    content.push({
      type: "image",
      source: {
        type: "base64",
        media_type: blob.getContentType() || "image/png",
        data: Utilities.base64Encode(blob.getBytes()),
      },
    });
  }

  // プロンプトを組み立て
  var userPrompt =
    "この画像は焼肉店のインスペクション（臨店検査）で撮影された問題写真です。\n" +
    "画像を分析し、指摘コメントを記載してください。\n";

  if (imageBlobs.length > 1) {
    userPrompt += "（問題写真が2枚あります。両方を踏まえてコメントしてください。）\n";
  }

  userPrompt += "\n検査種別: " + context.sheetName;

  if (context.category) {
    userPrompt += "\n分類: " + context.category;
  }
  if (context.item) {
    userPrompt += "\n項目: " + context.item;
  }
  if (context.improvement) {
    userPrompt += "\n改善内容: " + context.improvement;
  }

  content.push({
    type: "text",
    text: userPrompt,
  });

  var payload = {
    model: CONFIG.MODEL,
    max_tokens: 1024,
    system: CONFIG.SYSTEM_PROMPT,
    messages: [
      {
        role: "user",
        content: content,
      },
    ],
  };

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(
    "https://api.anthropic.com/v1/messages",
    options
  );
  var responseCode = response.getResponseCode();
  var responseBody = JSON.parse(response.getContentText());

  if (responseCode !== 200) {
    throw new Error(
      "API Error (" + responseCode + "): " +
        (responseBody.error ? responseBody.error.message : "Unknown error")
    );
  }

  var textContent = responseBody.content.find(function (c) {
    return c.type === "text";
  });
  if (!textContent) {
    throw new Error("APIからテキスト応答がありませんでした。");
  }

  return textContent.text;
}
