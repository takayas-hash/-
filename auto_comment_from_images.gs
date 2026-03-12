/**
 * スプレッドシートの画像からAIコメントを自動生成するGoogle Apps Script
 *
 * 【使い方】
 * 1. スプレッドシートの「拡張機能」→「Apps Script」を開く
 * 2. このコードを貼り付ける
 * 3. CONFIG セクションの設定を自分のシートに合わせて変更する
 * 4. ANTHROPIC_API_KEY にClaude APIキーを設定する（スクリプトプロパティ推奨）
 * 5. メニュー「画像コメント自動生成」→「選択範囲のコメントを生成」を実行
 */

// ============================================================
// 設定（自分のスプレッドシートに合わせて変更してください）
// ============================================================
const CONFIG = {
  // 画像が貼られている列（A=1, B=2, ...）
  IMAGE_COLUMN: 4,

  // コメントを書き込む列
  COMMENT_COLUMN: 5,

  // データの開始行（ヘッダーを除いた最初のデータ行）
  DATA_START_ROW: 2,

  // シート名（空文字の場合はアクティブシートを使用）
  SHEET_NAME: "",

  // 指摘項目名の列（コンテキストとしてAIに渡す）
  ITEM_NAME_COLUMN: 2,

  // 指摘箇所の列（コンテキストとしてAIに渡す）
  LOCATION_COLUMN: 3,

  // 既にコメントがある場合に上書きするか
  OVERWRITE_EXISTING: false,

  // Claude API モデル
  MODEL: "claude-sonnet-4-20250514",

  // AIへのシステムプロンプト
  SYSTEM_PROMPT: `あなたは建築・設備のインスペクション（検査）の専門家です。
提供された画像を分析し、指摘事項に対する簡潔で的確なコメントを日本語で記載してください。

コメントには以下を含めてください：
- 画像から確認できる状況の説明
- 問題点や改善が必要な箇所の指摘（該当する場合）
- 対応状況の評価（是正済み・未是正・経過観察など）

コメントは簡潔に3〜5文程度でまとめてください。`,
};

// ============================================================
// メニュー追加
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("画像コメント自動生成")
    .addItem("選択範囲のコメントを生成", "generateCommentsForSelection")
    .addItem("シート全体のコメントを生成", "generateCommentsForSheet")
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
      "APIキーが設定されていません。メニュー「画像コメント自動生成」→「APIキーを設定」から設定してください。"
    );
  }
  return key;
}

// ============================================================
// メイン処理：選択範囲のコメント生成
// ============================================================
function generateCommentsForSelection() {
  const sheet = getTargetSheet();
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
// メイン処理：シート全体のコメント生成
// ============================================================
function generateCommentsForSheet() {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("データが見つかりません。");
    return;
  }

  processRows(sheet, CONFIG.DATA_START_ROW, lastRow);
}

// ============================================================
// 行を処理してコメント生成
// ============================================================
function processRows(sheet, startRow, endRow) {
  const apiKey = getApiKey();
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  // シート上の全画像を取得
  const images = sheet.getImages();

  for (let row = startRow; row <= endRow; row++) {
    // 既存コメントチェック
    const existingComment = sheet
      .getRange(row, CONFIG.COMMENT_COLUMN)
      .getValue();
    if (existingComment && !CONFIG.OVERWRITE_EXISTING) {
      skippedCount++;
      continue;
    }

    // その行の画像列に配置された画像を探す
    const imageBlob = findImageForRow(sheet, images, row);
    if (!imageBlob) {
      skippedCount++;
      continue;
    }

    // コンテキスト情報を取得
    const context = getRowContext(sheet, row);

    try {
      // Claude APIで画像を分析
      const comment = analyzeImageWithClaude(apiKey, imageBlob, context);

      // コメントをセルに書き込み
      sheet.getRange(row, CONFIG.COMMENT_COLUMN).setValue(comment);
      processedCount++;

      // API レート制限対策
      Utilities.sleep(1000);
    } catch (e) {
      Logger.log("行 " + row + " でエラー: " + e.message);
      errorCount++;
    }
  }

  // 結果を表示
  SpreadsheetApp.getUi().alert(
    "処理完了\n\n" +
      "生成: " + processedCount + " 件\n" +
      "スキップ: " + skippedCount + " 件\n" +
      "エラー: " + errorCount + " 件"
  );
}

// ============================================================
// 行に対応する画像を取得
// ============================================================
function findImageForRow(sheet, images, targetRow) {
  // OverGridImage（セルの上に配置された画像）を検索
  for (const image of images) {
    const anchor = image.getAnchorCell();
    const imageRow = anchor.getRow();
    const imageCol = anchor.getColumn();

    // 画像列にあり、対象行にある画像を検索
    if (imageRow === targetRow && imageCol === CONFIG.IMAGE_COLUMN) {
      return image.getBlob();
    }
  }

  // セル内画像（IMAGE関数やCellImage）をチェック
  const cell = sheet.getRange(targetRow, CONFIG.IMAGE_COLUMN);
  const richText = cell.getRichTextValue();
  if (richText) {
    const linkUrl = richText.getLinkUrl();
    if (linkUrl && isImageUrl(linkUrl)) {
      try {
        const response = UrlFetchApp.fetch(linkUrl);
        return response.getBlob();
      } catch (e) {
        Logger.log("画像URLの取得に失敗: " + linkUrl);
      }
    }
  }

  // セルに画像のURLが文字列として入っている場合
  const cellValue = cell.getValue().toString();
  if (cellValue && isImageUrl(cellValue)) {
    try {
      const response = UrlFetchApp.fetch(cellValue);
      return response.getBlob();
    } catch (e) {
      Logger.log("画像URLの取得に失敗: " + cellValue);
    }
  }

  return null;
}

// ============================================================
// URLが画像かどうか判定
// ============================================================
function isImageUrl(url) {
  if (!url) return false;
  const imageExtensions = [".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"];
  const lowerUrl = url.toLowerCase();
  return (
    imageExtensions.some((ext) => lowerUrl.includes(ext)) ||
    lowerUrl.includes("googleusercontent.com") ||
    lowerUrl.includes("drive.google.com")
  );
}

// ============================================================
// 行のコンテキスト情報を取得
// ============================================================
function getRowContext(sheet, row) {
  const context = {};

  if (CONFIG.ITEM_NAME_COLUMN) {
    context.itemName = sheet
      .getRange(row, CONFIG.ITEM_NAME_COLUMN)
      .getValue()
      .toString();
  }

  if (CONFIG.LOCATION_COLUMN) {
    context.location = sheet
      .getRange(row, CONFIG.LOCATION_COLUMN)
      .getValue()
      .toString();
  }

  return context;
}

// ============================================================
// Claude API で画像を分析
// ============================================================
function analyzeImageWithClaude(apiKey, imageBlob, context) {
  const base64Image = Utilities.base64Encode(imageBlob.getBytes());
  const mimeType = imageBlob.getContentType() || "image/png";

  // ユーザープロンプトを組み立て
  let userPrompt = "この画像を分析して、インスペクション（検査）のコメントを記載してください。";

  if (context.itemName) {
    userPrompt += "\n指摘項目: " + context.itemName;
  }
  if (context.location) {
    userPrompt += "\n指摘箇所: " + context.location;
  }

  const payload = {
    model: CONFIG.MODEL,
    max_tokens: 1024,
    system: CONFIG.SYSTEM_PROMPT,
    messages: [
      {
        role: "user",
        content: [
          {
            type: "image",
            source: {
              type: "base64",
              media_type: mimeType,
              data: base64Image,
            },
          },
          {
            type: "text",
            text: userPrompt,
          },
        ],
      },
    ],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(
    "https://api.anthropic.com/v1/messages",
    options
  );
  const responseCode = response.getResponseCode();
  const responseBody = JSON.parse(response.getContentText());

  if (responseCode !== 200) {
    throw new Error(
      "API Error (" + responseCode + "): " +
      (responseBody.error ? responseBody.error.message : "Unknown error")
    );
  }

  // レスポンスからテキストを抽出
  const textContent = responseBody.content.find((c) => c.type === "text");
  if (!textContent) {
    throw new Error("APIからテキスト応答がありませんでした。");
  }

  return textContent.text;
}

// ============================================================
// 対象シートを取得
// ============================================================
function getTargetSheet() {
  if (CONFIG.SHEET_NAME) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEET_NAME
    );
    if (!sheet) {
      throw new Error("シート「" + CONFIG.SHEET_NAME + "」が見つかりません。");
    }
    return sheet;
  }
  return SpreadsheetApp.getActiveSheet();
}

// ============================================================
// Google Drive内の画像をBlobとして取得（IMAGE関数で使用されるDriveのURLに対応）
// ============================================================
function getImageFromDrive(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (e) {
    Logger.log("Drive画像の取得に失敗: " + e.message);
    return null;
  }
}

// ============================================================
// CellImage（セル内に挿入された画像）を取得する試み
// Note: Google Apps Scriptでは CellImage の直接取得に制限があるため、
// OverGridImage（セル上に配置された画像）の利用を推奨します。
// ============================================================
function findCellImageForRow(sheet, row) {
  try {
    const cell = sheet.getRange(row, CONFIG.IMAGE_COLUMN);
    // getFormula() で =IMAGE() 関数を検出
    const formula = cell.getFormula();
    if (formula && formula.toUpperCase().startsWith("=IMAGE(")) {
      // =IMAGE("URL") からURLを抽出
      const match = formula.match(/=IMAGE\("([^"]+)"/i);
      if (match && match[1]) {
        const imageUrl = match[1];
        const response = UrlFetchApp.fetch(imageUrl);
        return response.getBlob();
      }
    }
  } catch (e) {
    Logger.log("CellImage取得エラー: " + e.message);
  }
  return null;
}
