function createMonthlyForecasts() {
  // === 1. 設定エリア ===
  const saveFolderId = "1MiKdID8SJ7A1YRKdEzYubRbKSE83Hpfu"; // 完成した各店舗のファイルを保存するフォルダ
  const templateFolderId = "1mex-tUM7Plxa3WfJ13D_zzP9FPXi14z1"; // 「原本」が入っているフォルダ
  const sourceFolderId = "1bi6o7AMqzHCh0e15dxKTf2FGxYbZSWPc"; // 「貼り付け用」が入っているフォルダ

  // 貼り付け用ファイルを探すときの名前（先頭一致）。日付・時刻が末尾に付与されても見つかるように前方一致で検索する
  const sourceFileNamePrefix = "【見通しFM】";
  // 原本ファイルを探すときのキーワード（部分一致）。「損益管理」「損益帳票」などの呼称変更があっても見つかるようにする
  const templateFileNameKeyword = "見通しフォーマット【原本】";

  const sourceStartRow = 5; // コピーするデータの開始行
  const sourceStartCol = 1; // コピーするデータの開始列（A列）
  const targetRow = 2; // 貼り付けを開始する行番号
  const targetCol = 2; // 貼り付けを開始する列番号
  const targetSheetName = "損益帳票_貼り付け用"; // データ貼り付け先のシート名

  // 数式を値に変換する対象のシート名（E1, G1の書き込み先も同じシート）
  const valueSheetName = "取込・更新用";

  // 値に変換する列
  const targetCols = ["T", "AD", "AN", "AX", "BH", "BR", "CB", "CL", "CV", "DF", "DP", "DZ", "EJ", "ET", "FD", "FN", "FX", "GH", "GR"];

  // =========================================================

  // 2. フォルダの取得
  const saveFolder = DriveApp.getFolderById(saveFolderId);
  const templateFolder = DriveApp.getFolderById(templateFolderId);
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);

  // 3. 貼り付け用フォルダから【貼り付け用ファイル】を探す（前方一致。末尾に日時が付与されていても対象にする）
  const sourceFile = findFileByPrefix_(sourceFolder, sourceFileNamePrefix);
  if (!sourceFile) {
    console.log("処理対象の「" + sourceFileNamePrefix + "」から始まるファイルがありませんでした。処理をスキップします。");
    return;
  }
  const sourceSpreadsheet = SpreadsheetApp.open(sourceFile);

  // 4. 原本フォルダから【原本ファイル】を探す（部分一致）
  const templateFile = findFileByKeyword_(templateFolder, templateFileNameKeyword);
  if (!templateFile) {
    console.error("エラー: 原本フォルダの中に「" + templateFileNameKeyword + "」を含むファイルが見つかりません。");
    return;
  }

  // 5. 店舗ごとにループ処理
  const sheets = sourceSpreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const storeName = sheet.getName(); // タブ名を取得

    // ファイル名を作成
    const newFileName = "損益管理_見通しフォーマット【" + storeName + "】";

    // 原本をコピーして保存用フォルダに新しいファイルを作成
    const newFile = templateFile.makeCopy(newFileName, saveFolder);
    const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());

    // 6. データのコピー＆ペースト処理
    // コピー範囲は決め打ちにせず、そのシートの実データ量から自動計算する
    // （損益項目の行数が変わってもコード修正が不要になる）
    const sourceLastRow = sheet.getLastRow();
    const sourceLastCol = sheet.getLastColumn();
    const sourceData = sheet.getRange(
      sourceStartRow,
      sourceStartCol,
      sourceLastRow - sourceStartRow + 1,
      sourceLastCol - sourceStartCol + 1
    ).getValues();

    const targetSheet = newSpreadsheet.getSheetByName(targetSheetName);

    if (!targetSheet) {
      console.error("エラー: 原本ファイルの中に「" + targetSheetName + "」という名前のシートが見つかりません。");
      return;
    }

    if (sourceData.length > 0 && sourceData[0].length > 0) {
      targetSheet.getRange(
        targetRow,
        targetCol,
        sourceData.length,
        sourceData[0].length
      ).setValues(sourceData);
    }

    SpreadsheetApp.flush();

    // 7. 取込・更新用シートの処理（セルの入力と値への変換）
    const valueSheet = newSpreadsheet.getSheetByName(valueSheetName);

    if (valueSheet) {
      const match = storeName.match(/《(.*?)》(.*)/);
      if (match) {
        const storeCode = match[1]; // 店舗コード
        const storeFullName = match[2]; // 店舗名

        const rangeE1 = valueSheet.getRange("E1");
        const rangeG1 = valueSheet.getRange("G1");

        // ★追加：エラー回避のため、E1とG1の「入力規則」を解除（無視）する
        rangeE1.clearDataValidations();
        rangeG1.clearDataValidations();

        // その上で値をセットする
        rangeE1.setValue(storeCode);
        rangeG1.setValue(storeFullName);
      }

      const lastRow = valueSheet.getMaxRows();

      targetCols.forEach(col => {
        const range = valueSheet.getRange(col + "1:" + col + lastRow);
        range.setValues(range.getValues());
      });
    } else {
      console.warn("警告: 対象のシート「" + valueSheetName + "」が見つかりませんでした。");
    }

    SpreadsheetApp.flush();
  }

  // 8. 処理が終わった「貼り付け用ファイル」をゴミ箱に移動する
  sourceFile.setTrashed(true);

  console.log("✨ 全店舗の見通しファイルの作成・値化・店舗情報の入力が完了し、貼り付け用データを破棄しました。");
}

// フォルダ内から、名前が指定の文字列で始まるファイルを1件探す
function findFileByPrefix_(folder, prefix) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().indexOf(prefix) === 0) {
      return file;
    }
  }
  return null;
}

// フォルダ内から、名前に指定の文字列を含むファイルを1件探す
function findFileByKeyword_(folder, keyword) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().indexOf(keyword) !== -1) {
      return file;
    }
  }
  return null;
}
