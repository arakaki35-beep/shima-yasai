/**
 * メイン関数：Excelを取得してスプシに保存
 */
function updateVegetableData() {
  try {
    // 1. 最新のExcel URLを取得
    const excelUrl = getLatestExcelUrl();
    
    // 2. Excelをパース
    const data = fetchAndParseExcel(excelUrl);
    
    // 3. スプシに保存
    saveToSheet(data);
    
    Logger.log('✅ 処理完了: ' + data.length + '件のデータを処理しました');
    
  } catch (error) {
    Logger.log('❌ エラー発生: ' + error.message);
    // 必要に応じてメール通知を追加
    // MailApp.sendEmail('your-email@example.com', 'GASエラー通知', error.message);
  }
}

/**
 * 沖縄県のページから最新のyasai Excelファイル URLを取得
 */
function getLatestExcelUrl() {
  const pageUrl = "https://www.pref.okinawa.lg.jp/shigoto/shinkooroshi/1011552/1024140/1024142.html";
  const html = UrlFetchApp.fetch(pageUrl).getContentText();
  
  // "_res"以降のパスを直接抽出
  const match = html.match(/href="[^"]*(_res[^"]*yasai[^"]*\.xlsx?)"/);
  
  if (match && match[1]) {
    const excelPath = match[1];
    const excelUrl = 'https://www.pref.okinawa.lg.jp/' + excelPath;
    return excelUrl;
  }
  
  throw new Error('yasai Excelファイルが見つかりませんでした');
}

/**
 * ExcelファイルをダウンロードしてパースALL
 */
function fetchAndParseExcel(excelUrl) {
  // Excelダウンロード
  const response = UrlFetchApp.fetch(excelUrl);
  const blob = response.getBlob();
  
  // Sheetsに変換
  const fileMetadata = {
    title: 'temp_excel_' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  
  const file = Drive.Files.create(fileMetadata, blob);
  
  const spreadsheet = SpreadsheetApp.openById(file.id);
  const allData = [];
  
  // 全シート（月曜日〜土曜日）をループ
  const sheets = spreadsheet.getSheets();
  sheets.forEach(sheet => {
    // 非表示シートはスキップ
    if (sheet.isSheetHidden()) {
      return;
    }
    
    // シート名が「月曜日」「火曜日」...「土曜日」のいずれかかチェック
    const sheetName = sheet.getName();
    if (/^[月火水木金土]曜日$/.test(sheetName)) {
      const data = parseSheet(sheet);
      allData.push(...data);
    }
  });
  
  // 一時ファイル削除
  Drive.Files.remove(file.id);
  
  return allData;
}

/**
 * 1シートをパース
 */
function parseSheet(sheet) {
  const result = [];
  
  // E4セルから日付取得
  const dateCell = sheet.getRange('E4').getValue();
  if (!dateCell) {
    return result;
  }
  
  const date = convertWarekiToDate(dateCell.toString());
  const dateStr = Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
  
  // 7行目から最終行まで取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 7) return result;
  
  // A列（通し番号）、B列（品目名）、G列（平均価格）を取得
  const rangeA = sheet.getRange(7, 1, lastRow - 6, 1).getValues();
  const rangeB = sheet.getRange(7, 2, lastRow - 6, 1).getValues();
  const rangeG = sheet.getRange(7, 7, lastRow - 6, 1).getValues();
  
  for (let i = 0; i < rangeA.length; i++) {
    const rowNumber = rangeA[i][0];
    const vegetableName = rangeB[i][0];
    const avgPrice = rangeG[i][0];
    
    // A列が空なら終了（終端判定）
    if (!rowNumber || rowNumber === '') break;
    
    // B列とG列がある場合のみ保存
    if (vegetableName && avgPrice) {
      result.push([dateStr, vegetableName, avgPrice]);
    }
  }
  
  return result;
}

/**
 * 和暦→西暦変換
 */
function convertWarekiToDate(warekiStr) {
  const match = warekiStr.match(/令和(\d+)年(\d+)月(\d+)日/);
  if (match) {
    const year = 2018 + parseInt(match[1]); // 令和元年=2019
    const month = parseInt(match[2]) - 1;
    const day = parseInt(match[3]);
    return new Date(year, month, day);
  }
  throw new Error('日付変換エラー: ' + warekiStr);
}

/**
 * データをスプシに保存
 */
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('データ');
  
  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet('データ');
    sheet.appendRow(['日付', '品目名', '平均価格']);
  }
  
  // 重複チェック：同じ日付のデータは追加しない
  const existingData = sheet.getDataRange().getValues();
  const existingDates = new Set(existingData.map(row => row[0]));
  
  const newData = data.filter(row => !existingDates.has(row[0]));
  
  if (newData.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newData.length, 3).setValues(newData);
    Logger.log('新規追加: ' + newData.length + '件');
  } else {
    Logger.log('新規データなし（重複のためスキップ）');
  }
}

/**
 * Web公開用：JSON形式でデータを返す
 * デプロイ → 新しいデプロイ → ウェブアプリ で公開URL取得
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('データ');
  
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'データシートが見つかりません'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const path = e.parameter.path || 'vegetables-history';
  
  if (path === 'vegetables-list-with-prices') {
    return getVegetableListWithPrices(sheet);
  } else if (path === 'vegetables-history') {
    return getVegetablesHistory(sheet);
  } else if (path === 'vegetables') {
    return getLatestVegetables(sheet);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'error',
      message: '不明なエンドポイント'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 野菜一覧と最新価格を返す
 */
function getVegetableListWithPrices(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 各野菜の最新価格を取得
  const vegetablePrices = {};
  const vegetableList = new Set();
  
  for (let i = data.length - 1; i >= 1; i--) {
    const vegetableName = data[i][1];
    const price = data[i][2];
    const date = data[i][0];
    
    if (!vegetablePrices[vegetableName]) {
      vegetablePrices[vegetableName] = {
        price: price,
        date: Utilities.formatDate(new Date(date), 'JST', 'yyyy-MM-dd')
      };
      vegetableList.add(vegetableName);
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      data: {
        list: Array.from(vegetableList).sort(),
        prices: vegetablePrices
      }
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 全野菜の履歴データを返す（過去30日分）
 */
function getVegetablesHistory(sheet) {
  const data = sheet.getDataRange().getValues();
  
  // 過去30日間の日付を計算
  const today = new Date();
  const thirtyDaysAgo = new Date(today.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  // 野菜ごとにデータを整理
  const vegetableData = {};
  const datesSet = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const dateStr = Utilities.formatDate(new Date(data[i][0]), 'JST', 'yyyy-MM-dd');
    const vegetableName = data[i][1];
    const price = data[i][2];
    const date = new Date(data[i][0]);
    
    // 過去30日以内のデータのみ
    if (date >= thirtyDaysAgo) {
      if (!vegetableData[vegetableName]) {
        vegetableData[vegetableName] = [];
      }
      
      vegetableData[vegetableName].push({
        date: dateStr,
        price: price
      });
      
      datesSet.add(dateStr);
    }
  }
  
  // 日付をソート
  const dates = Array.from(datesSet).sort();
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      data: vegetableData,
      dates: dates
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 最新の野菜価格を返す（フォールバック用）
 */
function getLatestVegetables(sheet) {
  const data = sheet.getDataRange().getValues();
  
  // 最新の日付を取得
  const latestDate = data[data.length - 1][0];
  
  // 最新日付のデータのみ抽出
  const latestData = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === latestDate.toString()) {
      latestData.push({
        name: data[i][1],
        price: data[i][2]
      });
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      data: latestData
    }))
    .setMimeType(ContentService.MimeType.JSON);
}