// 編集履歴
// 2025/3/26: incrementGrades関数を追加（入野）
// 2025/3/3: onFormSubmit関数、それに付随するヘルパー関数群を実装（入野）

// 定数定義
const SHEET_NAMES = {
  MALE: "男子",
  FEMALE: "女子", 
  FORM_RESPONSES: "Form Responses"
};

const COLUMN_INDICES = {
  NAME: 0,        // A列: 氏名
  GRADE: 1,       // B列: 学年
  RANK: 2,        // C列: ランク
  TIMESTAMP: 0,   // Form ResponsesのA列: タイムスタンプ
  APPLICANT: 1,   // Form ResponsesのB列: 申込者
  OPPONENT: 2,    // Form ResponsesのC列: 対戦相手
  RESULT: 3       // Form ResponsesのD列: 勝敗
};

const RESULT_VALUES = {
  WIN: "勝利"
};

const LOG_LEVEL = {
  ERROR: "ERROR",
  WARNING: "WARNING", 
  INFO: "INFO",
  DEBUG: "DEBUG"
};

// ログ出力関数
function logMessage(level, message) {
  const timestamp = new Date().toISOString();
  Logger.log(`[${timestamp}] [${level}] ${message}`);
}

// フォーム送信時に実行される関数
function onFormSubmit(e) {
  try {
    logMessage(LOG_LEVEL.INFO, "ランク戦結果処理を開始します");
    
    // スプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formResponsesSheet = ss.getSheetByName(SHEET_NAMES.FORM_RESPONSES);
    const maleSheet = ss.getSheetByName(SHEET_NAMES.MALE);
    const femaleSheet = ss.getSheetByName(SHEET_NAMES.FEMALE);
    
    // 最新のフォーム回答を取得
    const lastRow = formResponsesSheet.getLastRow();
    const formData = formResponsesSheet.getRange(lastRow, 1, 1, 4).getValues()[0];
    
    // フォームデータを抽出
    const timestamp = formData[COLUMN_INDICES.TIMESTAMP];
    const applicantName = formData[COLUMN_INDICES.APPLICANT];
    const opponentName = formData[COLUMN_INDICES.OPPONENT];
    const matchResult = formData[COLUMN_INDICES.RESULT];
    
    logMessage(LOG_LEVEL.INFO, `申込者: ${applicantName}, 対戦相手: ${opponentName}, 結果: ${matchResult}`);
    
    // 勝利の場合のみ処理を実行
    if (matchResult === RESULT_VALUES.WIN) {
      // 男子シートと女子シートの両方で検索
      const maleApplicantRow = findPersonRow(maleSheet, applicantName);
      const maleOpponentRow = findPersonRow(maleSheet, opponentName);
      const femaleApplicantRow = findPersonRow(femaleSheet, applicantName);
      const femaleOpponentRow = findPersonRow(femaleSheet, opponentName);
      
      // 男子シートで両方見つかった場合
      if (maleApplicantRow > 0 && maleOpponentRow > 0) {
        logMessage(LOG_LEVEL.INFO, "男子シートで両選手を発見しました");
        updateRankings(maleSheet, maleApplicantRow, maleOpponentRow);
      } 
      // 女子シートで両方見つかった場合
      else if (femaleApplicantRow > 0 && femaleOpponentRow > 0) {
        logMessage(LOG_LEVEL.INFO, "女子シートで両選手を発見しました");
        updateRankings(femaleSheet, femaleApplicantRow, femaleOpponentRow);
      } 
      // 見つからない場合
      else {
        logMessage(LOG_LEVEL.WARNING, "申込者または対戦相手が見つかりませんでした");
      }
    } else {
      logMessage(LOG_LEVEL.INFO, "勝利以外の結果のため、ランキング更新は行いません");
    }
    
    logMessage(LOG_LEVEL.INFO, "ランク戦結果処理が完了しました");
  } catch (error) {
    logMessage(LOG_LEVEL.ERROR, `処理中にエラーが発生しました: ${error.message}`);
    logMessage(LOG_LEVEL.ERROR, `スタックトレース: ${error.stack}`);
  }
}

// シート内で指定された名前の行を検索する関数
function findPersonRow(sheet, name) {
  try {
    // シートの全データを一度に取得（効率化）、ヘッダー行を除く
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues(); // getValuesからgetDisplayValuesに変更
    
    // 名前が一致する行を検索
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].trim() === name.trim()) { // 前後の空白を削除して比較
        return i + 2; // ヘッダー行を考慮して+2
      }
    }
    
    // 見つからなかった場合
    logMessage(LOG_LEVEL.WARNING, `"${name}"という名前はシート内に見つかりませんでした`);
    return -1;
  } catch (error) {
    logMessage(LOG_LEVEL.ERROR, `名前検索中にエラーが発生しました: ${error.message}`);
    return -1;
  }
}

// ランキングを更新する関数
function updateRankings(sheet, applicantRow, opponentRow) {
  try {
    // 現在の順位を確認
    if (applicantRow < opponentRow) {
      logMessage(LOG_LEVEL.INFO, "申込者は既に対戦相手より上位にいるため、順位変更は不要です");
      return;
    }
    
    // 対象範囲のデータを一度に取得
    const startRow = Math.min(opponentRow, applicantRow);
    const endRow = Math.max(opponentRow, applicantRow);
    const rowCount = endRow - startRow + 1;
    
    // 名前と学年のデータを取得（A列とB列）
    const dataRange = sheet.getRange(startRow, 1, rowCount, 2);
    const data = dataRange.getValues();
    
    // 申込者のデータを保存
    const applicantData = data[applicantRow - startRow];
    
    // 対戦相手の直上に申込者を移動するためのデータ配列を作成
    const newData = [];
    
    for (let i = 0; i < data.length; i++) {
      if (i === opponentRow - startRow) {
        // 申込者のデータを追加
        newData.push(applicantData);
        // その直後に対戦相手のデータを追加
        newData.push(data[i]);
      } else if (i !== applicantRow - startRow) {
        // 申込者以外のデータを追加
        newData.push(data[i]);
      }
    }
    
    // 新しいデータをシートに書き込む
    dataRange.setValues(newData);
    
    // ランクを更新（C列）
    const rankRange = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1); // 1行目を除く
    const rankValues = rankRange.getValues();
    for (let i = 0; i < rankValues.length; i++) {
      rankValues[i][0] = i + 1; // 1から始まるランク
    }
    rankRange.setValues(rankValues);
    
    logMessage(LOG_LEVEL.INFO, `ランキングを更新しました: ${applicantData[0]}を${opponentRow}行目に移動しました`);
  } catch (error) {
    logMessage(LOG_LEVEL.ERROR, `ランキング更新中にエラーが発生しました: ${error.message}`);
  }
}

// 男子シートと女子シートのB列の数値を1増やす関数
function incrementGrades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const maleSheet = ss.getSheetByName(SHEET_NAMES.MALE);
  const femaleSheet = ss.getSheetByName(SHEET_NAMES.FEMALE);
  
  // 男子シートのB列を取得
  const maleLastRow = maleSheet.getLastRow();
  const maleGrades = maleSheet.getRange(2, COLUMN_INDICES.GRADE + 1, maleLastRow - 1, 1).getValues();
  
  // 女子シートのB列を取得
  const femaleLastRow = femaleSheet.getLastRow();
  const femaleGrades = femaleSheet.getRange(2, COLUMN_INDICES.GRADE + 1, femaleLastRow - 1, 1).getValues();
  
  // 男子シートのB列の数値を1増やす
  for (let i = 0; i < maleGrades.length; i++) {
    maleGrades[i][0] = maleGrades[i][0] + 1;
  }
  maleSheet.getRange(2, COLUMN_INDICES.GRADE + 1, maleGrades.length, 1).setValues(maleGrades);
  
  // 女子シートのB列の数値を1増やす
  for (let i = 0; i < femaleGrades.length; i++) {
    femaleGrades[i][0] = femaleGrades[i][0] + 1;
  }
  femaleSheet.getRange(2, COLUMN_INDICES.GRADE + 1, femaleGrades.length, 1).setValues(femaleGrades);
}


