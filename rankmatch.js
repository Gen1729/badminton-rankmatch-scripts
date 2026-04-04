// 編集履歴
// sortRankMatchSchedule関数において、並べ替えの順序を変更（2025/5/14　入野）
// processCancelRequest関数の修正（2025/4/28　入野）
// processRankMatchRequest関数を実装（2025/3/2　入野）

// 定数定義 - マジックナンバーやリテラル値を一箇所にまとめています
const START_ROW = 4;
const CLUB_MATCH_SLOTS = ["部活中（1試合目）", "部活中（2試合目）", "部活中（3試合目）"];
const CANCEL_FLAG = "キャンセル";
const OUTSIDE_CLUB_SLOT = "部活時間外";
const FRIDAY_DAY_INDEX = 5;
const HEADER_ROW = 1;
const APPLICANT_COLUMN = 0;
const OPPONENT_COLUMN = 1;
const DATE_COLUMN = 2;
const TIME_SLOT_COLUMN = 3;
const FORM_TIMESTAMP_INDEX = 0;
const FORM_APPLICANT_INDEX = 1;
const FORM_OPPONENT_INDEX = 2;
const FORM_DATE_INDEX = 3;
const FORM_TIMESLOT_INDEX = 4;
const FORM_CANCEL_INDEX = 5;
const FORM_MODIFIED_DATE_INDEX = 6;
const FORM_MODIFIED_TIMESLOT_INDEX = 7;
const NOTE_COLUMN = 8;
const ALLOWED_MATCHES_CELL = "C1";
const ALLOWED_DAYS_CELL = "E1";

// スプレッドシートの取得
const ss = SpreadsheetApp.getActiveSpreadsheet();
const formResponsesSheet = ss.getSheetByName("Form Responses");
const rankMatchScheduleSheet = ss.getSheetByName("ランク戦日程");

// ランク戦リクエストの処理を統括するメイン関数
function processRankMatchRequest(e) {
  try {
    // フォーム送信されたデータを取得し、ログに出力
    const formData = e.values;
    Logger.log("新規リクエスト受信: " + JSON.stringify(formData));

    // 各項目を抽出
    const timestamp = formData[FORM_TIMESTAMP_INDEX];
    const applicant = formData[FORM_APPLICANT_INDEX];
    const opponent = formData[FORM_OPPONENT_INDEX];
    const originalDate = new Date(formData[FORM_DATE_INDEX]);
    const timeSlot = formData[FORM_TIMESLOT_INDEX];
    const cancelFlag = formData[FORM_CANCEL_INDEX];
    // 修正リクエストの入力値を文字列として取得し、空白の場合はnullとする
    const modifiedDateRaw = formData[FORM_MODIFIED_DATE_INDEX] ? formData[FORM_MODIFIED_DATE_INDEX].toString().trim() : "";
    const modifiedTimeSlotRaw = formData[FORM_MODIFIED_TIMESLOT_INDEX] ? formData[FORM_MODIFIED_TIMESLOT_INDEX].toString().trim() : "";
    const modifiedDateValue = modifiedDateRaw !== "" ? new Date(modifiedDateRaw) : null;
    const modifiedTimeSlot = modifiedTimeSlotRaw !== "" ? modifiedTimeSlotRaw : null;

    // 今日の日付（0時正規化）
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // キャンセルリクエストの場合
    if (cancelFlag === CANCEL_FLAG) {
      Logger.log("キャンセルリクエストを処理します。");
      processCancelRequest(applicant, opponent, originalDate, timeSlot);
      return;
    }
    // 変更リクエストの場合（変更後の日付または時間帯が入力されている場合のみ）
    if (modifiedDateValue || modifiedTimeSlot) {
      Logger.log("日付/時間帯変更リクエストを処理します。");
      processModifyRequest(applicant, opponent, originalDate, timeSlot, modifiedDateValue, modifiedTimeSlot);
      return;
    }
    // 通常リクエストの場合
    Logger.log("通常リクエストを処理します。");
    processNormalRequest(applicant, opponent, originalDate, timeSlot);
  } catch (error) {
    // エラーハンドリング：processRankMatchRequest内で発生した例外をログ出力
    Logger.log("processRankMatchRequest エラー: " + error);
  }
}

// ----------------------------
// 通常リクエストの処理
// ----------------------------
function processNormalRequest(applicant, opponent, originalDate, timeSlot) {
  try {
    // 今日の日付（0時正規化）を取得
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let rejectionReasons = [];

    // ① 申込者と対戦相手が同一であれば拒否
    if (applicant === opponent) {
      rejectionReasons.push("自分自身との試合はリクエストできません");
      Logger.log("自分自身との試合リクエストを検出し、キャンセルしました。");
      return;
    }

    // ② 日付のバリデーション（過去の日付の場合はリジェクト）
    const allowedDaysOffset = Number(formResponsesSheet.getRange(ALLOWED_DAYS_CELL).getValue());
    rejectionReasons = rejectionReasons.concat(validateRequestDate(originalDate, today, allowedDaysOffset, ""));

    // ③ 部活中リクエストのチェック：金曜日かどうか及び重複の有無
    if (CLUB_MATCH_SLOTS.indexOf(timeSlot) !== -1) {
      if (originalDate.getDay() !== FRIDAY_DAY_INDEX) {
        rejectionReasons.push("部活中の時間帯は金曜日のみ選択可能です");
      } else {
        // 金曜日の場合のみ重複チェックを行う
        if (isDuplicateClubMatch(originalDate, timeSlot)) {
          rejectionReasons.push("選択された部活中の時間帯は既に予約されています");
        }
      }
    }

    // ④ 申込者の月間試合数制限チェック
    const allowedMatches = Number(formResponsesSheet.getRange(ALLOWED_MATCHES_CELL).getValue());
    const monthlyCount = countMonthlyMatchesFromFormResponses(applicant, originalDate);
    if (monthlyCount > allowedMatches) {
      rejectionReasons.push("申込者の試合数が月の上限に達しています");
    }

    if (rejectionReasons.length > 0) {
      // リクエスト拒否：拒否理由をログに出力
      Logger.log("通常リクエスト拒否: " + rejectionReasons.join(", "));
      // Form Responsesシートの備考欄に拒否理由を記入
      updateFormResponseNote(rejectionReasons.join("\n"));
    } else {
      // リクエスト承認：ランク戦日程シートに追加
      addToRankMatchSchedule(applicant, opponent, originalDate, timeSlot, "");
      Logger.log("通常リクエスト承認: " + applicant + " vs " + opponent + " on " + formatDateJP(originalDate) + ", " + timeSlot);
    }
    // ランク戦日程シートの並び替えを実施
    sortRankMatchSchedule();
  } catch (error) {
    Logger.log("processNormalRequest エラー: " + error);
  }
}

// ----------------------------
// キャンセルリクエストの処理
// ----------------------------
function processCancelRequest(applicant, opponent, originalDate, timeSlot) {
  try {
    // ランク戦日程シートから該当する行を削除
    const lastRow = rankMatchScheduleSheet.getLastRow();
    if (lastRow <= HEADER_ROW) {
      // ヘッダーのみの場合は一致するリクエストがないためメッセージを表示
      updateFormResponseNote("キャンセル対象のリクエストが見つかりませんでした。");
      Logger.log("キャンセル対象のリクエストが見つかりませんでした。");
      return;
    }
    const data = rankMatchScheduleSheet.getRange(HEADER_ROW + 1, 1, lastRow - 1, 4).getValues();
    let rowToDelete = -1;
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (
        row[APPLICANT_COLUMN] === applicant &&
        row[OPPONENT_COLUMN] === opponent &&
        isSameDate(new Date(row[DATE_COLUMN]), originalDate) &&
        row[TIME_SLOT_COLUMN] === timeSlot
      ) {
        rowToDelete = i + HEADER_ROW + 1;
        break;
      }
    }
    if (rowToDelete !== -1) {
      // スケジュールから削除
      rankMatchScheduleSheet.deleteRow(rowToDelete);
      // キャンセル成功のメッセージを備考欄に記入（現在のリクエスト行へ）
      updateFormResponseNote("リクエストは正常にキャンセルされました。");

      // Form Responsesシートで該当する元リクエストを検索し、I列に「キャンセル済」と入力
      const responsesLastRow = formResponsesSheet.getLastRow();
      if (responsesLastRow > 1) {
        // applicant(B列)～timeSlot(E列)を取得
        const startCol = APPLICANT_COLUMN + 2; // B列
        const numCols = TIME_SLOT_COLUMN - APPLICANT_COLUMN + 1; // 0→3 なら 4列
        const responsesData = formResponsesSheet
          .getRange(2, startCol, responsesLastRow - 1, numCols)
          .getValues();
        for (let j = 0; j < responsesData.length; j++) {
          const resp = responsesData[j];
          if (
            resp[APPLICANT_COLUMN] === applicant &&
            resp[OPPONENT_COLUMN] === opponent &&
            isSameDate(new Date(resp[DATE_COLUMN]), originalDate) &&
            resp[TIME_SLOT_COLUMN] === timeSlot
          ) {
            const formRow = j + 2;
            formResponsesSheet.getRange(formRow, 9).setValue("キャンセル済");
            break;
          }
        }
      }
      Logger.log("キャンセル対象のリクエストを削除しました: 行 " + rowToDelete);
    } else {
      // 一致するリクエストが見つからなかった場合のメッセージを備考欄に記入
      updateFormResponseNote("キャンセル対象のリクエストが見つかりませんでした。");
      Logger.log("キャンセル対象のリクエストが見つかりませんでした。");
    }
  } catch (error) {
    Logger.log("processCancelRequest エラー: " + error);
    // エラーが発生した場合も備考欄にエラーメッセージを記入
    updateFormResponseNote("キャンセル処理中にエラーが発生しました: " + error);
  }
}

// ----------------------------
// 日付または時間帯変更リクエストの処理
// ----------------------------
function processModifyRequest(applicant, opponent, originalDate, originalTimeSlot, modifiedDateValue, modifiedTimeSlot) {
  try {
    // 今日の日付（0時正規化）を取得
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let rejectionReasons = [];
    // 変更後の内容（入力がなければ元の値を維持）
    const newDate = modifiedDateValue ? new Date(modifiedDateValue) : originalDate;
    const newTimeSlot = modifiedTimeSlot ? modifiedTimeSlot : originalTimeSlot;

    // ① 申込者と対戦相手が同一なら拒否
    if (applicant === opponent) {
      rejectionReasons.push("自分自身との試合はリクエストできません");
      Logger.log("自分自身との試合リクエストを検出し、キャンセルしました。");
      updateFormResponseNote("自分自身との試合はリクエストできません");
      return;
    }

    // ② 新しい日付のバリデーション（新しい日付についてチェック）
    const allowedDaysOffset = Number(formResponsesSheet.getRange(ALLOWED_DAYS_CELL).getValue());
    rejectionReasons = rejectionReasons.concat(validateRequestDate(newDate, today, allowedDaysOffset, "変更後の"));

    // ③ 変更後の部活中リクエストチェック：金曜日かつ重複の有無
    if (CLUB_MATCH_SLOTS.indexOf(newTimeSlot) !== -1) {
      if (newDate.getDay() !== FRIDAY_DAY_INDEX) {
        rejectionReasons.push("変更後の部活中の時間帯は金曜日のみ選択可能です");
      } else {
        // 金曜日の場合のみ重複チェックを行う
        if (isDuplicateClubMatch(newDate, newTimeSlot)) {
          rejectionReasons.push("変更後の部活中の時間帯は既に予約されています");
        }
      }
    }

    // 元のリクエスト行を検索
    const lastRow = rankMatchScheduleSheet.getLastRow();
    if (lastRow <= HEADER_ROW) {
      // ヘッダーのみの場合は一致するリクエストがないためメッセージを表示
      updateFormResponseNote("変更対象のリクエストが見つかりませんでした。");
      Logger.log("変更対象のリクエストが見つかりませんでした。");
      return;
    }
    
    const data = rankMatchScheduleSheet.getRange(HEADER_ROW + 1, 1, lastRow - 1, 4).getValues();
    let rowToModify = -1;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[APPLICANT_COLUMN] === applicant && 
          row[OPPONENT_COLUMN] === opponent && 
          isSameDate(new Date(row[DATE_COLUMN]), originalDate) && 
          row[TIME_SLOT_COLUMN] === originalTimeSlot) {
        rowToModify = i + HEADER_ROW + 1; // ヘッダー行(1行目)を考慮して+2
        break;
      }
    }
    
    if (rowToModify === -1) {
      // 一致するリクエストが見つからなかった場合のメッセージを備考欄に記入
      updateFormResponseNote("変更対象のリクエストが見つかりませんでした。");
      Logger.log("変更対象のリクエストが見つかりませんでした。");
      return;
    }

    if (rejectionReasons.length > 0) {
      // 変更リクエスト拒否の場合：拒否理由を備考欄に記入
      updateFormResponseNote(rejectionReasons.join("\n"));
      Logger.log("変更リクエスト拒否: " + rejectionReasons.join(", "));
    } else {
      // 承認の場合：元のリクエスト行を削除し、変更後のリクエストを追加
      rankMatchScheduleSheet.deleteRow(rowToModify);
      addToRankMatchSchedule(applicant, opponent, newDate, newTimeSlot, "");
      updateFormResponseNote("リクエストは正常に変更されました。【変更前】" + formatDateJP(originalDate) + ", " + originalTimeSlot +
                  " → 【変更後】" + formatDateJP(newDate) + ", " + newTimeSlot);
      Logger.log("変更リクエスト承認: " + applicant + " vs " + opponent + "【変更前】" + formatDateJP(originalDate) + ", " + originalTimeSlot +
                  " → 【変更後】" + formatDateJP(newDate) + ", " + newTimeSlot);
    }
    // ランク戦日程シートの並び替えを実施
    sortRankMatchSchedule();
  } catch (error) {
    Logger.log("processModifyRequest エラー: " + error);
    updateFormResponseNote("変更処理中にエラーが発生しました: " + error);
  }
}

// ----------------------------
// 以下、ヘルパー関数群
// ----------------------------

// Form Responsesシートの最新行の備考欄を更新する関数
function updateFormResponseNote(note) {
  const lastRow = formResponsesSheet.getLastRow();
  if (lastRow > 1) { // ヘッダー行がある場合
    formResponsesSheet.getRange(lastRow, NOTE_COLUMN + 1).setValue(note); // インデックスは0始まりなので+1
    Logger.log("Form Responsesシートの備考欄を更新しました: " + note);
  }
}

// 日付を「yyyy年MM月dd日」の形式に変換する関数
function formatDateJP(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy年MM月dd日");
}

// 2つの日付が同一か判定する（時刻は無視）
function isSameDate(date1, date2) {
  if (!(date1 instanceof Date) || !(date2 instanceof Date) ||
      isNaN(date1.getTime()) || isNaN(date2.getTime())) {
    return false;
  }
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

// ランク戦日程シート内で、同一の日付・時間帯の部活中リクエストが既に存在するかチェックする関数
function isDuplicateClubMatch(date, timeSlot) {
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return false; // ヘッダーのみの場合は重複なし
  
  const data = rankMatchScheduleSheet.getRange(HEADER_ROW + 1, 1, lastRow - 1, 4).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (isSameDate(new Date(row[DATE_COLUMN]), date) && row[TIME_SLOT_COLUMN] === timeSlot) {
      return true;
    }
  }
  return false;
}

// フォーム回答シートから申込者の指定した月の承認済み試合数をカウントする関数
function countMonthlyMatchesFromFormResponses(applicant, date) {
  const lastRow = formResponsesSheet.getLastRow();
  if (lastRow <= 1) return 0; // ヘッダーのみの場合は0

  // 対象年月の開始日と終了日を設定
  const targetYear = date.getFullYear();
  const targetMonth = date.getMonth();
  const startDate = new Date(targetYear, targetMonth, 1);
  const endDate = new Date(targetYear, targetMonth + 1, 0);

  // 指定範囲のデータを一括取得
  const data = formResponsesSheet.getRange(2, 1, lastRow - 1, NOTE_COLUMN + 1).getValues();

  let count = 0;
  // 日付変更リクエストの場合に元の日付と変更後の日付を追跡するためのセット
  const processedRequests = new Set();
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowApplicant = row[FORM_APPLICANT_INDEX];
    const rowDate = new Date(row[FORM_DATE_INDEX]);
    const cancelFlag = row[FORM_CANCEL_INDEX];
    const note = row[NOTE_COLUMN]; // 備考欄の値を取得
    const modifiedDate = row[FORM_MODIFIED_DATE_INDEX] ? new Date(row[FORM_MODIFIED_DATE_INDEX]) : null;
    
    // 申込者が一致し、日付が対象月内で、キャンセルでなく、備考欄が空のリクエストのみカウント
    if (rowApplicant === applicant && 
        rowDate >= startDate && 
        rowDate <= endDate && 
        cancelFlag !== CANCEL_FLAG && 
        (!note || note.trim() === "")) {
      
      // 日付変更リクエストの場合、元の日付と変更後の日付が同じ月内なら1回だけカウント
      const requestKey = rowApplicant + "_" + row[FORM_OPPONENT_INDEX] + "_" + rowDate.getTime();
      
      // 変更後の日付が指定されていて、同じ月内の場合
      if (modifiedDate && 
          modifiedDate.getFullYear() === targetYear && 
          modifiedDate.getMonth() === targetMonth) {
        
        // この組み合わせが既にカウントされていなければカウント
        if (!processedRequests.has(requestKey)) {
          count++;
          processedRequests.add(requestKey);
        }
      } else if (!modifiedDate) {
        // 変更リクエストでない通常のリクエスト
        count++;
      }
    }
  }

  Logger.log(applicant + "の" + targetYear + "年" + (targetMonth + 1) + "月の試合数: " + count);
  return count;
}

// 申込者の指定した月（元の日付の月）の承認済み試合数をカウントする関数
function countMonthlyMatches(applicant, date) {
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return 0; // ヘッダーのみの場合は0

  // 対象年月の開始日と終了日を設定
  const targetYear = date.getFullYear();
  const targetMonth = date.getMonth();
  const startDate = new Date(targetYear, targetMonth, 1);
  const endDate = new Date(targetYear, targetMonth + 1, 0);

  // 指定範囲のデータを一括取得
  const data = rankMatchScheduleSheet.getRange(HEADER_ROW + 1, 1, lastRow - 1, 3).getValues();

  let count = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowDate = new Date(row[DATE_COLUMN]);
    if (rowDate >= startDate && rowDate <= endDate && row[APPLICANT_COLUMN] === applicant) {
      count++;
    }
  }

  Logger.log(applicant + "の" + targetYear + "年" + (targetMonth + 1) + "月の試合数: " + count);
  return count;
}

// ランク戦日程シートに新しいリクエストを追加する関数
function addToRankMatchSchedule(applicant, opponent, date, timeSlot, note) {
  // ランク戦日程シートの最終行の次の行に追加
  const lastRow = rankMatchScheduleSheet.getLastRow();
  const newRow = [applicant, opponent, date, timeSlot, note];
  rankMatchScheduleSheet.getRange(lastRow + 1, 1, 1, 5).setValues([newRow]);
  Logger.log("ランク戦日程シートに追加: " + applicant + " vs " + opponent + " on " + formatDateJP(date) + ", " + timeSlot);
}

// ランク戦日程シートの並び替えを行う関数
function sortRankMatchSchedule() {
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return; // ヘッダーのみの場合は処理しない
  
  // 時間帯の優先順位定義
  const timeSlotOrder = {
    [OUTSIDE_CLUB_SLOT]: -1,
    [CLUB_MATCH_SLOTS[0]]: 0,
    [CLUB_MATCH_SLOTS[1]]: 1,
    [CLUB_MATCH_SLOTS[2]]: 2
  };
  
  // データを取得して並び替え用の配列を作成
  const data = rankMatchScheduleSheet.getRange(HEADER_ROW + 1, 1, lastRow - 1, 5).getValues();
  const sortableData = data.map((row, index) => {
    return {
      index: index + HEADER_ROW + 1, // 実際の行番号（ヘッダー行を考慮）
      row: row,
      date: new Date(row[DATE_COLUMN]),
      timeSlotOrder: timeSlotOrder[row[TIME_SLOT_COLUMN]] || 999
    };
  });
  
  // 日付の昇順、同じ日付なら時間帯の優先順位で並び替え
  sortableData.sort((a, b) => {
    const dateDiff = a.date - b.date;
    if (dateDiff === 0) {
      return a.timeSlotOrder - b.timeSlotOrder;
    }
    return dateDiff;
  });
  
  // 並び替えたデータを新しい配列に格納
  const sortedData = sortableData.map(item => item.row);
  
  // 並び替えたデータをシートに書き戻す
  rankMatchScheduleSheet.getRange(HEADER_ROW + 1, 1, sortedData.length, 5).setValues(sortedData);
  
  Logger.log("ランク戦日程シートの並び替えが完了しました。");
}

// 日付が過去またはリクエスト受付期間を超えているかチェックする関数。  
// labelには「元の」または「新しい」を指定し、エラーメッセージに反映させます。
function validateRequestDate(date, today, allowedDaysOffset, label) {
  let errors = [];
  // 過去の日付でのリクエストは拒否する
  if (date.getTime() < today.getTime()) {
    errors.push(label + "日付が過去です");
  }
  let maxDate = new Date(today);
  maxDate.setDate(maxDate.getDate() + allowedDaysOffset);
  if (date.getTime() > maxDate.getTime()) {
    errors.push(label + "日付がリクエスト受付期間を超えています");
  }
  return errors;
}

// 利用可能な部活中時間帯を取得する関数
function getAvailableClubSlots(date) {
  const clubSlots = CLUB_MATCH_SLOTS;
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return clubSlots; // ヘッダーのみの場合は全て利用可能
  
  const data = rankMatchScheduleSheet.getRange(HEADER_ROW + 1, 1, lastRow - 1, 4).getValues();
  const usedSlots = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (isSameDate(new Date(row[DATE_COLUMN]), date) && clubSlots.indexOf(row[TIME_SLOT_COLUMN]) !== -1) {
      usedSlots.push(row[TIME_SLOT_COLUMN]);
    }
  }
  
  // 利用されていない時間帯を返す
  return clubSlots.filter(slot => usedSlots.indexOf(slot) === -1);
}


