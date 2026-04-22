const ss = SpreadsheetApp.getActiveSpreadsheet();
const maleSheet = ss.getSheetByName("男子");
const femaleSheet = ss.getSheetByName("女子");
const rankMatchScheduleSheet = ss.getSheetByName("ランク戦日程");
const configSheet = ss.getSheetByName("設定一覧");
const maleSheetByDepartment = ss.getSheetByName("男子（学科別）");
const femaleSheetByDepartment = ss.getSheetByName("女子（学科別）");

const scheduleFormResponsesSheet = ss.getSheetByName("日程報告");
const resultFormResponsesSheet = ss.getSheetByName("結果報告");

const MATCH_SCHEDULING_FORM_ID = '';
const MATCH_RESULT_FORM_ID = '';
const MATCH_RESULT_CHECK_FORM_ID = '';

const MAX_RANK_DIFFERENCE_CELL = 'B1';
const MATCH_ACCEPT_DAY_LIMIT_CELL = 'B3';
const SAME_OPPONENT_COOLDOWN_DAYS_CELL = 'B5';
const MONTH_APPLICATION_LIMIT_CELL = 'B7';
const FRIDAY_MATCH_NUMBER_CELL = 'B21';

const MAX_RANK_DIFFERENCE = configSheet.getRange(MAX_RANK_DIFFERENCE_CELL).getValue();
const MATCH_ACCEPT_DAY_LIMIT = configSheet.getRange(MATCH_ACCEPT_DAY_LIMIT_CELL).getValue();
const SAME_OPPONENT_COOLDOWN_DAYS = configSheet.getRange(SAME_OPPONENT_COOLDOWN_DAYS_CELL).getValue();
const FRIDAY_MATCH_NUMBER = configSheet.getRange(FRIDAY_MATCH_NUMBER_CELL).getValue();

const timeSlotSortOrder = { '部活時間外': 0, '部活中': 1, 'その他': 2 };

const HEADER_ROW_OFFSET = 1;

// 実行中キャッシュ
// シート書き込み時に dirty を立て、次回参照時に再取得する
const dataCache = {
  rankMatch: null,
  rankMatchDirty: true,
  maleRank: null,
  maleRankDirty: true,
  femaleRank: null,
  femaleRankDirty: true,
};

function resetDataCache() {
  dataCache.rankMatch = null;
  dataCache.rankMatchDirty = true;
  dataCache.maleRank = null;
  dataCache.maleRankDirty = true;
  dataCache.femaleRank = null;
  dataCache.femaleRankDirty = true;
}

function markRankMatchDirty() {
  dataCache.rankMatchDirty = true;
}

function markMaleRankDirty() {
  dataCache.maleRankDirty = true;
}

function markFemaleRankDirty() {
  dataCache.femaleRankDirty = true;
}

function getRankMatchData() {
  if (!dataCache.rankMatchDirty && dataCache.rankMatch) {
    return dataCache.rankMatch;
  }
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if (lastRow <= HEADER_ROW_OFFSET) {
    dataCache.rankMatch = [];
  } else {
    dataCache.rankMatch = rankMatchScheduleSheet
      .getRange(HEADER_ROW_OFFSET + 1, 1, lastRow - 1, RANK_MATCH_SHEET_MAX_COLUMN)
      .getValues();
  }
  dataCache.rankMatchDirty = false;
  return dataCache.rankMatch;
}

function getMaleRankData() {
  if (!dataCache.maleRankDirty && dataCache.maleRank) {
    return dataCache.maleRank;
  }
  const lastRow = maleSheet.getLastRow();
  if (lastRow <= HEADER_ROW_OFFSET) {
    dataCache.maleRank = [];
  } else {
    dataCache.maleRank = maleSheet
      .getRange(HEADER_ROW_OFFSET + 1, 1, lastRow - 1, RANKING_SHEET_MAX_COLUMN)
      .getValues();
  }
  dataCache.maleRankDirty = false;
  return dataCache.maleRank;
}

function getFemaleRankData() {
  if (!dataCache.femaleRankDirty && dataCache.femaleRank) {
    return dataCache.femaleRank;
  }
  const lastRow = femaleSheet.getLastRow();
  if (lastRow <= HEADER_ROW_OFFSET) {
    dataCache.femaleRank = [];
  } else {
    dataCache.femaleRank = femaleSheet
      .getRange(HEADER_ROW_OFFSET + 1, 1, lastRow - 1, RANKING_SHEET_MAX_COLUMN)
      .getValues();
  }
  dataCache.femaleRankDirty = false;
  return dataCache.femaleRank;
}

// 日程シートの列の定数
const APPLICANT_ID_COLUMN = 0;
const APPLICANT_NAME_COLUMN = 1;
const OPPONENT_ID_COLUMN = 2;
const OPPONENT_NAME_COLUMN = 3;
const MATCH_DATE_COLUMN = 4;
const MATCH_TIMESLOT_COLUMN = 5;
const MATCH_RESULT_FLAG_COLUMN = 6;
const MATCH_RESULT_COLUMN = 7;
const MODIFY_FLAG_COLUMN = 11;
const SCHEDULE_FORM_TIMESTAMP_COLUMN = 12;
const RESULT_FORM_TIMESTAMP_COLUMN = 13;

// ランク戦日程シートの列数
const RANK_MATCH_SHEET_MAX_COLUMN = 14;

// ランキングシートの列の定数
const PLAYER_ID_COLUMN = 1;
const PLAYER_NAME_COLUMN = 2;
const CAN_PLAY_FLAG_COLUMN = 4;
const MATCH_LIMIT_COLUMN = 5;
const MATCH_MORE_FLAG_COLUMN = 6;

// ランキング表シートの列数
const RANKING_SHEET_MAX_COLUMN = 7;

// フォームを受け取った時の分岐
// 日程報告か結果報告か
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  const responseSheet = e.range.getSheet();
  const responseRow = e.range.getRow();
  const sheetName = responseSheet.getName();
  let lockAcquired = false;

  try {
    lock.waitLock(30000);
    lockAcquired = true;
  } catch (err) {
    const message = '他の処理が終わっていないため、このリクエストをキャンセルします。';
    console.log(message);
    writeLogsInFormResponse(responseSheet, responseRow, message);
    return;
  }

  try {
    resetDataCache();
    if (sheetName === '日程報告') {
      handleSchedule(e, responseSheet, responseRow);
      return;
    }
    if (sheetName === '結果報告') {
      handleResult(e, responseSheet, responseRow);
      return;
    }
    throw new Error(`想定外の送信先シートです: ${sheetName}`);
  } catch (err) {
    writeLogsInFormResponse(responseSheet, responseRow, err);
    console.error(err);
    throw err;
  } finally {
    if(lockAcquired){
      lock.releaseLock();
    }
  }
}

// 日程報告を受け取った時の関数
function handleSchedule(e, responseSheet, responseRow){
  try {
    const formData = e.values;
    console.log("新規日程報告受信: " + JSON.stringify(formData));
    const timestamp = formData[0];
    const applicantRowString = formData[1];
    const opponentRowString = formData[2];
    const originalDate = new Date(formData[3]);
    const timeSlot = formData[4];
    const cancelFlag = formData[5];
    const modifiedDate = formData[6] ? new Date(formData[6]) : null;
    const modifiedTimeSlot = formData[7] ? formData[7] : null;

    const today = new Date();
    today.setHours(0,0,0,0);
    const applicationScope = new Date();
    applicationScope.setHours(0,0,0,0);
    applicationScope.setDate(applicationScope.getDate() + MATCH_ACCEPT_DAY_LIMIT);

    const applicant = parsePlayerLabel(applicantRowString);
    const opponent = parsePlayerLabel(opponentRowString);
    
    if(applicant.id === opponent.id){
      console.log('対戦する人が同一人物です。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '対戦する人が同一人物です。入力は無効です。');
      return;
    }

    if(applicant.gender !== opponent.gender){
      console.log('対戦する人の性別が違います。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '対戦する人の性別が違います。入力は無効です。');
      return;
    }

    if(originalDate.getTime() < today.getTime()){
      console.log('日付が過去のものです。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '日付が過去のものです。入力は無効です。');
      return;
    }

    if(applicationScope.getTime() < originalDate.getTime()){
      console.log('日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。');
      return;
    }

    if(modifiedDate && applicationScope.getTime() < modifiedDate.getTime()){
      console.log('日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。');
      return;
    }

    if(cancelFlag === 'キャンセル'){
      console.log('キャンセル操作を実行します。');
      processCancelRequest(applicant,opponent,originalDate,timeSlot,responseSheet,responseRow);
      SpreadsheetApp.flush();
      return;
    }

    if(modifiedDate && modifiedTimeSlot){
      if(originalDate.getTime() === modifiedDate.getTime() && timeSlot === modifiedTimeSlot){
        console.log('変更前と変更後の日付/時間帯が同じです。');
        writeLogsInFormResponse(responseSheet, responseRow, '変更前と変更後の日付/時間帯が同じです。');
        return;
      }
      console.log('日付/時間帯変更操作を実行します。');
      processModifyRequest(applicant,opponent,originalDate,timeSlot,modifiedDate,modifiedTimeSlot,responseSheet,responseRow);
      sortRankMatchSchedule();
      SpreadsheetApp.flush();
      return;
    }

    if(modifiedDate || modifiedTimeSlot){
      console.log('日程追加の場合はどちらも空白でなければなりません。また、日付/時間帯変更の場合は日付と時間のどちらの入力も必要です。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '日程追加の場合はどちらも空白でなければなりません。また、日付/時間帯変更の場合は日付と時間のどちらの入力も必要です。入力は無効です。');
      return;
    }

    console.log('日程追加操作を実行します。');
    processNormalRequest(applicant,opponent,originalDate,timeSlot,responseSheet,responseRow);
    sortRankMatchSchedule();
    SpreadsheetApp.flush();

  } catch (err) {
    console.log('日程報告中に予期せぬエラーが発生しました。' + err);
    writeLogsInFormResponse(responseSheet, responseRow, '日程報告中に予期せぬエラーが発生しました。' + err);
  }
}

// 結果報告を受け取った時の関数
function handleResult(e, responseSheet, responseRow){
  try {
    const formData = e.values;
    console.log("結果報告受信: " + JSON.stringify(formData));
    const timestamp = formData[0];
    const applicantRowString = formData[1];
    const opponentRowString = formData[2];
    const matchResult = formData[3];
    const game1Score = formData[4] ? formData[4] : null;
    const game2Score = formData[5] ? formData[5] : null;
    const game3Score = formData[6] ? formData[6] : null;

    const applicant = parsePlayerLabel(applicantRowString);
    const opponent = parsePlayerLabel(opponentRowString);

    if(applicant.id === opponent.id){
      console.log('対戦する人が同一人物です。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '対戦する人が同一人物です。入力は無効です。');
      return;
    }

    if(applicant.gender !== opponent.gender){
      console.log('対戦する人の性別が違います。入力は無効です。');
      writeLogsInFormResponse(responseSheet, responseRow, '対戦する人の性別が違います。入力は無効です。');
      return;
    }

    console.log('結果報告の書き込みを開始します。');
    writeMatchResult(applicant,opponent,matchResult,game1Score,game2Score,game3Score,responseSheet,responseRow);
    updateFormDropdown();
    SpreadsheetApp.flush();
  } catch (err) {
    console.log('結果報告中に予期せぬエラーが発生しました。' + err);
    writeLogsInFormResponse(responseSheet, responseRow, '結果報告中に予期せぬエラーが発生しました。' + err);
  }
}

// 以下日程報告フォームの処理

//--------------------
// キャンセル操作
//--------------------
function processCancelRequest(applicant,opponent,originalDate,timeSlot,responseSheet,responseRow){
  try {
    const applicantID = applicant.id;
    const opponentID  = opponent .id;

    console.log(applicantID);
    console.log(opponentID);

    const matchData = getRankMatchData();
    if(matchData.length === 0){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse(responseSheet, responseRow, '対象の試合が見つかりませんでした。入力を再度確認してください。');
      return;
    }

    let isMale = false;

    if(applicant.gender === '男'){
      isMale = true;
    }

    let deleteFlag = false;
    for(let idx = 0;idx < matchData.length;idx++){
      const row = matchData[idx];
      if(row[APPLICANT_ID_COLUMN] === applicantID && row[OPPONENT_ID_COLUMN] === opponentID && (new Date(row[MATCH_DATE_COLUMN])).getTime() === originalDate.getTime()){
        if(timeSlot === '金曜部活内'){
          if(row[MATCH_TIMESLOT_COLUMN] === '部活時間外' || row[MATCH_TIMESLOT_COLUMN] === 'その他'){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            continue;
          }
        }else{
          if(row[MATCH_TIMESLOT_COLUMN] !== timeSlot){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            continue;
          }
        }

        if(row[MATCH_RESULT_FLAG_COLUMN] !== ''){
          console.log('該当の試合は存在しますが、すでに結果報告を受け取っているためキャンセルはできません。');
          writeLogsInFormResponse(responseSheet, responseRow, '該当の試合は存在しますが、すでに結果報告を受け取っているためキャンセルはできません。');
          continue;
        }

        if(row[MATCH_TIMESLOT_COLUMN] !== '部活時間外' && row[MATCH_TIMESLOT_COLUMN] !== 'その他'){
          if(Number(row[MATCH_TIMESLOT_COLUMN][4]) !== FRIDAY_MATCH_NUMBER){
            narrowSchedule(originalDate,Number(row[MATCH_TIMESLOT_COLUMN][4]));
          }
        }

        rankMatchScheduleSheet.deleteRow(idx + HEADER_ROW_OFFSET + 1);
        markRankMatchDirty();
        manageChallenge(applicantID,false,isMale,responseSheet,responseRow);
        console.log('該当の試合を削除しました。');
        deleteFlag = true;
        break;
      }
    }
    if(!deleteFlag){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse(responseSheet, responseRow, '対象の試合が見つかりませんでした。入力を再度確認してください。');
    }

  } catch (err) {
    console.log('キャンセル処理中にエラーが発生しました。' + err);
    writeLogsInFormResponse(responseSheet, responseRow, 'キャンセル処理中にエラーが発生しました。' + err);
  }
}

//--------------------
// 日程/時間帯変更操作
//--------------------
function processModifyRequest(applicant,opponent,originalDate,timeSlot,modifiedDate,modifiedTimeSlot,responseSheet,responseRow){
  try {
    const applicantID = applicant.id;
    const opponentID  = opponent .id;

    console.log(applicantID);
    console.log(opponentID);

    const matchData = getRankMatchData();
    if(matchData.length === 0){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse(responseSheet, responseRow, '対象の試合が見つかりませんでした。入力を再度確認してください。');
      return;
    }

    let isMale = false;

    if(applicant.gender === '男'){
      isMale = true;
    }

    if(modifiedTimeSlot === '金曜部活内'){
      if(!isFriday(modifiedDate)){
        console.log('部活内を選んだ場合、日付は金曜日でなくてはいけません。');
        writeLogsInFormResponse(responseSheet, responseRow, '部活内を選んだ場合、日付は金曜日でなくてはいけません。');
        return;
      }
    }

    if(isSlotBooked(modifiedDate,modifiedTimeSlot)){
      console.log('変更後の日程はすでに埋まっています。');
      writeLogsInFormResponse(responseSheet, responseRow, '変更後の日程はすでに埋まっています。');
      return;
    }

    if(isMatchedRecently(applicantID,opponentID,modifiedDate,originalDate)){
      console.log('一部の例外を除き同じカードの対戦は' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。');
      writeLogsInFormResponse(responseSheet, responseRow, '一部の例外を除き同じカードの対戦は' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。');
      return;
    }

    let modifyFlag = false;
    for(let idx = 0;idx < matchData.length;idx++){
      const row = matchData[idx];
      if(row[APPLICANT_ID_COLUMN] === applicantID && row[OPPONENT_ID_COLUMN] === opponentID && (new Date(row[MATCH_DATE_COLUMN])).getTime() === originalDate.getTime()){
        if(timeSlot === '金曜部活内'){
          if(row[MATCH_TIMESLOT_COLUMN] === '部活時間外' || row[MATCH_TIMESLOT_COLUMN] === 'その他'){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            continue;
          }
        }else{
          if(row[MATCH_TIMESLOT_COLUMN] !== timeSlot){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            continue;
          }
        }

        if(row[MATCH_RESULT_FLAG_COLUMN] !== ''){
          console.log('該当の試合は存在しますが、すでに結果報告を受け取っているため変更はできません。');
          writeLogsInFormResponse(responseSheet, responseRow, '該当の試合は存在しますが、すでに結果報告を受け取っているため変更はできません。');
          continue;
        }

        if(row[MODIFY_FLAG_COLUMN] !== '可'){
          console.log('該当の試合は存在しますが、すでに一度日程/時間帯変更を行なっているため変更はできません。');
          writeLogsInFormResponse(responseSheet, responseRow, '該当の試合は存在しますが、すでに一度日程/時間帯変更を行なっているため変更はできません。');
          continue;
        }

        if(row[MATCH_TIMESLOT_COLUMN] !== '部活時間外' && row[MATCH_TIMESLOT_COLUMN] !== 'その他'){
          if(Number(row[MATCH_TIMESLOT_COLUMN][4]) !== FRIDAY_MATCH_NUMBER){
            narrowSchedule(originalDate,Number(row[MATCH_TIMESLOT_COLUMN][4]));
          }
        }

        rankMatchScheduleSheet.deleteRow(idx + HEADER_ROW_OFFSET + 1);
        markRankMatchDirty();
        manageChallenge(applicantID,false,isMale,responseSheet,responseRow);
        pushNewMatch(applicant,opponent,modifiedDate,modifiedTimeSlot,false,new Date(row[SCHEDULE_FORM_TIMESTAMP_COLUMN]),responseSheet,responseRow);
        console.log('該当の試合の日程/時間帯を変更しました。');
        modifyFlag = true;
        break;
      }
    }
    if(!modifyFlag){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse(responseSheet, responseRow, '対象の試合が見つかりませんでした。入力を再度確認してください。');
    }
  } catch (err) {
    console.log('日程/時間帯変更処理中にエラーが発生しました。' + err);
    writeLogsInFormResponse(responseSheet, responseRow, '日程/時間帯変更処理中にエラーが発生しました。' + err);
  }
}

//--------------------
// 日程追加操作
//--------------------
function processNormalRequest(applicant,opponent,originalDate,timeSlot,responseSheet,responseRow){
  try {
    const applicantID = applicant.id;
    const opponentID  = opponent .id;
    
    console.log(applicantID);
    console.log(opponentID);

    if(timeSlot === '金曜部活内'){
      if(!isFriday(originalDate)){
        console.log('部活内を選んだ場合、日付は金曜日でなくてはいけません。');
        writeLogsInFormResponse(responseSheet, responseRow, '部活内を選んだ場合、日付は金曜日でなくてはいけません。');
        return;
      }
    }

    if(isSlotBooked(originalDate,timeSlot)){
      console.log('その日程はすでに埋まっています。');
      writeLogsInFormResponse(responseSheet, responseRow, 'その日程はすでに埋まっています。');
      return;
    }

    let isMale = false;

    if(applicant.gender === '男'){
      isMale = true;
    }

    if(!canPlayMatch(applicantID,opponentID,isMale,responseSheet,responseRow)){
      console.log('この組み合わせの試合は行えません。');
      writeLogsInFormResponse(responseSheet, responseRow, 'この組み合わせの試合は行えません。');
      return;
    }

    pushNewMatch(applicant,opponent,originalDate,timeSlot,true,null,responseSheet,responseRow);
  } catch (err) {
    console.log('日程追加処理中にエラーが発生しました。' + err);
    writeLogsInFormResponse(responseSheet, responseRow, '日程追加処理中にエラーが発生しました。' + err);
  }
}

// その日が試合で埋まっているかを判定する関数
// 金曜日の強化練の時は判定するがそれ以外の場合は必ずfalseを返す
function isSlotBooked(date,slot){
  if(slot !== '金曜部活内')return false;

  const matchData = getRankMatchData();
  if(matchData.length === 0){
    return false;
  }

  let sameDayCount = 0;
  matchData.forEach((row) => {
    if((new Date(row[MATCH_DATE_COLUMN])).getTime() !== date.getTime())return;
    if(row[MATCH_TIMESLOT_COLUMN] === '部活時間外' || row[MATCH_TIMESLOT_COLUMN] === 'その他')return;
    sameDayCount += 1;
  })

  if(sameDayCount >= FRIDAY_MATCH_NUMBER){
    return true;
  }

  return false;
}

// 金曜日のマッチが何試合既に入っているかを数えて次が何試合目かを返す関数
function countFridayMatch(date){
  const matchData = getRankMatchData();
  if(matchData.length === 0){
    return 1;
  }

  let sameDayCount = 0;
  matchData.forEach((row) => {
    if((new Date(row[MATCH_DATE_COLUMN])).getTime() !== date.getTime())return;
    if(row[MATCH_TIMESLOT_COLUMN] === '部活時間外' || row[MATCH_TIMESLOT_COLUMN] === 'その他')return;
    sameDayCount += 1;
  })

  return sameDayCount + 1;
}

// 金曜の１、２試合目が無くなったらその後にあった試合を詰める関数
function narrowSchedule(date,matchNumber){
  const matchData = getRankMatchData();
  matchData.forEach((row,idx) => {
    if((new Date(row[MATCH_DATE_COLUMN])).getTime() !== date.getTime())return;
    if(row[MATCH_TIMESLOT_COLUMN] === '部活時間外' || row[MATCH_TIMESLOT_COLUMN] === 'その他')return;
    if(Number(row[MATCH_TIMESLOT_COLUMN][4]) > matchNumber){
      rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1 + idx,MATCH_TIMESLOT_COLUMN + 1).setValue('部活中(' + (Number(row[MATCH_TIMESLOT_COLUMN][4]) - 1) + '試合目)');
      markRankMatchDirty();
    }
    
  })
}

//この組み合わせの試合が可能か判定する関数
function canPlayMatch(applicantID,opponentID,isMale,responseSheet,responseRow){
  const rankData = isMale ? getMaleRankData() : getFemaleRankData();

  const applicantRowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === applicantID);
  const opponentRowIndex  = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === opponentID );
  
  if (applicantRowIndex === -1 || opponentRowIndex === -1) {
    console.log('プレイヤーが順位表から見つかりませんでした。');
    writeLogsInFormResponse(responseSheet, responseRow, 'プレイヤーが順位表から見つかりませんでした。');
    return false;
  }
  
  if(applicantRowIndex <= opponentRowIndex){
    console.log('申込者の順位の方が対戦相手よりも高いです。');
    writeLogsInFormResponse(responseSheet, responseRow, '申込者の順位の方が対戦相手よりも高いです。');
    return false;
  }

  if(Number(rankData[applicantRowIndex][MATCH_LIMIT_COLUMN]) <= 0){
    console.log('今月の挑戦権を使い切っています。');
    writeLogsInFormResponse(responseSheet, responseRow, '今月の挑戦権を使い切っています。');
    return false;
  }

  if(rankData[applicantRowIndex][CAN_PLAY_FLAG_COLUMN] !== '可' || rankData[opponentRowIndex][CAN_PLAY_FLAG_COLUMN] !== '可'){
    console.log('プレイヤーが試合できない状態です。');
    writeLogsInFormResponse(responseSheet, responseRow, 'プレイヤーが試合できない状態です。');
    return false;
  }

  let cannotPlayMatchPlayersCount = 0;

  for(let i = opponentRowIndex; i < applicantRowIndex; i++){
    if(rankData[i][CAN_PLAY_FLAG_COLUMN] !== '可')cannotPlayMatchPlayersCount += 1;
  }

  if(cannotPlayMatchPlayersCount <= 1){
    if(applicantRowIndex - opponentRowIndex <= MAX_RANK_DIFFERENCE){
      return true;
    }else{
      console.log('順位差が大きすぎです。');
      writeLogsInFormResponse(responseSheet, responseRow, '順位差が大きすぎです。');
      return false;
    }
  }else{
    if(applicantRowIndex - opponentRowIndex <= MAX_RANK_DIFFERENCE + cannotPlayMatchPlayersCount){
      return true;
    }else{
      console.log('順位差が大きすぎです。');
      writeLogsInFormResponse(responseSheet, responseRow, '順位差が大きすぎです。');
      return false;
    }
  }
}

// 新規日程を追加する関数
function pushNewMatch(applicant,opponent,date,slot,canUseModification,formSubmittedDate,responseSheet,responseRow){
  const applicantID = applicant.id;
  const opponentID  = opponent .id;
  const applicantName = applicant.name;
  const opponentName  = opponent .name;

  console.log(applicantID);
  console.log(opponentID);

  console.log(applicantName);
  console.log(opponentName);

  let isMale = false;

  if(applicant.gender === '男'){
    isMale = true;
  }

  if(isMatchedRecently(applicantID,opponentID,date,null)){
    console.log('一部の例外を除き同じカードの対戦は' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。');
    writeLogsInFormResponse(responseSheet, responseRow, '一部の例外を除き同じカードの対戦は' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。');
    return;
  }

  const modificationFlag = canUseModification ? '可' : '不可';

  const submitTime = formSubmittedDate ? formSubmittedDate : new Date();

  let slotString;

  if(slot === '部活時間外' || slot === 'その他'){
    slotString = slot;
  }else{
    const nextMatchNumber = countFridayMatch(date);
    slotString = '部活中(' + nextMatchNumber + '試合目)';
  }

  rankMatchScheduleSheet.appendRow([applicantID,applicantName,opponentID,opponentName,date,slotString,'','','','','',modificationFlag,submitTime,'']);
  markRankMatchDirty();
  manageChallenge(applicantID,true,isMale,responseSheet,responseRow);
}

// 挑戦権を管理する関数
// isUsedがtrueなら減らし、falseなら増やす
function manageChallenge(id,isUsed,isMale,responseSheet,responseRow){
  const rankData = isMale ? getMaleRankData() : getFemaleRankData();

  const rowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === id);
  
  if(rowIndex === -1){
    console.log('プレイヤーが順位表から見つかりませんでした。');
    writeLogsInFormResponse(responseSheet, responseRow, 'プレイヤーが順位表から見つかりませんでした。');

    return;
  }

  if(isMale){
    const nextValue = isUsed
      ? Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) - 1
      : Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) + 1;
    maleSheet.getRange(rowIndex + HEADER_ROW_OFFSET + 1,MATCH_LIMIT_COLUMN + 1).setValue(nextValue);
    rankData[rowIndex][MATCH_LIMIT_COLUMN] = nextValue;
  }else{
    const nextValue = isUsed
      ? Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) - 1
      : Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) + 1;
    femaleSheet.getRange(rowIndex + HEADER_ROW_OFFSET + 1,MATCH_LIMIT_COLUMN + 1).setValue(nextValue);
    rankData[rowIndex][MATCH_LIMIT_COLUMN] = nextValue;
  }
}

// クールダウン期間に引っかかっていないかを判定する関数
function isMatchedRecently(applicantID,opponentID,date,exceptionDate){
  const matchData = getRankMatchData();
  if(matchData.length === 0){
    return false;
  }

  const scopeBefore = new Date();
  scopeBefore.setHours(0,0,0,0);
  scopeBefore.setDate(date.getDate() - SAME_OPPONENT_COOLDOWN_DAYS);

  const scopeAfter = new Date();
  scopeAfter.setHours(0,0,0,0);
  scopeAfter.setDate(date.getDate() + SAME_OPPONENT_COOLDOWN_DAYS);

  const now = new Date();
  now.setHours(0,0,0,0);

  let isMatchedFlag = false;

  matchData.forEach((row) => {
    if(!(row[APPLICANT_ID_COLUMN] === opponentID && row[OPPONENT_ID_COLUMN] === applicantID))return;
    if(new Date(row[MATCH_DATE_COLUMN]).getTime() >= scopeBefore.getTime() && new Date(row[MATCH_DATE_COLUMN]).getTime() <= scopeAfter.getTime() && row[MATCH_RESULT_FLAG_COLUMN] !== '' && row[MATCH_RESULT_COLUMN] !== '敗北')isMatchedFlag = true;
    if(new Date(row[MATCH_DATE_COLUMN]).getTime() >= now.getTime() && new Date(row[MATCH_DATE_COLUMN]).getTime() <= scopeAfter.getTime() && row[MATCH_RESULT_FLAG_COLUMN] === '')isMatchedFlag = true;
  })

  matchData.forEach((row) => {
    if(!(row[APPLICANT_ID_COLUMN] === applicantID && row[OPPONENT_ID_COLUMN] === opponentID))return;
    if(exceptionDate && new Date(row[MATCH_DATE_COLUMN]).getTime() === exceptionDate.getTime())return;
    if(new Date(row[MATCH_DATE_COLUMN]).getTime() >= scopeBefore.getTime() && new Date(row[MATCH_DATE_COLUMN]).getTime() <= scopeAfter.getTime() && row[MATCH_RESULT_COLUMN] === '敗北')isMatchedFlag = true;
    if(new Date(row[MATCH_DATE_COLUMN]).getTime() >= now.getTime() && new Date(row[MATCH_DATE_COLUMN]).getTime() <= scopeAfter.getTime() && row[MATCH_RESULT_FLAG_COLUMN] === '')isMatchedFlag = true;
  })

  return isMatchedFlag;
}

// 金曜日かどうか判定する関数
// 金曜日じゃなくなる場合は数字を変更してください。(日=0,月=1,火=2,水=3,...,土=6)
function isFriday(date){
  return date.getDay() === 5;
}

// 以上日程報告フォームの処理
// -----------------------------------------------------
// 以下結果報告フォームの処理

// 試合日程に結果を書き込む関数
function writeMatchResult(applicant,opponent,matchResult,game1Score,game2Score,game3Score,responseSheet,responseRow){
  const applicantID = applicant.id;
  const opponentID  = opponent .id;

  console.log(applicantID);
  console.log(opponentID);

  const matchData = getRankMatchData();
  if(matchData.length === 0){
    console.log('対象の試合が見つかりませんでした。この結果報告は無効です。');
    writeLogsInFormResponse(responseSheet, responseRow, '対象の試合が見つかりませんでした。この結果報告は無効です。');
    return;
  }

  let isMale = false;

  if(applicant.gender === '男'){
    isMale = true;
  }

  const today = new Date();
  today.setHours(0,0,0,0);

  let isExecuted = false;

  let applytime;

  matchData.forEach((row,idx) => {
    if(isExecuted)return;
    if(row[APPLICANT_ID_COLUMN] === applicantID && row[OPPONENT_ID_COLUMN] === opponentID && row[MATCH_RESULT_FLAG_COLUMN] === ''){

      if((new Date(row[MATCH_DATE_COLUMN])).getTime() !== today.getTime()){
        console.log('同カードで結果未提出の試合がありますが、日付が今日のものではありません。');
        return;
      }

      rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1 + idx,MATCH_RESULT_FLAG_COLUMN + 1,1,5).setValues([['済',matchResult,game1Score,game2Score,game3Score]]);
      rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1 + idx,RESULT_FORM_TIMESTAMP_COLUMN + 1).setValue(new Date());
      markRankMatchDirty();

      applytime = new Date(row[SCHEDULE_FORM_TIMESTAMP_COLUMN]);

      isExecuted = true;
    }
  })

  if(!isExecuted){
    console.log('対象の試合が見つかりませんでした。この結果報告は無効です。');
    writeLogsInFormResponse(responseSheet, responseRow, '対象の試合が見つかりませんでした。この結果報告は無効です。');
    return;
  }

  console.log('試合結果を正常に書き込みました。');

  if(matchResult === '敗北'){
    removeWinningBonus(applicantID,isMale);
    return;
  }

  console.log('順位表の変動を開始します。');

  changeRanking(applicantID,opponentID,isMale,applytime,responseSheet,responseRow);
}

// ランキングを変更する関数
function changeRanking(applicantID,opponentID,isMale,applytime,responseSheet,responseRow){
  const rankData = isMale ? getMaleRankData() : getFemaleRankData();

  const applicantRowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === applicantID);
  let opponentRowIndex  = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === opponentID );

  if (applicantRowIndex === -1 || opponentRowIndex === -1) {
    console.log('プレイヤーが順位表から見つかりませんでした。');
    writeLogsInFormResponse(responseSheet, responseRow, 'プレイヤーが順位表から見つかりませんでした。');
    return;
  }
  
  if(applicantRowIndex <= opponentRowIndex){
    console.log('申込者の順位の方が対戦相手よりも高いです。');
    // writeLogsInFormResponse(responseSheet, responseRow, '申込者の順位の方が対戦相手よりも高いです。');
    return;
  }

  if(isMale){
    if(Number(rankData[applicantRowIndex][MATCH_LIMIT_COLUMN]) === 0 && rankData[applicantRowIndex][MATCH_MORE_FLAG_COLUMN] === '可'){
      console.log('連勝ボーナスを発動します。');
      rankData[applicantRowIndex][MATCH_LIMIT_COLUMN] = '1';
    }
  }else{
    if(Number(rankData[applicantRowIndex][MATCH_LIMIT_COLUMN]) === 0 && rankData[applicantRowIndex][MATCH_MORE_FLAG_COLUMN] === '可'){
      console.log('連勝ボーナスを発動します。');
      rankData[applicantRowIndex][MATCH_LIMIT_COLUMN] = '1';
    }
  }

  const isPromoted = isOpponentPromoted(applytime,opponentID);
  if(isPromoted){
    opponentRowIndex = applicantRowIndex - Math.min(applicantRowIndex - opponentRowIndex,MAX_RANK_DIFFERENCE);
  }

  if(isMale){
    const subRanking = rankData.slice(opponentRowIndex,applicantRowIndex + 1);
    const removedRankSubRanking = subRanking.map((row) => row.slice(1));
    maleSheet.getRange(HEADER_ROW_OFFSET + 1 + 1 + opponentRowIndex,2,applicantRowIndex - opponentRowIndex,RANKING_SHEET_MAX_COLUMN-1).setValues(removedRankSubRanking.slice(0,-1));
    maleSheet.getRange(HEADER_ROW_OFFSET + 1 + opponentRowIndex,2,1,RANKING_SHEET_MAX_COLUMN-1).setValues(removedRankSubRanking.slice(-1));
    markMaleRankDirty();
  }else{
    const subRanking = rankData.slice(opponentRowIndex,applicantRowIndex + 1);
    const removedRankSubRanking = subRanking.map((row) => row.slice(1));
    femaleSheet.getRange(HEADER_ROW_OFFSET + 1 + 1 + opponentRowIndex,2,applicantRowIndex - opponentRowIndex,RANKING_SHEET_MAX_COLUMN-1).setValues(removedRankSubRanking.slice(0,-1));
    femaleSheet.getRange(HEADER_ROW_OFFSET + 1 + opponentRowIndex,2,1,RANKING_SHEET_MAX_COLUMN-1).setValues(removedRankSubRanking.slice(-1));
    markFemaleRankDirty();
  }

  console.log('ランキングを正常に変更しました。');

  buildRankingsByDepartment();
}

// 連勝ボーナスを剥奪する関数
function removeWinningBonus(applicantID,isMale){
  const rankData = isMale ? getMaleRankData() : getFemaleRankData();

  const applicantRowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === applicantID);

  if(isMale){
    maleSheet.getRange(HEADER_ROW_OFFSET + 1 + applicantRowIndex,MATCH_MORE_FLAG_COLUMN + 1).setValue('不可');
    rankData[applicantRowIndex][MATCH_MORE_FLAG_COLUMN] = '不可';
  }else{
    femaleSheet.getRange(HEADER_ROW_OFFSET + 1 + applicantRowIndex,MATCH_MORE_FLAG_COLUMN + 1).setValue('不可');
    rankData[applicantRowIndex][MATCH_MORE_FLAG_COLUMN] = '不可';
  }
  return;
}

// 試合申し込みから対戦相手が勝ち上がったかどうかを判定する関数
function isOpponentPromoted(time,opponentID){
  const matchData = getRankMatchData();
  let isPromoted = false;
  matchData.forEach((row) => {
    if(row[APPLICANT_ID_COLUMN] !== opponentID)return;
    if(row[MATCH_RESULT_FLAG_COLUMN] === '')return;
    if(row[MATCH_RESULT_COLUMN] === '敗北')return;
    if(new Date(row[RESULT_FORM_TIMESTAMP_COLUMN]).getTime() < time.getTime())return;

    isPromoted = true;
  })

  return isPromoted;
}

// 以上結果報告の関数
// -----------------------------------------
// 以下その他の関数

// 性別、名前、IDを返す関数
function parsePlayerLabel(label) {
  const text = String(label || '').trim();
  const m = text.match(/^\((男|女)\)\s*(.+?)\s*\(([^()]+)\)$/);
  if (!m) {
    throw new Error(`プレイヤー表記が不正です: ${text}`);
  }
  return { gender: m[1], name: m[2].trim(), id: m[3].trim() };
}

//時系列順にソートする関数
function sortRankMatchSchedule(){
  const matchData = getRankMatchData();
  if(matchData.length === 0){
    return;
  }
  matchData.sort((a, b) => {
    if(new Date(a[MATCH_DATE_COLUMN]).getTime() - new Date(b[MATCH_DATE_COLUMN]).getTime() !== 0){
      return new Date(a[MATCH_DATE_COLUMN]).getTime() - new Date(b[MATCH_DATE_COLUMN]).getTime();
    }

    let a_order,b_order;
    if(a[MATCH_TIMESLOT_COLUMN] in timeSlotSortOrder === false){
      a_order = 1;
    }else{
      a_order = timeSlotSortOrder[a[MATCH_TIMESLOT_COLUMN]];
    }
    if(b[MATCH_TIMESLOT_COLUMN] in timeSlotSortOrder === false){
      b_order = 1;
    }else{
      b_order = timeSlotSortOrder[b[MATCH_TIMESLOT_COLUMN]];
    }

    if(a_order !== b_order || a_order !== 1)return a_order - b_order;
    else{
      const na = Number(a[MATCH_TIMESLOT_COLUMN].match(/\((\d+)試合目\)/)?.[1] ?? 0);
      const nb = Number(b[MATCH_TIMESLOT_COLUMN].match(/\((\d+)試合目\)/)?.[1] ?? 0);
      return na - nb;
    }
  });
  rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,matchData.length,RANK_MATCH_SHEET_MAX_COLUMN).setValues(matchData);
  markRankMatchDirty();

  console.log('日程を時系列順にソートしました。');
}

//フォームのプルダウンを変更する関数
function updateFormDropdown() {
  const choices = [];

  const maleData = getMaleRankData();
  for (let i = 0; i < maleData.length; i++) {
    const studentID = maleData[i][PLAYER_ID_COLUMN];
    const studentName = maleData[i][PLAYER_NAME_COLUMN];
    choices.push('(男) ' + studentName + ' (' + studentID + ')');
  }

  const femaleData = getFemaleRankData();
  for (let i = 0; i < femaleData.length; i++) {
    const studentID = femaleData[i][PLAYER_ID_COLUMN];
    const studentName = femaleData[i][PLAYER_NAME_COLUMN];
    choices.push('(女) ' + studentName + ' (' + studentID + ')');
  }
  
  const matchSchedulingForm = FormApp.openById(MATCH_SCHEDULING_FORM_ID);
  const matchResultForm = FormApp.openById(MATCH_RESULT_FORM_ID);
  const matchResultCheckForm = FormApp.openById(MATCH_RESULT_CHECK_FORM_ID);

  const matchSchedulingItems = matchSchedulingForm.getItems();
  const matchResultItems = matchResultForm.getItems();
  const matchResultCheckItems = matchResultCheckForm.getItems();

  for (let j = 0; j < matchSchedulingItems.length; j++) {
    var item = matchSchedulingItems[j];
    if(item.getTitle() === 'ランク戦を申し込む人' || item.getTitle() === 'ランク戦を受ける人'){
      const itemQuestion = item.asListItem();
      itemQuestion.setChoiceValues(choices);
    }
  }
  for (let j = 0; j < matchResultItems.length; j++) {
    var item = matchResultItems[j];
    if(item.getTitle() === 'ランク戦を申し込んだ人' || item.getTitle() === 'ランク戦を受けた人'){
      const itemQuestion = item.asListItem();
      itemQuestion.setChoiceValues(choices);
    }
  }
  for (let j = 0; j < matchResultCheckItems.length; j++) {
    var item = matchResultCheckItems[j];
    if(item.getTitle() === '誰に関する結果を確認するか' || item.getTitle() === '（任意）対戦相手'){
      const itemQuestion = item.asListItem();
      itemQuestion.setChoiceValues(choices);
    }
  }

  console.log('フォームの選択肢の順番を変更しました。');
}

// 月初めに挑戦権を回復させる関数(定期実行)
function restoreChallengeRight(){
  const now = new Date();
  const month = now.getMonth(); 
  const cell = MONTH_APPLICATION_LIMIT_CELL[0] + String(Number(MONTH_APPLICATION_LIMIT_CELL[1]) + month + 1);
  const monthApplicationLimit = configSheet.getRange(cell).getValue();
  maleSheet.getRange('F2:F' + maleSheet.getLastRow()).setValue(monthApplicationLimit);
  maleSheet.getRange('G2:G' + maleSheet.getLastRow()).setValue('可');
  femaleSheet.getRange('F2:F' + femaleSheet.getLastRow()).setValue(monthApplicationLimit);
  femaleSheet.getRange('G2:G' + femaleSheet.getLastRow()).setValue('可');
  markMaleRankDirty();
  markFemaleRankDirty();
  console.log('挑戦権を回復させました。');
}

// フォームに対し、処理のログを書き込む関数
function writeLogsInFormResponse(responseSheet, responseRow, message){
  const sheetName = responseSheet.getName();

  if(sheetName === '日程報告'){
    const previousMessage = scheduleFormResponsesSheet.getRange(responseRow,9).getValue();
    scheduleFormResponsesSheet.getRange(responseRow,9).setValue(previousMessage + message);
  }else{
    const previousMessage = resultFormResponsesSheet.getRange(responseRow,8).getValue();
    resultFormResponsesSheet.getRange(responseRow,8).setValue(previousMessage + message);
  }
}

// 学科別順位表を構築する関数
function buildRankingsByDepartment(){
  const maleData = getMaleRankData().map((row) => row.slice(0, 4));
  const femaleData = getFemaleRankData().map((row) => row.slice(0, 4));

  let maleMedicineData = [];
  let maleInsuranceData = [];

  for(let i = 0;i < maleData.length;i++){
    let maleRow = maleData[i];
    if(maleData[i][PLAYER_ID_COLUMN][4] === '1'){
      maleRow[0] = String(maleMedicineData.length + 1) + ' (' + maleRow[0] + ')';
      maleMedicineData.push(maleRow);
    }else{
      maleRow[0] = String(maleInsuranceData.length + 1) + ' (' + maleRow[0] + ')';
      maleInsuranceData.push(maleRow);
    }
  }

  let femaleMedicineData = [];
  let femaleInsuranceData = [];

  for(let i = 0;i < femaleData.length;i++){
    let femaleRow = femaleData[i];
    if(femaleData[i][PLAYER_ID_COLUMN][4] === '1'){
      femaleRow[0] = String(femaleMedicineData.length + 1) + ' (' + femaleRow[0] + ')';
      femaleMedicineData.push(femaleRow);
    }else{
      femaleRow[0] = String(femaleInsuranceData.length + 1) + ' (' + femaleRow[0] + ')';
      femaleInsuranceData.push(femaleRow);
    }
  }

  maleSheetByDepartment.getRange(2,1,maleSheetByDepartment.getLastRow(),maleSheetByDepartment.getLastColumn()).clearContent();
  femaleSheetByDepartment.getRange(2,1,femaleSheetByDepartment.getLastRow(),femaleSheetByDepartment.getLastColumn()).clearContent();
  
  maleSheetByDepartment  .getRange(HEADER_ROW_OFFSET + 1,2,maleMedicineData   .length,4).setValues(maleMedicineData   );
  maleSheetByDepartment  .getRange(HEADER_ROW_OFFSET + 1,8,maleInsuranceData  .length,4).setValues(maleInsuranceData  );
  femaleSheetByDepartment.getRange(HEADER_ROW_OFFSET + 1,2,femaleMedicineData .length,4).setValues(femaleMedicineData );
  femaleSheetByDepartment.getRange(HEADER_ROW_OFFSET + 1,8,femaleInsuranceData.length,4).setValues(femaleInsuranceData);
}