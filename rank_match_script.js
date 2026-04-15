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

// 日程シートの列の定数
const APPLICANT_ID_COLUMN = 0;
const APPLICANT_NAME_COLUMN = 1;
const OPPONENT_ID_COLUMN = 2;
const OPPONENT_NAME_COLUMN = 3;
const MATCH_DATE_COLUMN = 4;
const MATCH_TIMESLOT_COLUMN = 5;
const MATCH_RESULT_FLAG_COLUMN = 6;
const MATCH_RESULT_COLUMN = 7;
const MODIFY_FLAG_COLUMN = 8;
const SCHEDULE_FORM_TIMESTAMP_COLUMN = 9;
const RESULT_FORM_TIMESTAMP_COLUMN = 10;

// ランキングシートの列の定数
const PLAYER_ID_COLUMN = 1;
const PLAYER_NAME_COLUMN = 2;
const CAN_PLAY_FLAG_COLUMN = 4;
const MATCH_LIMIT_COLUMN = 5;
const MATCH_MORE_FLAG_COLUMN = 6;

// フォームを受け取った時の分岐
// 日程報告か結果報告か
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    console.log('他の処理が終わっていないためスキップします。');
    return;
  }
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();

    if(sheetName === '日程報告'){
      handleSchedule(e);
    }else if(sheetName === '結果報告'){
      handleResult(e);
    }
  } finally {
    lock.releaseLock();
  }
}

// 日程報告を受け取った時の関数
function handleSchedule(e){
  try {
    const formData = e.values;
    console.log("新規日程報告受信: " + JSON.stringify(formData));
    const timestamp = formData[0];
    const applicant = formData[1];
    const opponent = formData[2];
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

    if(applicant === opponent){
      console.log('対戦する人が同一人物です。入力は無効です。');
      writeLogsInFormResponse('対戦する人が同一人物です。入力は無効です。',true);
      return;
    }

    if(applicant[1] !== opponent[1]){
      console.log('対戦する人の性別が違います。入力は無効です。');
      writeLogsInFormResponse('対戦する人の性別が違います。入力は無効です。',true);
      return;
    }

    if(originalDate < today){
      console.log('日付が過去のものです。入力は無効です。');
      writeLogsInFormResponse('日付が過去のものです。入力は無効です。',true);
      return;
    }

    if(applicationScope < originalDate){
      console.log('日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。');
      writeLogsInFormResponse('日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。',true);
      return;
    }

    if(modifiedDate && applicationScope < modifiedDate){
      console.log('日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。');
      writeLogsInFormResponse('日付が未来すぎます。' + MATCH_ACCEPT_DAY_LIMIT + '日以内の日程のみ許可します。入力は無効です。',true);
      return;
    }

    if(cancelFlag === 'キャンセル'){
      console.log('キャンセル操作を実行します。');
      processCancelRequest(applicant,opponent,originalDate,timeSlot);
      SpreadsheetApp.flush();
      return;
    }

    if(modifiedDate && modifiedTimeSlot){
      if(originalDate === modifiedDate && timeSlot === modifiedTimeSlot){
        console.log('変更前と変更後の日付/時間帯が同じです。');
        return;
      }
      console.log('日付/時間帯変更操作を実行します。');
      processModifyRequest(applicant,opponent,originalDate,timeSlot,modifiedDate,modifiedTimeSlot);
      sortRankMatchSchedule();
      SpreadsheetApp.flush();
      return;
    }

    if(modifiedDate || modifiedTimeSlot){
      console.log('日程追加の場合はどちらも空白でなければなりません。また、日付/時間帯変更の場合は日付と時間のどちらの入力も必要です。入力は無効です。');
      writeLogsInFormResponse('日程追加の場合はどちらも空白でなければなりません。また、日付/時間帯変更の場合は日付と時間のどちらの入力も必要です。入力は無効です。',true);
      return;
    }

    console.log('日程追加操作を実行します。');
    processNormalRequest(applicant,opponent,originalDate,timeSlot);
    sortRankMatchSchedule();
    SpreadsheetApp.flush();

  } catch (err) {
    console.log('日程報告中に予期せぬエラーが発生しました。' + err);
    writeLogsInFormResponse('日程報告中に予期せぬエラーが発生しました。' + err,true);
  }
}

// 結果報告を受け取った時の関数
function handleResult(e){
  try {
    const formData = e.values;
    console.log("結果報告受信: " + JSON.stringify(formData));
    const timestamp = formData[0];
    const applicant = formData[1];
    const opponent = formData[2];
    const matchResult = formData[3];
    const game1Score = formData[4] ? formData[4] : null;
    const game2Score = formData[5] ? formData[5] : null;
    const game3Score = formData[6] ? formData[6] : null;

    if(applicant === opponent){
      console.log('対戦する人が同一人物です。入力は無効です。');
      writeLogsInFormResponse('対戦する人が同一人物です。入力は無効です。',false);
      return;
    }

    if(applicant[1] !== opponent[1]){
      console.log('対戦する人の性別が違います。入力は無効です。');
      writeLogsInFormResponse('対戦する人の性別が違います。入力は無効です。',false);
      return;
    }

    console.log('結果報告の書き込みを開始します。');
    writeMatchResult(applicant,opponent,matchResult,game1Score,game2Score,game3Score);
    updateFormDropdown();
    SpreadsheetApp.flush();
  } catch (err) {
    console.log('結果報告中に予期せぬエラーが発生しました。' + err);
    writeLogsInFormResponse('結果報告中に予期せぬエラーが発生しました。' + err,false);
  }
}

// 以下日程報告フォームの処理

//--------------------
// キャンセル操作
//--------------------
function processCancelRequest(applicant,opponent,originalDate,timeSlot){
  try {
    const applicantID = applicant.substring(applicant.length-9,applicant.length-1);
    const opponentID  = opponent .substring(opponent .length-9,opponent .length-1);

    console.log(applicantID);
    console.log(opponentID);

    const lastRow = rankMatchScheduleSheet.getLastRow();
    if(lastRow <= HEADER_ROW_OFFSET){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse('対象の試合が見つかりませんでした。入力を再度確認してください。',true);
      return;
    }

    let isMale = false;

    if(applicant[1] === '男'){
      isMale = true;
    }

    const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
    let deleteFlag = false;
    matchData.forEach((row,idx) => {
      if(row[APPLICANT_ID_COLUMN] === applicantID && row[OPPONENT_ID_COLUMN] === opponentID && (new Date(row[MATCH_DATE_COLUMN])).getTime() === originalDate.getTime()){
        
        if(timeSlot === '金曜部活内'){
          if(row[MATCH_TIMESLOT_COLUMN] === '部活時間外' || row[MATCH_TIMESLOT_COLUMN] === 'その他'){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }else{
          if(row[MATCH_TIMESLOT_COLUMN] !== timeSlot){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }

        if(row[MATCH_RESULT_FLAG_COLUMN] !== ''){
          console.log('該当の試合は存在しますが、すでに結果報告を受け取っているためキャンセルはできません。');
          writeLogsInFormResponse('該当の試合は存在しますが、すでに結果報告を受け取っているためキャンセルはできません。',true);
          return;
        }

        if(row[MATCH_TIMESLOT_COLUMN] !== '部活時間外' && row[MATCH_TIMESLOT_COLUMN] !== 'その他'){
          if(Number(row[MATCH_TIMESLOT_COLUMN][4]) !== FRIDAY_MATCH_NUMBER){
            narrowSchedule(originalDate,Number(row[MATCH_TIMESLOT_COLUMN][4]));
          }
        }

        rankMatchScheduleSheet.deleteRow(idx + HEADER_ROW_OFFSET + 1);
        manageChallenge(applicantID,false,isMale);
        console.log('該当の試合を削除しました。');
        deleteFlag = true;
      }
    })
    if(!deleteFlag){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse('対象の試合が見つかりませんでした。入力を再度確認してください。',true);
    }

  } catch (err) {
    console.log('キャンセル処理中にエラーが発生しました。' + err);
    writeLogsInFormResponse('キャンセル処理中にエラーが発生しました。' + err,true);
  }
}

//--------------------
// 日程/時間帯変更操作
//--------------------
function processModifyRequest(applicant,opponent,originalDate,timeSlot,modifiedDate,modifiedTimeSlot){
  try {
    const applicantID = applicant.substring(applicant.length-9,applicant.length-1);
    const opponentID  = opponent .substring(opponent .length-9,opponent .length-1);

    console.log(applicantID);
    console.log(opponentID);

    const lastRow = rankMatchScheduleSheet.getLastRow();
    if(lastRow <= HEADER_ROW_OFFSET){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse('対象の試合が見つかりませんでした。入力を再度確認してください。',true);
      return;
    }

    let isMale = false;

    if(applicant[1] === '男'){
      isMale = true;
    }

    if(modifiedTimeSlot === '金曜部活内'){
      if(!isFriday(modifiedDate)){
        console.log('部活内を選んだ場合、日付は金曜日でなくてはいけません。');
        writeLogsInFormResponse('部活内を選んだ場合、日付は金曜日でなくてはいけません。',true);
        return;
      }
    }

    if(isSlotBooked(modifiedDate,modifiedTimeSlot)){
      console.log('変更後の日程はすでに埋まっています。');
      writeLogsInFormResponse('変更後の日程はすでに埋まっています。',true);
      return;
    }

    if(isMatchedRecently(applicantID,opponentID,modifiedDate)){
      console.log('一部の例外を除き同じカードの対戦は前回の対戦から' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。');
      writeLogsInFormResponse('一部の例外を除き同じカードの対戦は前回の対戦から' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。',true);
      return;
    }

    const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
    let modifyFlag = false;
    matchData.forEach((row,idx) => {
      if(row[APPLICANT_ID_COLUMN] === applicantID && row[OPPONENT_ID_COLUMN] === opponentID && (new Date(row[MATCH_DATE_COLUMN])).getTime() === originalDate.getTime()){
        
        if(timeSlot === '金曜部活内'){
          if(row[MATCH_TIMESLOT_COLUMN] === '部活時間外' || row[MATCH_TIMESLOT_COLUMN] === 'その他'){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }else{
          if(row[MATCH_TIMESLOT_COLUMN] !== timeSlot){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }

        if(row[MATCH_RESULT_FLAG_COLUMN] !== ''){
          console.log('該当の試合は存在しますが、すでに結果報告を受け取っているため変更はできません。');
          writeLogsInFormResponse('該当の試合は存在しますが、すでに結果報告を受け取っているため変更はできません。',true);
          return;
        }

        if(row[MODIFY_FLAG_COLUMN] !== '可'){
          console.log('該当の試合は存在しますが、すでに一度日程/時間帯変更を行なっているため変更はできません。');
          writeLogsInFormResponse('該当の試合は存在しますが、すでに一度日程/時間帯変更を行なっているため変更はできません。',true);
          return;
        }

        if(row[MATCH_TIMESLOT_COLUMN] !== '部活時間外' && row[MATCH_TIMESLOT_COLUMN] !== 'その他'){
          if(Number(row[MATCH_TIMESLOT_COLUMN][4]) !== FRIDAY_MATCH_NUMBER){
            narrowSchedule(originalDate,Number(row[MATCH_TIMESLOT_COLUMN][4]));
          }
        }

        rankMatchScheduleSheet.deleteRow(idx + HEADER_ROW_OFFSET + 1);
        manageChallenge(applicantID,false,isMale);
        pushNewMatch(applicant,opponent,modifiedDate,modifiedTimeSlot,false);
        console.log('該当の試合の日程/時間帯を変更しました。');
        modifyFlag = true;
      }
    })
    if(!modifyFlag){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      writeLogsInFormResponse('対象の試合が見つかりませんでした。入力を再度確認してください。',true);
    }
  } catch (err) {
    console.log('日程/時間帯変更処理中にエラーが発生しました。' + err);
    writeLogsInFormResponse('日程/時間帯変更処理中にエラーが発生しました。' + err,true);
  }
}

//--------------------
// 日程追加操作
//--------------------
function processNormalRequest(applicant,opponent,originalDate,timeSlot){
  try {
    const applicantID = applicant.substring(applicant.length-9,applicant.length-1);
    const opponentID  = opponent .substring(opponent .length-9,opponent .length-1);
    
    console.log(applicantID);
    console.log(opponentID);

    if(timeSlot === '金曜部活内'){
      if(!isFriday(originalDate)){
        console.log('部活内を選んだ場合、日付は金曜日でなくてはいけません。');
        writeLogsInFormResponse('部活内を選んだ場合、日付は金曜日でなくてはいけません。',true);
        return;
      }
    }

    if(isSlotBooked(originalDate,timeSlot)){
      console.log('その日程はすでに埋まっています。');
      writeLogsInFormResponse('その日程はすでに埋まっています。',true);
      return;
    }

    let isMale = false;

    if(applicant[1] === '男'){
      isMale = true;
    }

    if(!canPlayMatch(applicantID,opponentID,isMale)){
      console.log('この組み合わせの試合は行えません。');
      writeLogsInFormResponse('この組み合わせの試合は行えません。',true);
      return;
    }

    pushNewMatch(applicant,opponent,originalDate,timeSlot,true);
  } catch (err) {
    console.log('日程追加処理中にエラーが発生しました。' + err);
    writeLogsInFormResponse('日程追加処理中にエラーが発生しました。' + err,true);
  }
}

// その日が試合で埋まっているかを判定する関数
// 金曜日の強化練の時は判定するがそれ以外の場合は必ずfalseを返す
function isSlotBooked(date,slot){
  if(slot !== '金曜部活内')return false;

  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= HEADER_ROW_OFFSET){
    return false;
  }

  const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
  let sameDayCount = 0;
  matchData.forEach((row) => {
    if((new Date(row[MATCH_DATE_COLUMN])).getTime() === date.getTime()){
      if(row[MATCH_TIMESLOT_COLUMN] !== '部活時間外' && row[MATCH_TIMESLOT_COLUMN] !== 'その他'){
        sameDayCount += 1;
      }
    }
  })

  if(sameDayCount >= FRIDAY_MATCH_NUMBER){
    return true;
  }

  return false;
}

// 金曜日のマッチが何試合既に入っているかを数えて次が何試合目かを返す関数
function countFridayMatch(date){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= HEADER_ROW_OFFSET){
    return 1;
  }

  const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
  let sameDayCount = 0;
  matchData.forEach((row) => {
    if((new Date(row[MATCH_DATE_COLUMN])).getTime() === date.getTime()){
      if(row[MATCH_TIMESLOT_COLUMN] !== '部活時間外' && row[MATCH_TIMESLOT_COLUMN] !== 'その他'){
        sameDayCount += 1;
      }
    }
  })

  return sameDayCount + 1;
}

// 金曜の１、２試合目が無くなったらその後にあった試合を詰める関数
function narrowSchedule(date,matchNumber){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
  matchData.forEach((row,idx) => {
    if((new Date(row[MATCH_DATE_COLUMN])).getTime() === date.getTime()){
      if(row[MATCH_TIMESLOT_COLUMN] !== '部活時間外' && row[MATCH_TIMESLOT_COLUMN] !== 'その他'){
        if(Number(row[MATCH_TIMESLOT_COLUMN][4]) > matchNumber){
          rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1 + idx,6).setValue('部活中(' + (Number(row[MATCH_TIMESLOT_COLUMN][4]) - 1) + '試合目)');
        }
      }
    }
  })
}

//この組み合わせの試合が可能か判定する関数
function canPlayMatch(applicantID,opponentID,isMale){

  let lastRow,rankData;
  
  if(isMale){
    lastRow = maleSheet.getLastRow();
    rankData = maleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();

  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();
  }

  const applicantRowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === applicantID);
  const opponentRowIndex  = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === opponentID );
  
  if (applicantRowIndex === -1 || opponentRowIndex === -1) {
    console.log('プレイヤーが順位表から見つかりませんでした。');
    writeLogsInFormResponse('プレイヤーが順位表から見つかりませんでした。',true);
    return false;
  }
  
  if(applicantRowIndex <= opponentRowIndex){
    console.log('申込者の順位の方が対戦相手よりも高いです。');
    writeLogsInFormResponse('申込者の順位の方が対戦相手よりも高いです。',true);
    return false;
  }

  if(Number(rankData[applicantRowIndex][MATCH_LIMIT_COLUMN]) <= 0){
    console.log('今月の挑戦権を使い切っています。');
    writeLogsInFormResponse('今月の挑戦権を使い切っています。',true);
    return false;
  }

  if(rankData[applicantRowIndex][CAN_PLAY_FLAG_COLUMN] !== '可' || rankData[opponentRowIndex][CAN_PLAY_FLAG_COLUMN] !== '可'){
    console.log('プレイヤーが試合できない状態です。');
    writeLogsInFormResponse('プレイヤーが試合できない状態です。',true);
    return false;
  }

  let cannotPlayMatchPlayersCount = 0;

  for(let i = opponentRowIndex; i < applicantRowIndex; i++){
    if(rankData[opponentRowIndex][CAN_PLAY_FLAG_COLUMN] !== '可')cannotPlayMatchPlayersCount += 1;
  }

  if(cannotPlayMatchPlayersCount <= 1){
    if(applicantRowIndex - opponentRowIndex <= MAX_RANK_DIFFERENCE){
      return true;
    }else{
      console.log('順位差が大きすぎです。');
      writeLogsInFormResponse('順位差が大きすぎです。',true);
      return false;
    }
  }else{
    if(applicantRowIndex - opponentRowIndex <= MAX_RANK_DIFFERENCE + cannotPlayMatchPlayersCount){
      return true;
    }else{
      console.log('順位差が大きすぎです。');
      writeLogsInFormResponse('順位差が大きすぎです。',true);
      return false;
    }
  }
}

// 新規日程を追加する関数
function pushNewMatch(applicant,opponent,date,slot,canUseModification){
  const applicantID = applicant.substring(applicant.length-9,applicant.length-1);
  const opponentID  = opponent .substring(opponent .length-9,opponent .length-1);
  const applicantName = applicant.substring(3,applicant.length-11);
  const opponentName  = opponent .substring(3,opponent .length-11);

  console.log(applicantID);
  console.log(opponentID);

  console.log(applicantName);
  console.log(opponentName);

  let isMale = false;

  if(applicant[1] === '男'){
    isMale = true;
  }

  if(isMatchedRecently(applicantID,opponentID,date)){
    console.log('一部の例外を除き同じカードの対戦は前回の対戦から' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。');
    writeLogsInFormResponse('一部の例外を除き同じカードの対戦は前回の対戦から' + SAME_OPPONENT_COOLDOWN_DAYS + '日以上空けなければなりません。',true);
    return;
  }

  const modificationFlag = canUseModification ? '可' : '不可';

  let slotString;

  if(slot === '部活時間外' || slot === 'その他'){
    slotString = slot;
  }else{
    const nextMatchNumber = countFridayMatch(date);
    slotString = '部活中(' + nextMatchNumber + '試合目)';
  }

  rankMatchScheduleSheet.appendRow([applicantID,applicantName,opponentID,opponentName,date,slotString,'','',modificationFlag,new Date(),'']);
  manageChallenge(applicantID,true,isMale);
}

// 挑戦権を管理する関数
// isUsedがtrueなら減らし、falseなら増やす
function manageChallenge(id,isUsed,isMale){
  let lastRow,rankData;
  
  if(isMale){
    lastRow = maleSheet.getLastRow();
    rankData = maleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();

  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();
  }

  const rowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === id);
  
  if(rowIndex === -1){
    console.log('プレイヤーが順位表から見つかりませんでした。');
    writeLogsInFormResponse('プレイヤーが順位表から見つかりませんでした。',true);
    return;
  }

  if(isMale){
    if(isUsed)maleSheet.getRange(rowIndex + HEADER_ROW_OFFSET + 1,MATCH_LIMIT_COLUMN + 1).setValue(Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) - 1);
    else maleSheet.getRange(rowIndex + HEADER_ROW_OFFSET + 1,MATCH_LIMIT_COLUMN + 1).setValue(Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) + 1);
  }else{
    if(isUsed)femaleSheet.getRange(rowIndex + HEADER_ROW_OFFSET + 1,MATCH_LIMIT_COLUMN + 1).setValue(Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) - 1);
    else femaleSheet.getRange(rowIndex + HEADER_ROW_OFFSET + 1,MATCH_LIMIT_COLUMN + 1).setValue(Number(rankData[rowIndex][MATCH_LIMIT_COLUMN]) + 1);
  }
}

function isMatchedRecently(applicantID,opponentID,date){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= HEADER_ROW_OFFSET){
    return false;
  }

  const scope = new Date();
  scope.setHours(0,0,0,0);
  scope.setDate(date.getDate() - SAME_OPPONENT_COOLDOWN_DAYS);

  const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
  let recentlyCounterMatchedFlag = false;
  matchData.forEach((row) => {
    if(!(row[APPLICANT_ID_COLUMN] === opponentID && row[OPPONENT_ID_COLUMN] === applicantID))return;
    if(new Date(row[MATCH_DATE_COLUMN]) > scope && row[MATCH_RESULT_FLAG_COLUMN] !== '' && row[MATCH_RESULT_COLUMN] !== '敗北'){
      recentlyCounterMatchedFlag = true;
    }
  })
  if(recentlyCounterMatchedFlag)return true;

  let recentlyReMatchedFlag = false;
  matchData.forEach((row) => {
    if(!(row[APPLICANT_ID_COLUMN] === applicantID && row[OPPONENT_ID_COLUMN] === opponentID))return;
    if(new Date(row[MATCH_DATE_COLUMN]) > scope && row[MATCH_RESULT_COLUMN] === '敗北'){
      recentlyReMatchedFlag = true;
    }
  })

  if(recentlyReMatchedFlag)return true;
  else return false;
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
function writeMatchResult(applicant,opponent,matchResult,game1Score,game2Score,game3Score){
  const applicantID = applicant.substring(applicant.length-9,applicant.length-1);
  const opponentID  = opponent .substring(opponent .length-9,opponent .length-1);

  console.log(applicantID);
  console.log(opponentID);

  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= HEADER_ROW_OFFSET){
    console.log('対象の試合が見つかりませんでした。この結果報告は無効です。');
    writeLogsInFormResponse('対象の試合が見つかりませんでした。この結果報告は無効です。',false);
    return;
  }

  let isMale = false;

  if(applicant[1] === '男'){
    isMale = true;
  }

  const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();

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

      rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1 + idx,MATCH_RESULT_FLAG_COLUMN + 1,1,2).setValues([['済',matchResult]]);
      rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1 + idx,RESULT_FORM_TIMESTAMP_COLUMN + 1).setValue(new Date());

      applytime = new Date(row[SCHEDULE_FORM_TIMESTAMP_COLUMN]);

      isExecuted = true;
    }
  })

  if(!isExecuted){
    console.log('対象の試合が見つかりませんでした。この結果報告は無効です。');
    writeLogsInFormResponse('対象の試合が見つかりませんでした。この結果報告は無効です。',false);
    return;
  }

  console.log('試合結果を正常に書き込みました。');

  if(matchResult === '敗北'){
    removeWinningBonus(applicantID,isMale);
    return;
  }

  console.log('順位表の変動を開始します。');

  changeRanking(applicantID,opponentID,isMale,applytime);
}

// ランキングを変更する関数
function changeRanking(applicantID,opponentID,isMale,applytime){

  let lastRow,rankData;
  
  if(isMale){
    lastRow = maleSheet.getLastRow();
    rankData = maleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();

  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();
  }

  const applicantRowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === applicantID);
  const opponentRowIndex  = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === opponentID );

  if (applicantRowIndex === -1 || opponentRowIndex === -1) {
    console.log('プレイヤーが順位表から見つかりませんでした。');
    writeLogsInFormResponse('プレイヤーが順位表から見つかりませんでした。',false);
    return;
  }
  
  if(applicantRowIndex <= opponentRowIndex){
    console.log('申込者の順位の方が対戦相手よりも高いです。');
    writeLogsInFormResponse('申込者の順位の方が対戦相手よりも高いです。',false);
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
    opponentRowIndex = applicantRowIndex - min(applicantRowIndex - opponentRowIndex,MAX_RANK_DIFFERENCE);
  }

  if(isMale){
    const subRanking = rankData.slice(opponentRowIndex,applicantRowIndex + 1);
    const removedRankSubRanking = subRanking.map((row) => row.slice(1));
    maleSheet.getRange(HEADER_ROW_OFFSET + 1 + 1 + opponentRowIndex,2,applicantRowIndex - opponentRowIndex,6).setValues(removedRankSubRanking.slice(0,-1));
    maleSheet.getRange(HEADER_ROW_OFFSET + 1 + opponentRowIndex,2,1,6).setValues(removedRankSubRanking.slice(-1));
  }else{
    const subRanking = rankData.slice(opponentRowIndex,applicantRowIndex + 1);
    const removedRankSubRanking = subRanking.map((row) => row.slice(1));
    femaleSheet.getRange(HEADER_ROW_OFFSET + 1 + 1 + opponentRowIndex,2,applicantRowIndex - opponentRowIndex,6).setValues(removedRankSubRanking.slice(0,-1));
    femaleSheet.getRange(HEADER_ROW_OFFSET + 1 + opponentRowIndex,2,1,6).setValues(removedRankSubRanking.slice(-1));
  }

  console.log('ランキングを正常に変更しました。');

  buildRankingsByDepartment();
}

// 連勝ボーナスを剥奪する関数
function removeWinningBonus(applicantID,isMale){
  let lastRow,rankData;
  
  if(isMale){
    lastRow = maleSheet.getLastRow();
    rankData = maleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();
  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,7).getValues();
  }

  const applicantRowIndex = rankData.findIndex((row) => row[PLAYER_ID_COLUMN] === applicantID);

  if(isMale){
    maleSheet.getRange(HEADER_ROW_OFFSET + 1 + applicantRowIndex,MATCH_MORE_FLAG_COLUMN + 1).setValue('不可');
  }else{
    femaleSheet.getRange(HEADER_ROW_OFFSET + 1 + applicantRowIndex,MATCH_MORE_FLAG_COLUMN + 1).setValue('不可');
  }
  return;
}

// 試合申し込みから対戦相手が勝ち上がったかどうかを判定する関数
function isOpponentPromoted(time,opponentID){
  const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
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

//時系列順にソートする関数
function sortRankMatchSchedule(){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= HEADER_ROW_OFFSET){
    return;
  }

  const matchData = rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).getValues();
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
  rankMatchScheduleSheet.getRange(HEADER_ROW_OFFSET + 1,1,lastRow-1,11).setValues(matchData);

  console.log('日程を時系列順にソートしました。');
}

//フォームのプルダウンを変更する関数
function updateFormDropdown() {
  const choices = [];

  const maleData = maleSheet.getRange('B2:C' + maleSheet.getLastRow()).getValues(); // イベント情報を取得
  for (let i = 0; i < maleData.length; i++) {
    const studentID = maleData[i][PLAYER_ID_COLUMN];
    const studentName = maleData[i][PLAYER_NAME_COLUMN];
    choices.push('(男) ' + studentName + ' (' + studentID + ')');
  }

  const femaleData = femaleSheet.getRange('B2:C' + femaleSheet.getLastRow()).getValues(); // イベント情報を取得
  for (let i = 0; i < femaleData.length; i++) {
    const studentID = femaleData[i][PLAYER_ID_COLUMN];
    const studentName = femaleData[i][PLAYER_NAME_COLUMN];
    choices.push('(女) ' + studentName + ' (' + studentID + ')');
  }

  
  const matchSchedulingForm = FormApp.openById(MATCH_SCHEDULING_FORM_ID);
  const matchResultForm = FormApp.openById(MATCH_RESULT_FORM_ID);

  const matchSchedulingItems = matchSchedulingForm.getItems();
  const matchResultItems = matchResultForm.getItems();
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
  console.log('挑戦権を回復させました。');
}

// フォームに対し、処理のログを書き込む関数
function writeLogsInFormResponse(message,isScheduleForm){
  if(isScheduleForm){
    const lastRow = scheduleFormResponsesSheet.getLastRow();
    if(lastRow === HEADER_ROW_OFFSET)return;
    const previousMessage = scheduleFormResponsesSheet.getRange(lastRow,9).getValue();
    scheduleFormResponsesSheet.getRange(lastRow,9).setValue(previousMessage + message);
  }else{
    const lastRow = resultFormResponsesSheet.getLastRow();
    if(lastRow === HEADER_ROW_OFFSET)return;
    const previousMessage = resultFormResponsesSheet.getRange(lastRow,8).getValue();
    resultFormResponsesSheet.getRange(lastRow,8).setValue(previousMessage + message);
  }
}

// 学科別順位表を構築する関数
function buildRankingsByDepartment(){
  const maleLastRow = maleSheet.getLastRow();
  const maleData = maleSheet.getRange(HEADER_ROW_OFFSET + 1,1,maleLastRow-1,4).getValues();
  const femaleLastRow = femaleSheet.getLastRow();
  const femaleData = femaleSheet.getRange(HEADER_ROW_OFFSET + 1,1,femaleLastRow-1,4).getValues();

  let maleMedicineData = [];
  let maleInsuranceData = [];

  for(let i = 0;i < maleLastRow-1;i++){
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

  for(let i = 0;i < femaleLastRow-1;i++){
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