const ss = SpreadsheetApp.getActiveSpreadsheet();
const maleSheet = ss.getSheetByName("男子");
const femaleSheet = ss.getSheetByName("女子");
const rankMatchScheduleSheet = ss.getSheetByName("ランク戦日程");

const MATCH_SCHEDULING_FORM_ID = '';
const MATCH_RESULT_FORM_ID = '';

const MAX_RANK_DIFFERENCE = 'B1';
const MATCH_ACCEPT_DAY_LIMIT = 'B3';
const SAME_OPPONENT_COOLDOWN_DAYS = 'B5';
const MONTH_APPLICATION_LIMIT = 'B7';


// フォームを受け取った時の分岐
// 日程報告か結果報告か
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    Logger.log('他の処理が終わっていないためスキップします。');
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
    applicationScope.setDate(applicationScope.getDate() + 7 + 1);

    if(applicant === opponent){
      console.log('対戦する人が同一人物です。入力は無効です。');
      return;
    }

    if(applicant[1] !== opponent[1]){
      console.log('対戦する人の性別が違います。入力は無効です。');
      return;
    }

    if(originalDate < today){
      console.log('日付が過去のものです。入力は無効です。');
      return;
    }

    if(applicationScope < originalDate){
      console.log('日付が未来すぎます。' + '7' + '日以内の日程のみ許可します。入力は無効です。');
      return;
    }

    if(modifiedDate && applicationScope < modifiedDate){
      console.log('日付が未来すぎます。' + '7' + '日以内の日程のみ許可します。入力は無効です。');
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
      return;
    }

    console.log('日程追加操作を実行します。');
    processNormalRequest(applicant,opponent,originalDate,timeSlot);
    sortRankMatchSchedule();
    SpreadsheetApp.flush();

  } catch (err) {
    console.log('日程報告中に予期せぬエラーが発生しました。' + err);
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
      return;
    }

    if(applicant[1] !== opponent[1]){
      console.log('対戦する人の性別が違います。入力は無効です。');
      return;
    }

    console.log('結果報告の書き込みを開始します。');
    writeMatchResult(applicant,opponent,matchResult,game1Score,game2Score,game3Score);
    updateFormDropdown();
    SpreadsheetApp.flush();
  } catch (err) {
    console.log('結果報告中に予期せぬエラーが発生しました。' + err);
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
    if(lastRow <= 1){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      return;
    }

    let isMale = false;

    if(applicant[1] === '男'){
      isMale = true;
    }

    const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();
    let deleteFlag = false;
    matchData.forEach((row,idx) => {
      if(row[0] === applicantID && row[2] === opponentID && (new Date(row[4])).getTime() === originalDate.getTime()){
        
        if(timeSlot === '金曜の強化練@青葉体育館'){
          if(row[5] === '部活時間外' || row[5] === 'その他'){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }else{
          if(row[5] !== timeSlot){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }

        if(row[6] !== ''){
          console.log('該当の試合は存在しますが、すでに結果報告を受け取っているためキャンセルはできません。');
          return;
        }

        if(row[5] !== '部活時間外' && row[5] !== 'その他'){
          if(Number(row[5][4]) !== 3){
            narrowSchedule(originalDate,Number(row[5][4]));
          }
        }

        rankMatchScheduleSheet.deleteRow(idx + 4);
        manageChallenge(applicantID,false,isMale);
        console.log('該当の試合を削除しました。');
        deleteFlag = true;
      }
    })
    if(!deleteFlag)console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');

  } catch (err) {
    console.log('キャンセル処理中にエラーが発生しました。' + err);
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
    if(lastRow <= 1){
      console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
      return;
    }

    let isMale = false;

    if(applicant[1] === '男'){
      isMale = true;
    }

    if(modifiedTimeSlot === '金曜の強化練@青葉体育館'){
      if(!isFriday(modifiedDate)){
        console.log('強化練を選んだ場合、日付は金曜日でなくてはいけません。');
        return;
      }
    }

    if(isSlotBooked(modifiedDate,modifiedTimeSlot)){
      console.log('変更後の日程はすでに埋まっています。');
      return;
    }

    const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();
    let modifyFlag = false;
    matchData.forEach((row,idx) => {
      if(row[0] === applicantID && row[2] === opponentID && (new Date(row[4])).getTime() === originalDate.getTime()){
        
        if(timeSlot === '金曜の強化練@青葉体育館'){
          if(row[5] === '部活時間外' || row[5] === 'その他'){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }else{
          if(row[5] !== timeSlot){
            console.log('対戦カードと日付は合っていますが、時間帯が正しくありません。入力を再度確認してください。');
            return;
          }
        }

        if(row[6] !== ''){
          console.log('該当の試合は存在しますが、すでに結果報告を受け取っているためキャンセルはできません。');
          return;
        }

        if(row[8] !== '可'){
          console.log('該当の試合は存在しますが、すでに一度日程/時間帯変更を行なっているため変更はできません。')
          return;
        }

        if(row[5] !== '部活時間外' && row[5] !== 'その他'){
          if(Number(row[5][4]) !== 3){
            narrowSchedule(originalDate,Number(row[5][4]));
          }
        }

        rankMatchScheduleSheet.deleteRow(idx + 4);
        manageChallenge(applicantID,false,isMale);
        pushNewMatch(applicant,opponent,modifiedDate,modifiedTimeSlot,false);
        console.log('該当の試合の日程/時間帯を変更しました。');
        modifyFlag = true;
      }
    })
    if(!modifyFlag)console.log('対象の試合が見つかりませんでした。入力を再度確認してください。');
  } catch (err) {
    console.log('日程/時間帯変更処理中にエラーが発生しました。' + err);
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

    if(timeSlot === '金曜の強化練@青葉体育館'){
      if(!isFriday(originalDate)){
        console.log('強化練を選んだ場合、日付は金曜日でなくてはいけません。');
        return;
      }
    }

    if(isSlotBooked(originalDate,timeSlot)){
      console.log('その日程はすでに埋まっています。');
      return;
    }

    let isMale = false;

    if(applicant[1] === '男'){
      isMale = true;
    }

    if(!canPlayMatch(applicantID,opponentID,isMale)){
      console.log('この組み合わせの試合は行えません。');
      return;
    }

    pushNewMatch(applicant,opponent,originalDate,timeSlot,true);
  } catch (err) {
    console.log('日程追加処理中にエラーが発生しました。' + err);
  }
}

// その日が試合で埋まっているかを判定する関数
// 金曜日の強化練の時は判定するがそれ以外の場合は必ずfalseを返す
function isSlotBooked(date,slot){
  if(slot !== '金曜の強化練@青葉体育館')return false;

  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= 1){
    return false;
  }

  const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();
  let sameDayCount = 0;
  matchData.forEach((row) => {
    if((new Date(row[4])).getTime() === date.getTime()){
      if(row[5] !== '部活時間外' && row[5] !== 'その他'){
        sameDayCount += 1;
      }
    }
  })

  if(sameDayCount >= 3){
    return true;
  }

  return false;
}

// 金曜日のマッチが何試合既に入っているかを数える関数
function countFridayMatch(date){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= 1){
    return 1;
  }

  const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();
  let sameDayCount = 0;
  matchData.forEach((row) => {
    if((new Date(row[4])).getTime() === date.getTime()){
      if(row[5] !== '部活時間外' && row[5] !== 'その他'){
        sameDayCount += 1;
      }
    }
  })

  return sameDayCount + 1;
}

// 金曜の１、２試合目が無くなったらその後にあった試合を詰める関数
function narrowSchedule(date,matchNumber){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();
  matchData.forEach((row,idx) => {
    if((new Date(row[4])).getTime() === date.getTime()){
      if(row[5] !== '部活時間外' && row[5] !== 'その他'){
        if(Number(row[5][4]) > matchNumber){
          rankMatchScheduleSheet.getRange(1 + 1 + idx,6).setValue('部活中(' + (Number(row[5][4]) - 1) + '試合目)');
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
    rankData = maleSheet.getRange(2,1,lastRow-1,7).getValues();

  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(2,1,lastRow-1,7).getValues();
  }

  const applicantRowIndex = rankData.findIndex((row) => row[1] === applicantID);
  const opponentRowIndex  = rankData.findIndex((row) => row[1] === opponentID );
  
  if (applicantRowIndex === -1 || opponentRowIndex === -1) {
    console.log('プレイヤーが順位表から見つかりませんでした。');
    return false;
  }
  
  if(applicantRowIndex <= opponentRowIndex){
    console.log('申込者の順位の方が対戦相手よりも高いです。');
    return false;
  }

  if(Number(rankData[applicantRowIndex][5]) <= 0){
    console.log('今月の挑戦権を使い切っています。');
    return false;
  }

  if(rankData[applicantRowIndex][4] !== '可' || rankData[opponentRowIndex][4] !== '可'){
    console.log('プレイヤーが試合できない状態です。');
    return false;
  }

  let cannotPlayMatchPlayersCount = 0;

  for(let i = opponentRowIndex; i < applicantRowIndex; i++){
    if(rankData[opponentRowIndex][4] !== '可')cannotPlayMatchPlayersCount += 1;
  }

  if(cannotPlayMatchPlayersCount <= 1){
    if(applicantRowIndex - opponentRowIndex <= 3){
      return true;
    }else{
      return false;
    }
  }else{
    if(applicantRowIndex - opponentRowIndex <= 3 + cannotPlayMatchPlayersCount){
      return true;
    }else{
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
    console.log('同じカードの対戦は前回の対戦から' + '21' + '日以上空けなければなりません。日程/時間帯変更の場合はキャンセルと同じ扱いになります。');
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

  rankMatchScheduleSheet.appendRow([applicantID,applicantName,opponentID,opponentName,date,slotString,'','',modificationFlag]);
  manageChallenge(applicantID,true,isMale);
}

// 挑戦権を管理する関数
// isUsedがtrueなら減らし、falseなら増やす
function manageChallenge(id,isUsed,isMale){
  let lastRow,rankData;
  
  if(isMale){
    lastRow = maleSheet.getLastRow();
    rankData = maleSheet.getRange(2,1,lastRow-1,7).getValues();

  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(2,1,lastRow-1,7).getValues();
  }

  const rowIndex = rankData.findIndex((row) => row[1] === id);
  
  if(rowIndex === -1){
    console.log('プレイヤーが順位表から見つかりませんでした。');
    return;
  }

  if(isMale){
    if(isUsed)maleSheet.getRange(rowIndex+2,6).setValue(Number(rankData[rowIndex][5]) - 1);
    else maleSheet.getRange(rowIndex+2,6).setValue(Number(rankData[rowIndex][5]) + 1);
  }else{
    if(isUsed)femaleSheet.getRange(rowIndex+2,6).setValue(Number(rankData[rowIndex][5]) - 1);
    else femaleSheet.getRange(rowIndex+2,6).setValue(Number(rankData[rowIndex][5]) + 1)
  }
}

function isMatchedRecently(applicantID,opponentID,date){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= 1){
    return false;
  }

  const scope = new Date();
  scope.setHours(0,0,0,0);
  scope.setDate(date.getDate() - 21);

  const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();
  let recentlyFlag = false;
  matchData.forEach((row) => {
    if(!((row[0] === applicantID && row[2] === opponentID) || (row[0] === opponentID && row[2] === applicantID)))return;
    if(new Date(row[4]) > scope){
      recentlyFlag = true;
    }
  })
  if(recentlyFlag)return true;
  else return false;
}

// 金曜日かどうか判定する関数
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
  if(lastRow <= 1){
    console.log('対象の試合が見つかりませんでした。この結果報告は無効です。');
    return;
  }

  let isMale = false;

  if(applicant[1] === '男'){
    isMale = true;
  }

  const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();

  const today = new Date();
  today.setHours(0,0,0,0);

  let isExecuted = false;

  matchData.forEach((row,idx) => {
    if(isExecuted)return;
    if(row[0] === applicantID && row[2] === opponentID && row[6] === ''){

      if((new Date(row[4])).getTime() !== today.getTime()){
        console.log('同カードで結果未提出の試合がありますが、日付が今日のものではありません。')
        return;
      }

      rankMatchScheduleSheet.getRange(1 + 1 + idx,6 + 1,1,2).setValues([['済',matchResult]]);
      isExecuted = true;
    }
  })

  if(!isExecuted){
    console.log('対象の試合が見つかりませんでした。この結果報告は無効です。');
    return;
  }

  console.log('試合結果を正常に書き込みました。');

  if(matchResult === '敗北'){
    removeWinningBonus(applicantID,isMale);
    return;
  }

  console.log('順位表の変動を開始します。');

  changeRanking(applicantID,opponentID,isMale);
}

// ランキングを変更する関数
function changeRanking(applicantID,opponentID,isMale){

  let lastRow,rankData;
  
  if(isMale){
    lastRow = maleSheet.getLastRow();
    rankData = maleSheet.getRange(2,1,lastRow-1,7).getValues();

  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(2,1,lastRow-1,7).getValues();
  }

  const applicantRowIndex = rankData.findIndex((row) => row[1] === applicantID);
  const opponentRowIndex  = rankData.findIndex((row) => row[1] === opponentID );

  if (applicantRowIndex === -1 || opponentRowIndex === -1) {
    console.log('プレイヤーが順位表から見つかりませんでした。');
    return;
  }
  
  if(applicantRowIndex <= opponentRowIndex){
    console.log('申込者の順位の方が対戦相手よりも高いです。');
    return;
  }

  if(isMale){
    if(Number(rankData[applicantRowIndex][5]) === 0 && rankData[applicantRowIndex][6] === '可'){
      console.log('連勝ボーナスを発動します。');
      rankData[applicantRowIndex][5] = '1';
    }
  }else{
    if(Number(rankData[applicantRowIndex][5]) === 0 && rankData[applicantRowIndex][6] === '可'){
      console.log('連勝ボーナスを発動します。');
      rankData[applicantRowIndex][5] = '1';
    }
  }

  if(isMale){
    const subRanking = rankData.slice(opponentRowIndex,applicantRowIndex + 1);
    const removedRankSubRanking = subRanking.map((row) => row.slice(1));
    maleSheet.getRange(2 + 1 + opponentRowIndex,2,applicantRowIndex - opponentRowIndex,6).setValues(removedRankSubRanking.slice(0,-1));
    maleSheet.getRange(2 + opponentRowIndex,2,1,6).setValues(removedRankSubRanking.slice(-1));
  }else{
    const subRanking = rankData.slice(opponentRowIndex,applicantRowIndex + 1);
    const removedRankSubRanking = subRanking.map((row) => row.slice(1));
    femaleSheet.getRange(2 + 1 + opponentRowIndex,2,applicantRowIndex - opponentRowIndex,6).setValues(removedRankSubRanking.slice(0,-1));
    femaleSheet.getRange(2 + opponentRowIndex,2,1,6).setValues(removedRankSubRanking.slice(-1));
  }

  console.log('ランキングを正常に変更しました。');
}

// 連勝ボーナスを剥奪する関数
function removeWinningBonus(applicantID,isMale){
  let lastRow,rankData;
  
  if(isMale){
    lastRow = maleSheet.getLastRow();
    rankData = maleSheet.getRange(2,1,lastRow-1,7).getValues();
  }else{
    lastRow = femaleSheet.getLastRow();
    rankData = femaleSheet.getRange(2,1,lastRow-1,7).getValues();
  }

  const applicantRowIndex = rankData.findIndex((row) => row[1] === applicantID);

  if(isMale){
    maleSheet.getRange(2 + applicantRowIndex,7).setValue('不可');
  }else{
    femaleSheet.getRange(2 + applicantRowIndex,7).setValue('不可');
  }
  return;
}

//時系列順にソートする関数
function sortRankMatchSchedule(){
  const lastRow = rankMatchScheduleSheet.getLastRow();
  if(lastRow <= 1){
    return;
  }

  const matchData = rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).getValues();
  matchData.sort((a, b) => new Date(a[4]).getTime() - new Date(b[4]).getTime());
  rankMatchScheduleSheet.getRange(1 + 1,1,lastRow-1,9).setValues(matchData);

  console.log('日程を時系列順にソートしました。');
}

//フォームのプルダウンを変更する関数
function updateFormDropdown() {
  const choices = [];

  const maleData = maleSheet.getRange('B2:C' + maleSheet.getLastRow()).getValues(); // イベント情報を取得
  for (let i = 0; i < maleData.length; i++) {
    const studentID = maleData[i][0];
    const studentName = maleData[i][1];
    choices.push('(男) ' + studentName + ' (' + studentID + ')');
  }

  const femaleData = femaleSheet.getRange('B2:C' + femaleSheet.getLastRow()).getValues(); // イベント情報を取得
  for (let i = 0; i < femaleData.length; i++) {
    const studentID = femaleData[i][0];
    const studentName = femaleData[i][1];
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

//月初めに挑戦権を回復させる関数(定期実行)
function restoreChallengeRight(){
  maleSheet.getRange('F2:F' + maleSheet.getLastRow()).setValue('2');
  maleSheet.getRange('G2:G' + maleSheet.getLastRow()).setValue('可');
  femaleSheet.getRange('F2:F' + femaleSheet.getLastRow()).setValue('2');
  femaleSheet.getRange('G2:G' + femaleSheet.getLastRow()).setValue('可');
  console.log('挑戦権を回復させました。');
}