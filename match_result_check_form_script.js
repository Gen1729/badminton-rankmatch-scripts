const ss = SpreadsheetApp.openById('');
const rankMatchScheduleSheet = ss.getSheetByName("ランク戦日程");

function onFormSubmit(e) {
  try{
    const itemResponses = e.response.getItemResponses();
    const formData = itemResponses.map(item => item.getResponse());
    console.log("結果確認受信: " + JSON.stringify(formData));
    const applicant = formData[0];
    const opponent = formData[1] ? formData[1] : null;

    const applicantNameAndId = applicant.substring(3);
    const opponentNameAndId = opponent ? opponent.substring(3) : null;

    const applicantID = applicantNameAndId.match(/\(([^)]+)\)/)?.[1] ?? 'xxxxxxxx';
    const opponentID = opponent ? (opponentNameAndId.match(/\(([^)]+)\)/)?.[1] ?? 'xxxxxxxx') : null;

    const email = e.response.getRespondentEmail();

    const lastRow = rankMatchScheduleSheet.getLastRow();

    if(lastRow <= 1){
      return;
    }

    let victoryCount = 0;
    let defeatCount = 0;

    const matchData = rankMatchScheduleSheet.getRange(2,1,lastRow,14).getValues();
    const filteredData = matchData.filter((row) => {
      if(!(row[0] === applicantID) && !(row[2] === applicantID))return false;
      if(opponent && !(row[0] === opponentID) && !(row[2] === opponentID))return false;
      if(row[6] === '')return false;
      if(row[7] === '敗北'){
        if(row[0] === applicantID)defeatCount++;
        else victoryCount++;
      }else{
        if(row[0] === applicantID)victoryCount++;
        else defeatCount++;
      }
      return true;
    })
    sendResultEmail(email,filteredData,victoryCount,defeatCount,applicantNameAndId,opponentNameAndId);
    
  } catch (err) {
    console.log('エラーが発生しました。' + err);
  }
}

function sendResultEmail(mailAdress,data,victory,defeat,applicant,opponent) {
  const recipient = mailAdress;
  const subject = "ランク戦結果確認";
  let body = '';
  if(opponent){
    body = `以下、${applicant} さん 対 ${opponent} さんの過去のランク戦結果です。`;
  }else{
    body = `以下、${applicant} さんの過去のランク戦結果です。`;
  }

  const matchResult = data.map(row => {
    const scores = [row[8], row[9], row[10]].map(s => s || 'x-x').join(' ');
    const d = new Date(row[4]);
    const formatted = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
    return `${formatted} ${row[7]} ${row[1]} (${row[0]}) vs ${row[3]} (${row[2]}) ${scores}`;
  }).join('\n');

  GmailApp.sendEmail(recipient, subject, body + '\n' + matchResult + '\n\n' + victory + '勝' + defeat + '敗');
}
