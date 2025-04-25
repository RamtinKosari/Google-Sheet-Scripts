function sendDailyWords() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English');
  
    //-- Select Range
    const data = sheet.getRange(13, 2, 1210 - 13 + 1, 29).getValues();
  
    //-- Map Words
    const words = data.map((row, index) => ({
      rowIndex: index + 13,
      word: row[0],
      partOfSpeech: row[3],
      meaning: row[6],
      synonyms: row[12],
      pronunciation: row[17],
      use: parseInt(row[26]) || 0,
      count: parseInt(row[28]) || 0
    }))
    //-- Skip Empty Words
    .filter(w => w.word && w.word.toString().trim() !== '')
    //-- Use Selected Words
    .filter(w => w.count === 1);
    //-- Sort
    words.sort((a, b) => a.count - b.count);
    //-- Pick
    const selectedWords = words.slice(0, 20);
    //-- Build Email Body as HTML Table
    let emailBody = `
      <p>Here are your <b>20 daily English words</b>:</p>
      <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif;">
        <tr style="background-color: #f2f2f2;">
          <th>Word</th>
          <th>Part of Speech</th>
          <th>Meaning</th>
          <th>Synonyms</th>
          <th>Pronunciation</th>
        </tr>
    `;
    selectedWords.forEach(w => {
      emailBody += `
        <tr>
          <td>${w.word}</td>
          <td>${w.partOfSpeech}</td>
          <td>${w.meaning}</td>
          <td>${w.synonyms}</td>
          <td>${w.pronunciation}</td>
        </tr>
      `;
    });
    emailBody += `</table>`;
    //-- Send Email
    if (selectedWords.length > 0) {
      MailApp.sendEmail({
        to: 'ramtinkosari@gmail.com',
        subject: 'Daily 20 English Words',
        htmlBody: emailBody
      });
      //-- Increase Count
      selectedWords.forEach(w => {
        sheet.getRange(w.rowIndex, 28).setValue(w.count + 1);
      });
    } else {
      Logger.log('No words with Review Status = 1 found.');
    }
  }
  