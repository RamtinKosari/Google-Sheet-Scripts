function sendEmailIfTimeExpired() {
    //-- DO NOT EDIT THIS LINE --//
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    /**
     * Recipient Email
     * @abstract Email address of the recipient
     * @description The script will send an email to this address if the time has expired.
     * Note that you need to grant permission to the script to send emails via your email.
     */
    var recipient = "ramtinkosari@gmail.com"; // Replace with your email
    /**
     * Array of Cells to Check
     * @abstract Cells to check for time expiration
     * @description The script will check the value of these cells and if the value is a date or number, it will compare it with the current time.
     * @example Here is an example of the array:
     * ['BJ23', 'BJ25', 'BJ27', 'BJ29', 'BJ31', ...]
     */
    var cellsToCheck = [
        'BJ23', 'BJ25', 'BJ27', 'BJ29', 'BJ31', 'BJ33', 'BJ35',
        'BJ37', 'BJ39', 'BJ41', 'BJ43', 'BJ45', 'BJ47', 'BJ49',
        'BJ51', 'BJ53', 'BJ55', 'BJ57', 'BJ59', 'BJ61', 'BJ63',
        'BJ65', 'BJ67', 'BJ69', 'BJ71', 'BJ73', 'BJ75', 'BJ77'
    ];
    /**
     * Prefix Column
     * @abstract Column for email messages
     * @description The script will get the value of the cell in this column to include in the email message.
     * @example Here is an example of the prefix column:
     * AT23 has Value of 'Plan 1', so the email message will be 'Time of Plan 1 has been Passed.'
     * in this example, we assume that the plan's cell is in column AT. so it is in same row as the time cell.
     */
    var prefixColumn = 'AT';
    /**
     * Email Interval
     * @abstract Interval for sending emails
     * @description The script will check if the last email was sent within this interval.
     * @example Here is an example of the email interval:
     * 10 * 60 * 1000 = 10 minutes
     * The script will send an email if the last email was sent more than 10 minutes ago.
     */
    var emailInterval = 10* 60 * 1000;
    /**
     * Current Time
     */
    var currentTime = new Date().getTime();
    //-- DO NOT EDIT BELOW THIS LINE --//
    var scriptProperties = PropertiesService.getScriptProperties();
    cellsToCheck.forEach(cell => {
        var range = sheet.getRange(cell);
        //-- Get the Value of the Cell
        var timeLeft = range.getValue();
        //-- Convert Date to Time
        if (timeLeft instanceof Date) {
            timeLeft = timeLeft.getTime();
        }
        //-- Get the Value of the Plan's Column
        var targetCell = prefixColumn + cell.match(/\d+/)[0];
        //-- Get the Content of the Cell
        var cellContent = sheet.getRange(targetCell).getValue();
        //-- Check if the Time has Passed
        if (typeof timeLeft === 'number' && (timeLeft - currentTime) <= 0) {
            //-- Get the Last Email Key
            var lastEmailKey = `email_sent_${cell}`;
            //-- Get the Last Email Time
            var lastEmailTime = scriptProperties.getProperty(lastEmailKey);
            //-- Send Email if the Last Email was Sent More than the Interval
            if (!lastEmailTime || (currentTime - parseInt(lastEmailTime, 10)) > emailInterval) {
                //-- Email Subject
                var subject = "Time Alert: Deadline Reached";
                //-- Email Message
                var message = `Time of Plan '${cellContent}' has been Passed since ${new Date(timeLeft).toLocaleString()}.`;
                //-- Send Email
                GmailApp.sendEmail(recipient, subject, message);
                //-- Save the Last Email Time
                scriptProperties.setProperty(lastEmailKey, currentTime.toString());
            }
        }
    });
}