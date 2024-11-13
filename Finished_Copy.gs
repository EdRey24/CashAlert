/*
const STUDENT_ID_COL = 'Student ID';
const FIRST_NAME_COL = 'First Name';
const LAST_NAME_COL = 'Last Name';
const RECIPIENT_COL = 'Email';
const TOTAL_COL = 'Total';
const EMAIL_SENT_COL = 'Email Sent';
const SEND_EMAIL_COL = 'Send Email';
const PAYER_FIRST_NAME_COL = 'Payer First Name';
const PAYER_LAST_NAME_COL = 'Payer Last Name';
const PAYER_STUDENT_ID_COL = 'Payer Student ID';
const PAYMENT_AMOUNT_COL = 'Payment Amount';
const PAYMENT_REASON_COL = 'Reason';
const OFFICER_COL = 'Officer';
const DATE_COL = 'Date';
const PAYMENT_SENT_COL = 'Payment Email Sent';
const weeklyEmailTemplate = HtmlService.createTemplateFromFile('WeeklyEmail');
const paymentEmailTemplate =
	HtmlService.createTemplateFromFile('PaymentsEmail');
const emailJS = HtmlService.createTemplateFromFile('WeeklyEmailJavaScript');
const tableEntries = [];

//Add new fundraisers here
const FUNDRAISER_COLS = [
	'Fundraiser #1',
	'Fundraiser #2',
];

//Formats time for Email Sent to CST
const formattedDate = Utilities.formatDate(
	new Date(),
	'CST',
	"MM/dd/yyyy' at 'hh:mm:ss a"
);

//Formats debt amount into USD currency
const formatter = new Intl.NumberFormat('en-US', {
	style: 'currency',
	currency: 'USD',
	minimumFractionDigits: 2,
});

//This function creates the menu for the Cash Alert on Google Sheets
function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Cash Alert')
		.addItem('Send Payment Confirmation', 'sendPaymentConfirmation')
		.addToUi();
}

//This function is what builds the tables from the data for the emails
function buildTable(currentRow) {
	//Iterates through every fundrasier we have in the FUNDRAISER_COL array
	for (var i = 0; i < FUNDRAISER_COLS.length; i++) {
		//Checks if the amount still owed for the fundraiser is over $0 and if so it adds it to the table for the email
		if (currentRow[FUNDRAISER_COLS[i]] > 0) {
			emailJS.fundraiser = FUNDRAISER_COLS[i];
			emailJS.fundraiser_amount = formatter.format(
				currentRow[FUNDRAISER_COLS[i]]
			);
			var finishedTable = emailJS.evaluate().getContent();
			tableEntries.push(finishedTable);
		}
	}
}

function buildPaymentTable(currentRow, dataRows) {
	var match;
	for (var i = 0; i < dataRows.length; i++) {
		if (matchPerson(dataRows[i]) == true) break;
	}

	function matchPerson(newRow) {
		if (currentRow[PAYER_STUDENT_ID_COL] == newRow[STUDENT_ID_COL]) {
			match = newRow;
			return true;
		}
	}

	for (var i = 0; i < FUNDRAISER_COLS.length; i++) {
    if ((FUNDRAISER_COLS[i] == currentRow[PAYMENT_REASON_COL]) && match[TOTAL_COL] == 0){
      emailJS.fundraiser = FUNDRAISER_COLS[i];
      emailJS.fundraiser_amount = formatter.format(
					match[FUNDRAISER_COLS[i]] + currentRow[PAYMENT_AMOUNT_COL]
				);
      paymentEmailTemplate.amount = formatter.format(match[TOTAL_COL]);
      var finishedTable = emailJS.evaluate().getContent();
			paymentEmailTemplate.amount = formatter.format(match[TOTAL_COL]);
			tableEntries.push(finishedTable);
    }
    //Checks if the amount still owed for the fundraiser is over $0 and if so it adds it to the table for the email
		else if (match[FUNDRAISER_COLS[i]] > 0) {
			emailJS.fundraiser = FUNDRAISER_COLS[i];
			if (FUNDRAISER_COLS[i] == currentRow[PAYMENT_REASON_COL]) {
        emailJS.fundraiser_amount = formatter.format(
					match[FUNDRAISER_COLS[i]] + currentRow[PAYMENT_AMOUNT_COL]
				);
			} else {
        emailJS.fundraiser_amount = formatter.format(match[FUNDRAISER_COLS[i]]);
      }
			var finishedTable = emailJS.evaluate().getContent();
			paymentEmailTemplate.amount = formatter.format(match[TOTAL_COL]);
			tableEntries.push(finishedTable);
		}
	}
}

function sendWeeklyEmail() {
	sheet = SpreadsheetApp.getActiveSheet();
  // Gets the data from the passed sheet
	const dataRange = sheet.getDataRange();
	// Fetches values for each currentRow in the Range HT Andrew Roberts
	const data = dataRange.getValues();

	// Assumes currentRow 1 contains our column headings
	const heads = data.shift();

	// Gets the index of the column named 'Email Status' (Assumes header names are unique)
	const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

	// Converts 2d array into an object array
	const allData = data.map((tableEntries) =>
		heads.reduce((o, k, i) => ((o[k] = tableEntries[i] || ''), o), {})
	);

	// Creates an array to record sent emails
	const out = [];

	// Loops through all the rows of data
	allData.forEach(function (currentRow) {
		// Only sends emails if Send Email checkbox is checked AND if it's on the right sheet
		if (
			currentRow[SEND_EMAIL_COL] == true
		) {
			try {
				weeklyEmailTemplate.first_name = currentRow[FIRST_NAME_COL];
				weeklyEmailTemplate.last_name = currentRow[LAST_NAME_COL];
				weeklyEmailTemplate.amount = formatter.format(currentRow[TOTAL_COL]);
				buildTable(currentRow);
				//Sets the HTML variable dataRangeValues to the contents of array tableEntries which contains the table entries for the email
				weeklyEmailTemplate.weeklyTableEntries = tableEntries;
				//Stores the finished email message into the variable htmlMessage
				var htmlMessage = weeklyEmailTemplate.evaluate().getContent();
				//Sends the finished email
				GmailApp.sendEmail(
					currentRow[RECIPIENT_COL],
					'***MONEY PENDING***',
					"Your email doesn't support HTML",
					{ name: 'Student Council', htmlBody: htmlMessage }
				);
				// Edits cell to record email sent date
				out.push([formattedDate]);
				//Resets the table entries for the next email
				tableEntries.length = 0;
			} catch (e) {
				// modify cell to record error
				out.push([e.message]);
			}
		} else {
			out.push([currentRow[EMAIL_SENT_COL]]);
		}
	});
	// Updates the sheet with new data
	sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
}

function sendPaymentConfirmation() {
	sheet = SpreadsheetApp.getActiveSheet();
  // Gets the data from the passed sheet
	const dataRange = sheet.getDataRange();

	// Fetches values for each currentRow in the Range HT Andrew Roberts
	const data = dataRange.getValues();

	// Assumes currentRow 1 contains our column headings
	const heads = data.shift();

	// Gets the index of the column named 'Email Status' (Assumes header names are unique)
	const emailSentColIdx = heads.indexOf(PAYMENT_SENT_COL);

	// Converts 2d array into an object array
	const allData = data.map((tableEntries) =>
		heads.reduce((o, k, i) => ((o[k] = tableEntries[i] || ''), o), {})
	);

	// Creates an array to record sent emails
	const out = [];

	const dataRows = allData;

	// Loops through all the rows of data
	allData.forEach(function (currentRow) {
		// Only sends emails if Send Email checkbox is checked AND if it's on the right sheet
		if (
			currentRow[PAYMENT_SENT_COL] == '' &&
      currentRow[DATE_COL] != ''
		) {
			try {
				paymentEmailTemplate.first_name = currentRow[PAYER_FIRST_NAME_COL];
				paymentEmailTemplate.last_name = currentRow[PAYER_LAST_NAME_COL];
				paymentEmailTemplate.payment = formatter.format(
					currentRow[PAYMENT_AMOUNT_COL]
				);
				paymentEmailTemplate.fundraiser = currentRow[PAYMENT_REASON_COL];
				paymentEmailTemplate.officer = currentRow[OFFICER_COL];
				var date = currentRow[DATE_COL];
				paymentEmailTemplate.date =
					date.getMonth() + 1 + '/' + date.getDate() + '/' + date.getFullYear();
				buildPaymentTable(currentRow, dataRows);
				//Sets the HTML variable dataRangeValues to the contents of array tableEntries which contains the table entries for the email
				paymentEmailTemplate.paymentTableEntries = tableEntries;
				//Stores the finished email message into the variable htmlMessage
				var htmlMessage = paymentEmailTemplate.evaluate().getContent();
				var email = String(currentRow[PAYER_STUDENT_ID_COL]);
				email = email.concat('@ljisd.com');
				//Sends the finished email
				GmailApp.sendEmail(
					email,
					'***PAYMENT CONFIRMATION***',
					"Your email doesn't support HTML",
					{ name: 'Student Council', htmlBody: htmlMessage }
				);
				// Edits cell to record email sent date
				out.push([formattedDate]);
				//Resets the table entries for the next email
				tableEntries.length = 0;
			} catch (e) {
				// modify cell to record error
				out.push([e.message]);
			}
		} else {
			out.push([currentRow[PAYMENT_SENT_COL]]);
		}
	});
	// Updates the sheet with new data
	sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
}
*/