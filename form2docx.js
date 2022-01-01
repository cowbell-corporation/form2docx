/**
 * @file form2docx
 * @author COWBELL Corporation <info@cowbell.jp>
 * @copyright COWBELL Corporation 2022-
 * @license The MIT License
 * @version 1.0.0
 */


/**
 * User settings (* : required)
 * @const {string} fileId              - ID of the template that is converted to Google Docs formatted document (*)
 * @const {string} adminEmail          - Administrator's email address (*)
 * @const {string} subject             - Subject of the email to the administrator
 * @const {string} mailBody            - Body of the email to the administrator (use template literals)
 * @const {string} errorSubject        - Subject of the error notification email
 * @const {string} errorMailBody       - Body of the error notification email (use template literals)
 * @const {string} suffixField         - Google form field name to be used as a suffix in the generated file
 * @const {stging} fileFormat          - Format of the generated file (docx or odt) (*)
 * @const {Array.<string>} placeHolder - Google form fields associated with the placeholder name in the template (*)
 */
const fileId        = '';
const adminEmail    = '';
const subject       = '[form2docx] Your Google Forms accepted the post.';
const mailBody      = `
Your Google Forms accepted the post.
The Google Docs document with the same content as the attached Word file is at the following URL.
`;
const errorSubject  = '[form2docx] An error has occurred.';
const errorMailBody = `
The form2docx script was not executed for the following reasons.
Check your configuration.
`;
const suffixField   = 'name';
const fileFormat    = 'docx';
const placeHolder   = [
	'name',
	'email',
	'message'
];


/**
 * Initialize
 * @const {string} templateFile - Object of the template that is converted to Google Docs format
 */
const templateFile = DriveApp.getFileById(fileId);


/**
 * Main
 * @param {Object} e - Event object
 */
 function form2docx(e) {

	try {
		// If there is no event object,
		// set the object of the first record submitted to the form as itemResponses
		// (for debugging purposes)
		const itemResponses = (e !== undefined) ? e.response.getItemResponses() : FormApp.getActiveForm().getResponses()[0].getItemResponses();

		const docId = createNewGdoc(itemResponses);
		const docx = convertGdoc2Docx(docId);
		const mailBody = createMailBody(docId);

		sendEmail(adminEmail, subject, mailBody, {attachments: docx});		
	}
	catch(e) {
		sendErrorMail(e);
	}

}


/**
 * Get date and time
 * @returns {string} String in yyyy-MM-dd_HH-mm-ss format
 */
function getDate() {

	const date = new Date();
	const formattedDate = Utilities.formatDate(date, 'JST',  'yyyy-MM-dd_HH-mm-ss');

	return formattedDate;

}


/**
 * Generate a new Google document based on the input in the form
 * @param {Object} itemResponses - Event object
 * @returns {string} ID of the generated Google document
 */
function createNewGdoc(itemResponses) {

	const newFile = templateFile.makeCopy();  
	const newId = newFile.getId();
	const newGdoc = DocumentApp.openById(newId);
	let newGdocBody = newGdoc.getBody();
	let key, value, fileSuffix;

	itemResponses.forEach(function(itemResponse) {
		key = itemResponse.getItem().getTitle();
		value = itemResponse.getResponse();
		newGdocBody = newGdocBody.replaceText(`{{${key}}}`, value);
		if(key === suffixField) {
			fileSuffix = value;
		}
	});

	newGdoc.saveAndClose();
	fileSuffix != undefined ? newFile.setName(getDate() + `_${fileSuffix}`) : newFile.setName(getDate());

	return newId;

}


/**
 * Convert the generated Google Docs to Docx format
 * @param {string} docId - ID of the generated Google document
 * @returns {Object} Generated Docx object
 */
function convertGdoc2Docx(docId) {

	const params = {
		'headers' : {
			Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
		},
		'muteHttpExceptions' : true
	};
	const filename = `${DriveApp.getFileById(docId).getName()}.${fileFormat}`;
	const url = `https://docs.google.com/document/d/${docId}/export?format=${fileFormat}`;
	const docx = UrlFetchApp.fetch(url, params).getBlob().setName(filename);

	return docx;

}


/**
 * Generate the body of the email to the administrator
 * @param {string} docId - ID of the generated Google document
 * @returns {string} The body of the email with the generated Google Docs URL appended
 */
function createMailBody(docId) {

	return mailBody + `\nhttps://docs.google.com/document/d/${docId}`;

}


/**
 * Send e-mail
 * @param {string} recipient - Recipient address
 * @param {string} subject - Subject
 * @param {string} body - Mail body
 * @param {Object} options - Sending options
 */
function sendEmail(recipient, subject, body, options) {

	GmailApp.sendEmail(
		recipient,
		subject,
		body,
		options
	);

}


/**
 * Send e-mail to the administrator when an error occurs
 * @param {string} message - Error messages inserted into the body of the email
 */
function sendErrorMail(message) {
	
	GmailApp.sendEmail(
		adminEmail,
		errorSubject,
		`${errorMailBody}\n${message}`
	);

}