/*
Beautify the code
https://beautifier.io/
*/


/*
  Project: FEUP's mobility IN changes to LA requests (English form)
  Function: emailsCOORD_EN()
  Description: sends emails to coordinators for added and deleted course units in English language.
  Copyright: https://github.com/jppego
 */
function emailsCOORD_EN() {

	//----------------------------------------
	// ----- sheet vars
	var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // fetch the active spreadsheet

	var responsesSheet = SpreadsheetApp.getActive().getSheetByName("RAW_DATA"); // location of form records
	var changesSheet = SpreadsheetApp.getActive().getSheetByName("CHANGES_LA"); // summary of changes to LA
	var contactsSheet = SpreadsheetApp.getActive().getSheetByName("CONTACTS"); // contacts of mobility coordinators
	var countersSheet = SpreadsheetApp.getActive().getSheetByName("COUNTERS"); // records counters 
	var deletedSheet = SpreadsheetApp.getActive().getSheetByName("DELETED_UC"); // deleted course units data


	// ----- rows to process vars
	var startRow = countersSheet.getRange(2, 2).getValue(); // first row to process
	var numRows = countersSheet.getRange(4, 2).getValue(); // number of rows to process


	if (numRows == 0) return; // If there is nothing to do the script halts execution

	//----------------------------------------


	// Switch to the RAW DATA sheet to do some work.
	spreadsheet.setActiveSheet(responsesSheet); // set first

	var responsesRange = responsesSheet.getRange(startRow, 1, numRows, 44); //fetch data from "RAW DATA"
	// C 01 - Timestamp
	// C 02 - Full name
	// C 03 - U.Porto's student number
	// C 04 - Email address (e.g. up202000000@fe.up.pt)
	// C 05 - Link to your personal page in sigarra. 
	// C 06 - Home university
	// C 07 - Country of home university (not your country of origin!)
	// C 08 - Mobility period
	// C 09 - I can attend classes in the following language(s)
	// C 10 - Approved course unit #1
	// C 11 - Approved course unit #2
	// C 12 - Approved course unit #3
	// C 13 - Approved course unit #4
	// C 14 - Approved course unit #5
	// C 15 - Approved course unit #6
	// C 16 - Approved course unit #7
	// C 17 - Approved course unit #8
	// C 18 - Approved course unit #9
	// C 19 - Approved course unit #10
	// C 20 - Observations on approved LA
	// C 21 - Delete course unit #1
	// C 22 - Delete course unit #2
	// C 23 - Delete course unit #3
	// C 24 - Delete course unit #4
	// C 25 - Delete course unit #5
	// C 26 - Delete course unit #6
	// C 27 - Delete course unit #7
	// C 28 - Delete course unit #8
	// C 29 - Delete course unit #9
	// C 30 - Delete course unit #10
	// C 31 - Observations about deleted course units
	// C 32 - Add course unit #1
	// C 33 - Add course unit #2
	// C 34 - Add course unit #3
	// C 35 - Add course unit #4
	// C 36 - Add course unit #5
	// C 37 - Add course unit #6
	// C 38 - Add course unit #7
	// C 39 - Add course unit #8
	// C 40 - Add course unit #9
	// C 41 - Add course unit #10
	// C 42 - File with transcript of records, approved LA, documents submitted for your application at U.Porto's. (Max 1 PDF file up to 10MB)
	// C 43 - Comments to the mobility coordinators about your request
	// C 44 - Comments to COOP about this form


	// Fetch values for each row in the Range.
	var responsesData = responsesRange.getValues();


	//----------------------------------------
	// 

	// Switch to the CHANGES_LA sheet to do some work.
	spreadsheet.setActiveSheet(changesSheet); // set first

	var changesRange = changesSheet.getRange(startRow, 2, numRows, 6); //fetch data from "CHANGES_LA"
	// Column 01 - EMAIL COORD
	// Column 02 - EMAIL STUD
	// Column 03 - CURRENT LA
	// Column 04 - DELETED COURSES
	// Column 05 - ADDED COURSES
	// Column 06 - COORD EMAILS

	var changesData = changesRange.getValues();


	//----------------------------------------
	// 

	// Switch to the "DELETED_UC" sheet to do some work.
	spreadsheet.setActiveSheet(deletedSheet); // set first

	var deletedRange = deletedSheet.getRange(startRow, 2, numRows, 6); //fetch data from "DELETED_UC"
	// Column 01 - EMAIL COORD
	// Column 02 - EMAIL STUD
	// Column 03 - CURRENT LA
	// Column 04 - DELETED COURSES
	// Column 05 - ADDED COURSES
	// Column 06 - COORD EMAILS

	var deletedData = deletedRange.getValues();



	//----------------------------------------------------------------------------------------------------------------------------------------------------------------
	//
	// Parses the unresponded requests and sends email to mobility coordinators
	for (var i in responsesData) {

		//fetches the data of the form responses
		var row_Responses = responsesData[i]; // fetch an array with the FROM RESPONSES data from record row

		var STUD_Name = row_Responses[1]; // Student's full name | Column 02
		var STUD_Number = row_Responses[2]; // Student's number | Column 03
		var emailAddress = row_Responses[3]; // Student's email address | Column 04

		var STUD_page = row_Responses[4]; // Student's personal page | Column 05
		var STUD_University = row_Responses[5]; // Student"s home university | Column 06
		var STUD_Country = row_Responses[6]; // Student"s home country | Column 07
		var mobilityPeriod = row_Responses[7]; // Mobility period | Column 08
		var languageSTUD = row_Responses[8]; // Mobility period | Column 09


		var obsApproved = row_Responses[19]; // Observations field of approved LA | Column 20
		var obsDeleted = row_Responses[30]; // Observations field of deleted course units | Column 31

		var fileToR = row_Responses[41]; // Transcript of Records | Column 42

		var commentsCOORD = row_Responses[42]; // Comments to the coordinator | Column 43
		var commentsCOOP = row_Responses[43]; // Comments to COOP | Column 43

		// fetches the data from LA changes
		var row_Changes = changesData[i]; // fetch an array with the CHANGES_LA data from record row

		//var flagEMAIL_COORD = row_Changes[0]; // status of records (1 - answered ; 0 - not answered )
		//var flagEMAIL_STUD = row_Changes[1]; // status of records (1 - answered ; 0 - not answered )
		var approvedUC = row_Changes[2]; // list of course units previously approved
		var eliminatedUC = row_Changes[3]; // list of course units to eliminate
		var addedUC = row_Changes[4]; // list of course units to add
		var emailCOORD = row_Changes[5]; // list of coordinators' emails corresponding to the added course units


		// fetches the data from "DELETED_UC"
		var row_Deleted = deletedData[i]; // fetch an array with the CHANGES_LA data from record row

		var email_deletedCOORD = row_Deleted[5]; // list of coordinators' emails corresponding to the added course units




		//----------------------------------------------------------------------------------------------------------------------------------------------------------------
		//   
		//----------------------------------------------------------------------------------------------------------------------------------------------------------------
		//
		// Email message to mobility coordinators of added course units
		var messageCOORD = [];
		messageCOORD.push("Dear Mobility Coordinator(s), <br> <br>");

		messageCOORD.push("Student " + STUD_Name + " wants to make a change to the Learning Agreement (LA) which <b>adds course units</b> of your responsibility. <br><br>");

		messageCOORD.push("We ask you to proceed with the analysis of this LA change as soon as possible, answering to this email with it's approval/refuse. <br><br>");

		messageCOORD.push("In case you agree with the proposed changes, we will send your decision to FEUP' mobility coordinator for signature. <br><br>");

		messageCOORD.push("The Mobility IN Team (incoming@server.com). <br><br>");


		messageCOORD.push("<b>Student's data:</b><br>");
		messageCOORD.push("Name: " + STUD_Name + "<br>");
		messageCOORD.push("Student number: " + STUD_Number + "<br>");
		messageCOORD.push("Email: " + emailAddress + "<br>");
		messageCOORD.push("Sigarra's page: " + STUD_page + "<br>");
		messageCOORD.push("Univ. of origin: " + STUD_University + ", " + STUD_Country + "<br>");
		messageCOORD.push("Mobility period: " + mobilityPeriod + "<br>");
		messageCOORD.push("Language proficiency: " + languageSTUD + "<br><br>");

		messageCOORD.push("Transcript of records: <a href='" + fileToR + " ' target='_blank'>ToR</a> | <b> Only accessible to mobility coordinators</b>. <br><br>");

		messageCOORD.push("----------------------------------------------<br>");
		messageCOORD.push("<b>Added course units:</b><br>" + addedUC + "<br><br>");

		messageCOORD.push("----------------------------------------------<br>");
		messageCOORD.push("<b>Previously approved course units:</b><br>" + approvedUC + "<br>");
		messageCOORD.push("<b>Observations to approved LA:</b><br>" + obsApproved + "<br><br>");

		messageCOORD.push("----------------------------------------------<br>");
		messageCOORD.push("<b>Eliminated course units:</b><br>" + eliminatedUC);
		messageCOORD.push("<b>Observations to deleted course units:</b><br>" + obsDeleted + "<br><br>");

		messageCOORD.push("----------------------------------------------<br>");
		messageCOORD.push("<b>Comments to Coordinator(s):</b><br>" + commentsCOORD + "<br><br>");

		messageCOORD.push("----------------------------------------------<br>");
		messageCOORD.push("<b>\"Never send a human to do a machine\'s job.\"</b>, The Matrix (1999) <br><br>");




		// Combine content into a single string
		//The join() method creates and returns a new string by concatenating all of the elements in an array (or an array-like object), separated by commas or a specified separator string. If the array has only one item, then that item will be returned without using the separator.
		//https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/join
		var html_messageCOORD = messageCOORD.join('');



		// Email subject
		var subject_COORD = "Changes to the Learning Agreement - Added course units | " + STUD_Name;

		// Email addresses for cc and replyTo
		var email_CC = emailAddress + ", incoming@server.com";
		var email_replyTo = "incoming@server.com";

		// Send email to mobility Coordinator
		MailApp.sendEmail({
			from: "incoming@server.com",
			//to: emailCOORD,
			cc: email_CC,
			bcc: emailCOORD,
			replyTo: email_replyTo,
			subject: subject_COORD,
			htmlBody: html_messageCOORD
		});




		//----------------------------------------------------------------------------------------------------------------------------------------------------------------
		//   
		//----------------------------------------------------------------------------------------------------------------------------------------------------------------
		//    
		// Email message to mobility Coordinator of deleted course units

		if (email_deletedCOORD == "") { // if there are no email addresses, does nothing
		} else {
			//
			messageCOORD = [];
			messageCOORD.push("Dear Mobility Coordinator(s), <br> <br>");

			messageCOORD.push("Student " + STUD_Name + " wants to make a change to the Learning Agreement (LA) which <b>deletes course units</b> of your responsibility. <br><br>");

			messageCOORD.push("The freed vacancies can be used by other students, but only after the change to the LA is concluded. <br><br>");

			messageCOORD.push("The Mobility IN Team (incoming@server.com). <br><br>");


			messageCOORD.push("<b>Student's data:</b><br>");
			messageCOORD.push("Name: " + STUD_Name + "<br>");
			messageCOORD.push("Student number: " + STUD_Number + "<br>");
			messageCOORD.push("Email: " + emailAddress + "<br>");
			messageCOORD.push("Sigarra's page: " + STUD_page + "<br>");
			messageCOORD.push("Univ. of origin: " + STUD_University + ", " + STUD_Country + "<br>");
			messageCOORD.push("Mobility period: " + mobilityPeriod + "<br>");
			messageCOORD.push("Language proficiency: " + languageSTUD + "<br><br>");

			messageCOORD.push("Transcript of records: <a href='" + fileToR + " ' target='_blank'>ToR</a> | <b> Only accessible to mobility coordinators</b>. <br><br>");

			messageCOORD.push("----------------------------------------------<br>");
			messageCOORD.push("<b>Added course units:</b><br>" + addedUC + "<br><br>");

			messageCOORD.push("----------------------------------------------<br>");
			messageCOORD.push("<b>Previously approved course units:</b><br>" + approvedUC + "<br>");
			messageCOORD.push("<b>Observations to approved LA:</b><br>" + obsApproved + "<br><br>");

			messageCOORD.push("----------------------------------------------<br>");
			messageCOORD.push("<b>Eliminated course units:</b><br>" + eliminatedUC);
			messageCOORD.push("<b>Observations to deleted course units:</b><br>" + obsDeleted + "<br><br>");

			messageCOORD.push("----------------------------------------------<br>");
			messageCOORD.push("<b>Comments to Coordinator(s):</b><br>" + commentsCOORD + "<br><br>");

			messageCOORD.push("----------------------------------------------<br>");
			messageCOORD.push("<b>\"Never send a human to do a machine\'s job.\"</b>, The Matrix (1999) <br><br>");

			// Combine content into a single string
			//The join() method creates and returns a new string by concatenating all of the elements in an array (or an array-like object), separated by commas or a specified separator string. If the array has only one item, then that item will be returned without using the separator.
			//https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/join
			html_messageCOORD = messageCOORD.join('');



			// Email subject
			var subject_COORD = "Changes to the Learning Agreement - Deleted course units | " + STUD_Name;

			// Email addresses for cc and replyTo
			var email_CC = "incoming@server.com";
			var email_replyTo = "incoming@server.com";

			// Send email to mobility Coordinator
			MailApp.sendEmail({
				from: "incoming@server.com",
				to: email_deletedCOORD,
				cc: email_CC,
				//bcc: email_deletedCOORD,
				replyTo: email_replyTo,
				subject: subject_COORD,
				htmlBody: html_messageCOORD
			});
		}; // end if


	} // loop for



	//set the status of the requests to answered


  // https://stackoverflow.com/questions/24894648/get-today-date-in-google-appscript 
  var date = new Date();
	// uncomment the next line when ready to deploy
	var changesRange = changesSheet.getRange(startRow, 2, numRows, 1).setValue(date);


	// uncomment the next line when ready to deploy
  //var changesRange = changesSheet.getRange(startRow, 2, numRows, 1).setValue(1);


} // end function emailsCOORD_EN() 




/*
  Project: FEUP's mobility IN changes to LA requests (English form)
  Function: emailsSTUD_EN()
  Description: sends confirmation email to student in English language.
  Copyright: https://github.com/jppego
 */
function emailsSTUD_EN() {

	//----------------------------------------
	// ----- sheet vars
	var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // fetch the active spreadsheet

	var responsesSheet = SpreadsheetApp.getActive().getSheetByName("RAW_DATA"); // location of form records
	var changesSheet = SpreadsheetApp.getActive().getSheetByName("CHANGES_LA"); // summary of changes to LA
	var contactsSheet = SpreadsheetApp.getActive().getSheetByName("CONTACTS"); // contacts of mobility coordinators
	var countersSheet = SpreadsheetApp.getActive().getSheetByName("COUNTERS"); // records counters 


	// ----- rows to process vars
	var startRow = countersSheet.getRange(2, 3).getValue(); // first row to process
	var numRows = countersSheet.getRange(4, 3).getValue(); // number of rows to process


	if (numRows == 0) return; // If there is nothing to do the script halts execution

	//----------------------------------------


	// Switch to the RAW DATA sheet to do some work.
	spreadsheet.setActiveSheet(responsesSheet); // set first


	var responsesRange = responsesSheet.getRange(startRow, 1, numRows, 44); //fetch data from "RAW DATA"
	// C 01 - Timestamp
	// C 02 - Full name
	// C 03 - U.Porto's student number
	// C 04 - Email address (e.g. up202000000@fe.up.pt)
	// C 05 - Link to your personal page in sigarra. 
	// C 06 - Home university
	// C 07 - Country of home university (not your country of origin!)
	// C 08 - Mobility period
	// C 09 - I can attend classes in the following language(s)
	// C 10 - Approved course unit #1
	// C 11 - Approved course unit #2
	// C 12 - Approved course unit #3
	// C 13 - Approved course unit #4
	// C 14 - Approved course unit #5
	// C 15 - Approved course unit #6
	// C 16 - Approved course unit #7
	// C 17 - Approved course unit #8
	// C 18 - Approved course unit #9
	// C 19 - Approved course unit #10
	// C 20 - Observations on approved LA
	// C 21 - Delete course unit #1
	// C 22 - Delete course unit #2
	// C 23 - Delete course unit #3
	// C 24 - Delete course unit #4
	// C 25 - Delete course unit #5
	// C 26 - Delete course unit #6
	// C 27 - Delete course unit #7
	// C 28 - Delete course unit #8
	// C 29 - Delete course unit #9
	// C 30 - Delete course unit #10
	// C 31 - Observations about deleted course units
	// C 32 - Add course unit #1
	// C 33 - Add course unit #2
	// C 34 - Add course unit #3
	// C 35 - Add course unit #4
	// C 36 - Add course unit #5
	// C 37 - Add course unit #6
	// C 38 - Add course unit #7
	// C 39 - Add course unit #8
	// C 40 - Add course unit #9
	// C 41 - Add course unit #10
	// C 42 - File with transcript of records, approved LA, documents submitted for your application at U.Porto's. (Max 1 PDF file up to 10MB)
	// C 43 - Comments to the mobility coordinators about your request
	// C 44 - Comments to COOP about this form


	// Fetch values for each row in the Range.
	var responsesData = responsesRange.getValues();


	//----------------------------------------
	// 

	// Switch to the CHANGES_LA sheet to do some work.
	spreadsheet.setActiveSheet(changesSheet); // set first

	var changesRange = changesSheet.getRange(startRow, 2, numRows, 6); //fetch data from "CHANGES_LA"
	// Column 01 - EMAIL COORD
	// Column 02 - EMAIL STUD
	// Column 03 - CURRENT LA
	// Column 04 - DELETED COURSES
	// Column 05 - ADDED COURSES
	// Column 06 - COORD EMAILS

	var changesData = changesRange.getValues();




	//----------------------------------------------------------------------------------------------------------------------------------------------------------------
	//
	// Parses the unresponded requests and sends email to mobility coordinators
	for (var i in responsesData) {

		//fetches the data of the form responses
		var row_Responses = responsesData[i]; // fetch an array with the FROM RESPONSES data from record row
		var STUD_Name = row_Responses[1]; // Student's full name | Column 02
		var STUD_Number = row_Responses[2]; // Student's number | Column 03
		var emailAddress = row_Responses[3]; // Student's email address | Column 04

		var STUD_page = row_Responses[4]; // Student's personal page | Column 05
		var STUD_University = row_Responses[5]; // Student"s home university | Column 06
		var STUD_Country = row_Responses[6]; // Student"s home country | Column 07
		var mobilityPeriod = row_Responses[7]; // Mobility period | Column 08
		var languageSTUD = row_Responses[8]; // Mobility period | Column 09


		var obsApproved = row_Responses[19]; // Observations field of approved LA | Column 20
		var obsDeleted = row_Responses[30]; // Observations field of deleted course units | Column 31

		var fileToR = row_Responses[41]; // Transcript of Records | Column 42

		var commentsCOORD = row_Responses[42]; // Comments to the coordinator | Column 43
		var commentsCOOP = row_Responses[43]; // Comments to COOP | Column 43

		// fetches the data from LA changes
		var row_Changes = changesData[i]; // fetch an array with the CHANGES_LA data from record row

		//var flagEMAIL_COORD = row_Changes[0]; // status of records (1 - answered ; 0 - not answered )
		//var flagEMAIL_STUD = row_Changes[1]; // status of records (1 - answered ; 0 - not answered )
		var approvedUC = row_Changes[2]; // list of course units previously approved
		var eliminatedUC = row_Changes[3]; // list of course units to eliminate
		var addedUC = row_Changes[4]; // list of course units to add



		//----------------------------------------------------------------------------------------------------------------------------------------------------------------
		//   
		//----------------------------------------------------------------------------------------------------------------------------------------------------------------
		//
		// Email message to student 
		//https://stackoverflow.com/questions/10720832/line-break-in-a-message

		var messageSTUD = [];
		messageSTUD.push("Dear " + STUD_Name + ", <br> <br>");

		messageSTUD.push("Your request for a change to the Learning Agreement (LA) will be sent to the mobility coordinators during the night period.<br><br>");

		messageSTUD.push("We kindly ask you to wait for the decision. If necessary, the mobility coordinators or the Mobility IN Team from COOP - Cooperation Unit will contact you. <br><br>");

		messageSTUD.push("<b>After you have approval of all the changes</b> you requested, <b>you must insert the approved changes in the webpage where you applied for mobility</b>. Only then will we be able to have your new learning agreement signed by FEUP's mobility coordinator. <br>");

		messageSTUD.push("Please, <b>follow the procedure in the Online Application Guide</b> that you can find <a href='https://sigarra.up.pt/up/en/web_base.gera_pagina?p_pagina=122272' target='_blank'>here</a>.  <br><br>"); // link to FAQ opens in a new tab

		messageSTUD.push("The Mobility IN Team (incoming@server.com). <br><br>");


		messageSTUD.push("----------------------------------------------<br>");
		messageSTUD.push("<b>Changes requested:</b><br>");
		messageSTUD.push("Name: " + STUD_Name + "<br>");
		messageSTUD.push("Student number: " + STUD_Number + "<br>");
		messageSTUD.push("Email: " + emailAddress + "<br>");
		messageSTUD.push("Sigarra's page: " + STUD_page + "<br>");
		messageSTUD.push("Univ. of origin: " + STUD_University + ", " + STUD_Country + "<br>");
		messageSTUD.push("Mobility period: " + mobilityPeriod + "<br>");
		messageSTUD.push("Language proficiency: " + languageSTUD + "<br><br>");

		messageSTUD.push("Transcript of records: <a href='" + fileToR + " ' target='_blank'>ToR</a> | <b> Only accessible to mobility coordinators</b>.  <br><br>");

		messageSTUD.push("----------------------------------------------<br>");
		messageSTUD.push("<b>Added course units:</b><br>" + addedUC + "<br><br>");

		messageSTUD.push("----------------------------------------------<br>");
		messageSTUD.push("<b>Previously approved course units:</b><br>" + approvedUC + "<br>");
		messageSTUD.push("<b>Observations to approved LA:</b><br>" + obsApproved + "<br><br>");

		messageSTUD.push("----------------------------------------------<br>");
		messageSTUD.push("<b>Eliminated course units:</b><br>" + eliminatedUC);
		messageSTUD.push("<b>Observations to deleted course units:</b><br>" + obsDeleted + "<br><br>");

		messageSTUD.push("----------------------------------------------<br>");
		messageSTUD.push("<b>Comments to Coordinator(s):</b><br>" + commentsCOORD + "<br><br>");

		messageSTUD.push("----------------------------------------------<br>");
		messageSTUD.push("<b>Comments to COOP:</b><br>" + commentsCOOP + "<br><br>");


		// Combine content into a single string
		//The join() method creates and returns a new string by concatenating all of the elements in an array (or an array-like object), separated by commas or a specified separator string. If the array has only one item, then that item will be returned without using the separator.
		//https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/join
		var html_messageSTUD = messageSTUD.join('');


		// Email subject
		var subject_STUD = "Changes to the Learning Agreement- Your request was recorded | " + STUD_Name;

		// Email addresses for cc and replyTo
		email_CC = "incoming@server.com";
		email_replyTo = "incoming@server.com";

		// Send email to Student
		MailApp.sendEmail({
			from: "incoming@server.com",
			to: emailAddress,
			cc: email_CC, // uncomment when ready to deploy
			//bcc: ,
			replyTo: email_replyTo,
			subject: subject_STUD,
			htmlBody: html_messageSTUD
		});

	} // loop for


  // https://stackoverflow.com/questions/24894648/get-today-date-in-google-appscript 
  var date = new Date();
	// uncomment the next line when ready to deploy
  var changesRange = changesSheet.getRange(startRow, 3, numRows, 1).setValue(date);

	//set the status of the requests to answered
	// uncomment the next line when ready to deploy
	//var changesRange = changesSheet.getRange(startRow, 3, numRows, 1).setValue(1);


} // end function emailsSTUD_EN()