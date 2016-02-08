/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{name: 'Send Sign-Up Form', functionName: 'sendSignUp_'}, {name: 'Send Confirmation Form', functionName: 'sendConfirmation_'}, {name: 'Move Wait List', functionName: 'moveWaitlist_'}];
  SpreadsheetApp.getActive().addMenu('RSVPlease', menu);
  PropertiesService.getScriptProperties().setProperty('signUpFormID', '');
  PropertiesService.getScriptProperties().setProperty('nameColumn', 0);
  PropertiesService.getScriptProperties().setProperty('descriptionColumn', 1);
  PropertiesService.getScriptProperties().setProperty('contentColumn', 2);
  PropertiesService.getScriptProperties().setProperty('confirmationQuestionAdded', false);
  PropertiesService.getScriptProperties().setProperty('numConfirmationsSent', 0);
  Logger.log("Went through startup stuff");
}

/**
 * A function that builds a form
 */
function buildForm_(rowNum, values, signUpForm) {
  var nameColumn = parseInt(PropertiesService.getScriptProperties().getProperty('nameColumn'));
  var descriptionColumn = parseInt(PropertiesService.getScriptProperties().getProperty('descriptionColumn'));
  var contentColumn = parseInt(PropertiesService.getScriptProperties().getProperty('contentColumn'));
  var confirmationQuestionAdded = PropertiesService.getScriptProperties().getProperty('confirmationQuestionAdded');
  var numConfirmationsSent = parseInt(PropertiesService.getScriptProperties().getProperty('numConfirmationsSent'));
  while (values[rowNum][descriptionColumn] === "Short answer" || values[rowNum][descriptionColumn] === "Long answer" || values[rowNum][descriptionColumn] === "Multiple choice") {
    if (values[rowNum][descriptionColumn] === "Short answer") {
      Logger.log("Creating short answer question: " + values[rowNum][contentColumn]);
      shortQuestion = signUpForm.addTextItem();
      shortQuestion.setTitle(values[rowNum][contentColumn]);
      shortQuestion.setRequired(true);
    }
    else if (values[rowNum][descriptionColumn] === "Long answer") {
      Logger.log("Creating long answer question: " + values[rowNum][contentColumn]);
      longQuestion = signUpForm.addParagraphTextItem();
      longQuestion.setTitle(values[rowNum][contentColumn]);
      longQuestion.setRequired(true);
    }
    else if (values[rowNum][descriptionColumn] === "Multiple choice") {
      Logger.log("Creating multiple choice question: " + values[rowNum][contentColumn]);
      var mcQuestion = signUpForm.addMultipleChoiceItem();
      mcQuestion.setRequired(true);
      var colNum = contentColumn;
      mcQuestion.setTitle(values[rowNum][colNum]);
      colNum += 1;
      while (values[rowNum][colNum] != "") {
        Logger.log("Adding multiple choice: "+ values[rowNum][colNum]);
        mcQuestion.createChoice(values[rowNum][colNum]);
        colNum += 1;
      }
    }
    rowNum += 1;
  }
}

/**
 * A function that returns the row that starts with the specified cell contents.
 */
function getRowWithItem_(values, itemName) {
  var nameColumn = parseInt(PropertiesService.getScriptProperties().getProperty('nameColumn'));
  var rowNum = 0;
  while (values[rowNum][nameColumn] != itemName) {
    rowNum += 1;
  }
  return rowNum;
}
function getNextRowWithDescription_(values, itemDescription) {
  var descriptionColumn = parseInt(PropertiesService.getScriptProperties().setProperty('descriptionColumn', 1));
  var rowNum = 0;
  while (values[rowNum][descriptionColumn] != itemDescription) {
    rowNum += 1;
  }
  return rowNum;
}
function whichQuestion_(form,questionTitle) {
  var items = form.getItems();
  for (var item = 0; item < items.length; item++) {
    if (items[item].getTitle == questionTitle) {
      return item;
    }
  }
}
function getQuestion_(form,questionTitle) {
  var items = form.getItems();
  for (var item = 0; item < items.length; item++) {
    if (items[item].getTitle == questionTitle) {
      return items[item];
    }
  }
}

/**
 * A set-up function that uses the data in the spreadsheet to create
 * a sign-up form and send it out to the emails in the list.
 */
function sendSignUp_() {
  var nameColumn = parseInt(PropertiesService.getScriptProperties().getProperty('nameColumn'));
  var descriptionColumn = parseInt(PropertiesService.getScriptProperties().getProperty('descriptionColumn'));
  var contentColumn = parseInt(PropertiesService.getScriptProperties().getProperty('contentColumn'));
  var confirmationQuestionAdded = PropertiesService.getScriptProperties().getProperty('confirmationQuestionAdded');
  var numConfirmationsSent = parseInt(PropertiesService.getScriptProperties().getProperty('numConfirmationsSent'));
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Data');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var rowNum = getRowWithItem_(values, "Invitation questions");
  var title = values[rowNum][contentColumn];
  rowNum += 1;
  var description = values[rowNum][contentColumn];
  var signUpForm = FormApp.create(title);
  PropertiesService.getScriptProperties().setProperty('signUpFormID', signUpForm.getId());
  Logger.log("Initialized sign up form");
  signUpForm.setDescription(description);

  // FILL IN QUESTIONS
  rowNum += 1;
  buildForm_(rowNum, values, signUpForm);
  Logger.log("Built sign up form");
  
  // WRITE EMAIL
  rowNum = getRowWithItem_(values, "Invitation email");
  var formURL = signUpForm.shortenFormUrl(signUpForm.getPublishedUrl());
  var subject = values[rowNum][contentColumn];
  rowNum += 1;
  var message = values[rowNum][contentColumn];
  message += " " + formURL + "\n";
  rowNum += 1;
  message += values[rowNum][contentColumn];
  rowNum += 1;
  // SEND EMAIL TO EMAIL LIST
  rowNum = getRowWithItem_(values, "Emails");
  while (rowNum < values.length && values[rowNum][contentColumn] != "") {
    Logger.log("Sending email to " + values[rowNum][contentColumn]);
    MailApp.sendEmail(values[rowNum][contentColumn], subject, message);
    rowNum += 1;
  }
  Logger.log("Done sending sign up emails");
}

/**
 * A set-up function that uses the data in the spreadsheet to create
 * a confirmation form and send it out to the first emails in the list.
 */
function sendConfirmation_() {
  var signUpForm = FormApp.openById(PropertiesService.getScriptProperties().getProperty('signUpFormID'));
  var nameColumn = parseInt(PropertiesService.getScriptProperties().getProperty('nameColumn'));
  var descriptionColumn = parseInt(PropertiesService.getScriptProperties().getProperty('descriptionColumn'));
  var contentColumn = parseInt(PropertiesService.getScriptProperties().getProperty('contentColumn'));
  var confirmationQuestionAdded = PropertiesService.getScriptProperties().getProperty('confirmationQuestionAdded');
  var numConfirmationsSent = parseInt(PropertiesService.getScriptProperties().getProperty('numConfirmationsSent'));
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Data');
  var range = sheet.getDataRange();
  var values = range.getValues();
  // GET CONFIRMATION EMAIL DETAILS
  Logger.log("Creating confirmation email");
  var rowNum = getRowWithItem_(values, "Confirmation email");
  var subject = values[rowNum][contentColumn];
  Logger.log("Confirmation email subject: " + subject);
  rowNum += 1;
  var messagePart1 = values[rowNum][contentColumn];
  Logger.log("Confirmation email messagePart1: " + messagePart1);
  rowNum += 1;
  var numSeats = values[getRowWithItem_(values, "Number of seats")][contentColumn];
  Logger.log("Confirmation email numSeats: " + numSeats);
  var messagePart2 = values[rowNum][contentColumn];
  Logger.log("Confirmation email messagePart2: " + messagePart2);
  rowNum += 1;
  message += " " + numSeats + " ";
  rowNum += 1;
  message += values[rowNum][contentColumn];
  Logger.log("Confirmation email message: " + message);
  // SEND EMAIL TO PEOPLE WHO SIGNED UP
  var signups = signUpForm.getResponses();
  var emailQuestion = whichQuestion_(signUpForm,"Email");
  var numConfirmations = -1;
  if (signups.length < numSeats) {
    numConfirmations = signups.length;
  }
  else {
    numConfirmations = numSeats;
  }
  // ADD CONFIRMATION QUESTION
  if (confirmationQuestionAdded == false) {
    rowNum = getRowWithItem_(values, "Confirmation question");
    buildForm_(rowNum, values, signUpForm);
    PropertiesService.getScriptProperties().setProperty('confirmationQuestionAdded', true);
    Logger.log("Added confirmation question to sign up form");
  }
  for (var response = 0; response < numConfirmations; response ++) {
    // SEND EMAIL
    var answers = signups[response].getItemResponses();
    var formURL = signups[response].getEditResponseUrl();
    var message = messagePart1 + " " + formURL + "\n" + messagePart2;
    Logger.log("Sending confirmation email to " + answers[emailQuestion]);
    MailApp.sendEmail(answers[emailQuestion], subject, message);
  }
  PropertiesService.getScriptProperties().setProperty('numConfirmationsSent', numConfirmations);
  Logger.log("Finished sending " + numConfirmations + " confirmation emails");
}

/**
 * A function that marks people who have not filled out the confirmation as
 * Not Coming, and sends out confirmations to people on the wait list.
 */
function moveWaitlist_() {
  var signUpForm = FormApp.openById(PropertiesService.getScriptProperties().getProperty('signUpFormID'));
  var nameColumn = parseInt(PropertiesService.getScriptProperties().getProperty('nameColumn'));
  var descriptionColumn = parseInt(PropertiesService.getScriptProperties().getProperty('descriptionColumn'));
  var contentColumn = parseInt(PropertiesService.getScriptProperties().getProperty('contentColumn'));
  var confirmationQuestionAdded = PropertiesService.getScriptProperties().getProperty('confirmationQuestionAdded');
  var numConfirmationsSent = parseInt(PropertiesService.getScriptProperties().getProperty('numConfirmationsSent'));
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Data');
  var range = sheet.getDataRange();
  var values = range.getValues();
  // SEND EMAIL TO WAITLISTED PEOPLE UNTIL SEATS ARE FILLED
  var signups = signUpForm.getResponses();
  var emailQuestion = whichQuestionIsEmail_(signUpForm);
  var numConfirmations = 0;
  var numSeats = values[getRowWithItem_(values, "Number of seats")][contentColumn];
  var response = 0;
  // ADD CONFIRMATION QUESTION
  if (confirmationQuestionAdded == false) {
    rowNum = getRowWithItem_(values, "Confirmation question");
    buildForm_(rowNum, values, signUpForm);
    confirmationQuestionAdded = true;
  }
  var rowNum = getRowWithItem_(values, "Confirmation email");
  var subject = values[rowNum][contentColumn];
  rowNum += 1;
  var messagePart1 = values[rowNum][contentColumn];
  rowNum += 1;
  var numSeats = values[getRowWithItem_(values, "Number of seats")][contentColumn];
  var messagePart2 = values[rowNum][contentColumn];
  rowNum += 1;
  message += " " + numSeats + " ";
  rowNum += 1;
  message += values[rowNum][contentColumn];
  var confirmQuestion = getQuestion_(signUpForm, values[getRowWithItem_(values, "Confirmation question")][contentColumn]);
  var emailQuestion = whichQuestion_(signUpForm,"Email");
  while (numConfirmations < numSeats && response < signups.length) {
    if (response < numConfirmationsSent) {
      var answer = signups[response].getResponseForItem(confirmationQuestion);
      if (answer != null && answer.getResponse() == values([getRowWithItem_(values, "Confirmation question")][contentColumn + 1])) {
        numConfirmations += 1;
      }
    }
    else {
      // student was wait listed and needs to be sent confirmation email
      numConfirmations += 1;
      var answers = signups[response].getItemResponses();
      var formURL = signups[response].getEditResponseUrl();
      var message = messagePart1 + " " + formURL + "\n" + messagePart2;
      MailApp.sendEmail(answers[emailQuestion], subject, message);
    }
    response += 1;
  }
  numConfirmationsSent = response;
}

// TODO: automatically send a confirmation form to a wait listed person when someone says they cannot attend
