/**
 * The primary class used for the overall functionality of the service
 */
class App{

  /**
   * App-wide Enums
   * 
   * For ease of change and to prevent bugs, wherever possible, ALL literals meant to be constant 
   * during execution should be stored as immutable enums. 
   * 
   * Caveat: The only exception is constants used by multiple classes.
   * Those should be placed in Global_consts.gs 
   */
  static get SheetNames(){
    return Object.freeze({
      PENDING_ACCEPT: 'Pending Loans (Yet to be accepted by Borrower)',
      REJECTED_BY_BORROWER: 'Rejected Loans (Rejected by Borrower)',
      OUTSTANDING: 'Outstanding Loans',
      COMPLETED: 'Completed Loans',
      AUTHORISED_USERS: 'Authorised Users'
    });
  }
  static get OtherHeaders(){
    return Object.freeze({
      ISSUER_EMAIL: 'Email Address',
      BORROWER_EMAIL: 'Borrower Email'
    });
  }
  static get Assets(){
    return Object.freeze({
      RETURN_BUTTON_FILEID: '1fqZqSt_Bwf9FU9ozm-d4dmLIt8T7Ivh1'
    });
  }

  static resetAllIds(){
    Loan.resetId();
    Task.resetId();
  }

  constructor(){
    this.form = FormApp.getActiveForm();
    this.title = this.form.getTitle();
    this.url = 'https://script.google.com/a/macros/hci.edu.sg/s/AKfycbwPMJWklqXHzAOcmbVvQ54krUQspiYoKb44UUviIh9dZmUroH66cvAcBCLWLSreVzu0/exec';

    //Ensure that the linked sheet is named correctly
    this.sheets = APP_CONSTS.SS.getSheets();
    this.sheets.find(sheet => sheet.getFormUrl()).setName(App.SheetNames.PENDING_ACCEPT);  // Get the linked sheet
    this.initAuthorisedUsersSheet();
  }
  
  /**
   * Initialises the Authorised Users sheet using WebApp.Authorisation.AUTHORISED_USERS_AT_INIT. 
   * Can subsequently be updated from the Google sheet interface. 
   */
  initAuthorisedUsersSheet(){
    if(this.sheets.find(sheet => sheet.getName() === App.SheetNames.AUTHORISED_USERS)) return; // If already initialised, return
    
    const newSheet = APP_CONSTS.SS.insertSheet(App.SheetNames.AUTHORISED_USERS);
    const values = [['Email']].concat(WebApp.Authorisation.AUTHORISED_USERS_AT_INIT.map(user => [user]));
    newSheet.getRange(1, 1, values.length, 1).setValues(values);

    newSheet.getRange(1, 1).setFontWeight('bold');
    newSheet.setFrozenRows(1);
  }

  /**
   * Accepts a taskId and searches for the task information in across all sheets (if sheetName is unspecified), 
   * and returns an instance of Loan with property .tasks that is an array of Task objects
   * 
   * @param {string} taskId The taskId of the task to get
   * @param {string} [sheetName] Optional parameter to declare the name of the sheet within which to search for the task
   * @return the Loan object within which the queried task resides, and null if the task does not exist in the sheet/spreadsheet
   */
  getLoanByTaskId(taskId, sheetName){
    for(const key in App.SheetNames){
      const curSheetName = App.SheetNames[key];
      if(sheetName && sheetName !== curSheetName) continue;  // If sheetName is specified, skip the irrelevant sheets
      if(!this.sheets.find(sheet => curSheetName===sheet.getName())) continue;  // If the sheet in App.SheetNames has yet to be created, skip it

      const values = parsedValues(curSheetName);
      const loanRow = values.findIndex(value => value.some(cell => cell.taskId === taskId)) + 1;
      const loanRecord = values[loanRow-1];
      if(loanRecord) return new Loan(undefined, loanRow, curSheetName, loanRecord);
    }
    return null;
  }

  /**
   * Queries a loan object for a taskId, and returns the task if it exists within the loan
   * 
   * @param {object} loan The loan to be queried
   * @param {string} taskId The taskId used to identify the task
   * @return the Task object desired, and null if the task does not exist in the sheet/spreadsheet
   */
  getTaskFromLoan(loan, taskId){
    for(let i=0; i<loan.tasks.length; i++){
      if(loan.tasks[i].taskId === taskId) return loan.tasks[i];
    }
    return null;
  }

  /**
   * Accepts an email and searches across all sheets (if sheetName is unspecified) for all loans 
   * whose borrower email matches the given email across all sheets. It then returns an array of loans
   * 
   * @param {string} borrowerEmail The borrower email to match
   * @param {string} [sheetName] Optional parameter to declare the name of the sheet within which to search for the loan. The 
   * search will then be exclusive to that sheet. If unspecified, the function will search every sheet and record all results. 
   * @return an array of loans which matches the borrowerEmail provided. If there is no matching loan, an empty array will be returned. If borrowerEmail is falsey, it will return null. 
   */
  getLoansByBorrowerEmail(borrowerEmail, sheetName){
    if(!borrowerEmail) return null;
    
    const loans = [];
    for(const key in App.SheetNames){
      const curSheetName = App.SheetNames[key];
      if(sheetName && sheetName !== curSheetName) continue;

      const values = parsedValues(curSheetName);
      const borrowerEmailIndex = values[0].findIndex(cell => cell === App.OtherHeaders.BORROWER_EMAIL);

      // Creating the loan array (and extracting the loan row numbers )
      const loanRowNumsArray = [];
      const loanRecords = values.filter((value, i) => {
        if (value[borrowerEmailIndex] === borrowerEmail) return loanRowNumsArray.push(i+1);
        return false;
      });
      loans.push(...loanRecords.map((loanRecord, i) => new Loan(undefined, loanRowNumsArray[i], curSheetName, loanRecord)));
    }
    return loans;
  }

  /**
   * Creates a new sheet of name newSheetName if not created. 
   * Populates the headers with those in the main sheet named 'Pending Loans (Yet to be accepted by Borrower)'
   * 
   * @param {string} newSheetName The name of the sheet to be created
   * @return The created/desired sheet 
   */
  createNewSheetWithHeaders(newSheetName){
    const ss = APP_CONSTS.SS;
    const refSheet = ss.getSheetByName(App.SheetNames.PENDING_ACCEPT);
    const refSheetLastColNum = refSheet.getLastColumn();
    const refHeaderRange = refSheet.getRange(1, 1, 1, refSheetLastColNum);

    if(!ss.getSheetByName(newSheetName)) ss.insertSheet(newSheetName);
    const newSheet = ss.getSheetByName(newSheetName);
    newSheet.getRange(1, 1, 1, refSheetLastColNum).setValues(refHeaderRange.getDisplayValues());
    refHeaderRange.copyFormatToRange(newSheet, 1, refSheetLastColNum, 1, 1);
    newSheet.setFrozenRows(1);

    this.sheets = ss.getSheets(); // Update this.sheets
    return newSheet;
  }

  /**
   * Sends an approval-request email for the Borrower acceptance or Issuer check tasks
   * (For the issuer check task, there is no longer a need for an approval process)
   * 
   * @param {string} taskId The id of the task to be performed
   * @param {string} [querySheetName] Optional parameter specifying the name of the sheet to query with taskId
   */
  sendApproval(taskId, querySheetName){
    const loan = this.getLoanByTaskId(taskId, querySheetName);
    const task = this.getTaskFromLoan(loan, taskId);

    const template = HtmlService.createTemplateFromFile('approval_email.html');
    template.title = this.title;
    template.loanDetails = loan.loanDetails;
    template.actionUrl = `${this.url}?taskId=${taskId}`;

    const subject = task.taskName + " required - " + this.title;
    const options = {
      htmlBody: template.evaluate().getContent()
    };
    GmailApp.sendEmail(task.approveEmail, subject, "", options);
  }
  
  /**
   * Upon acceptance or rejection of a task, notify the other party of the decision via email
   * 
   * @param {string} taskId The id of the task performed
   * @param {string} [querySheetName] Optional parameter specifying the name of the sheet to query with taskId
   */
  sendNotification(taskId, querySheetName){
    const loan = this.getLoanByTaskId(taskId, querySheetName);
    const task = this.getTaskFromLoan(loan, taskId);

    const template = HtmlService.createTemplateFromFile('notification_email.html');
    template.title = this.title;
    template.loanDetails = loan.loanDetails;
    template.taskStatus = task.status;
    template.taskStatusEnums = Task.Status;
    template.taskName = task.taskName;

    const subject = task.taskName === Task.Names.BORROWER_ACCEPTANCE ?
      `${task.taskName.split(' ')[0]} ${task.status} - ${this.title}` :
      `Loan Successfully Returned - ${this.title}`;
    const options = {
      htmlBody: template.evaluate().getContent()
    };
    GmailApp.sendEmail(task.notifEmail, subject, "", options);
  }

  onFormSubmit(){
    const values = parsedValues(App.SheetNames.PENDING_ACCEPT);
    const headers = values[0];
    let lastRow = values.length;
    let startOwnHeadersColumn = headers.indexOf(Loan.Id.HEADER) + 1;
    if(!startOwnHeadersColumn) startOwnHeadersColumn = headers.length + 1;
    
    const newHeaders = [Loan.Id.HEADER, Loan.Status.HEADER];
    const newValues = [Loan.createId(), Loan.Status.Pending.PENDING_ACCEPT];

    const flow = APP_CONSTS.FLOWS.defaultFlow;
    let currentTaskId;
    flow.forEach((flowObj, i) => {

      // Assigning task-specific property values. task[0] is Borrower Acceptance, task[1] is Issuer Check
      const status = Task.Status[i===0 ? 'PENDING' : 'WAITING'];
      const approveEmail = values[lastRow-1][headers.indexOf(App.OtherHeaders[i===0 ? 'BORROWER_EMAIL' : 'ISSUER_EMAIL'])];
      const notifEmail = values[lastRow-1][headers.indexOf(App.OtherHeaders[i===1 ? 'BORROWER_EMAIL' : 'ISSUER_EMAIL'])];

      newHeaders.push(Task.Headers[i]);
      newValues.push(JSON.stringify(new Task(flowObj, status, approveEmail, notifEmail)));
      if(i === 0) currentTaskId = JSON.parse(newValues[newValues.length-1]).taskId;
    });

    const sheet = APP_CONSTS.SS.getSheetByName(App.SheetNames.PENDING_ACCEPT);
    sheet.getRange(1, startOwnHeadersColumn, 1, newHeaders.length)
      .setValues([newHeaders])
      .setBackground("#34A853")
      .setFontColor("#FFFFFF");
    sheet.getRange(lastRow, startOwnHeadersColumn, 1, newValues.length).setValues([newValues]);

    this.sendApproval(currentTaskId, App.SheetNames.PENDING_ACCEPT);
  }

  /**
   * Moves a loanRecord (i.e. a row of data in a sheet) to the specific destination sheet, and
   * deletes the loanRecord on the source sheet. If the source and destination sheet are the same, 
   * it returns without performing any action. 
   * 
   * @param {array} loanRecord 2d array representation of the row of data to be transferred
   * @param {number} oldRowNum Row Number of the record in the source sheet
   * @param {sheet} sourceSheet Source sheet of the loanRecord
   * @param {string} destSheetName Name of the destination sheet
   */
  moveToSheet(loanRecord, oldRowNum, sourceSheet, destSheetName){
    if(sourceSheet.getName() === destSheetName) return;
    const destSheet = this.createNewSheetWithHeaders(destSheetName);
    destSheet.getRange(destSheet.getLastRow()+1, 1, 1, loanRecord[0].length).setValues(loanRecord);
    sourceSheet.deleteRow(oldRowNum);
  }

  approve({ taskId, comments }){
    const loan = this.getLoanByTaskId(taskId, App.SheetNames.PENDING_ACCEPT);
    const task = this.getTaskFromLoan(loan, taskId);
    if(!task) throw new Error(`${taskName} approval error`);

    const sourceSheet = APP_CONSTS.SS.getSheetByName(task.sourceSheet);
    const { taskName, rowNum: taskRowNum, colNum: taskColNum } = task;
    console.log(taskRowNum, taskColNum);
    delete task.rowNum;
    delete task.colNum;
    
    task.comments = comments;
    task.status = Task.Status.COMPLETED;
    task.timestamp = new Date();
    sourceSheet.getRange(taskRowNum, taskColNum).setValue(JSON.stringify(task));
    sourceSheet.getRange(taskRowNum, loan.loanStatusColNum)
      .setValue(Loan.Status.Accepted[
        taskName === Task.Names.BORROWER_ACCEPTANCE ?
          'ACCEPTED_BORROWER' : 'ACCEPTED_ISSUER'
      ]);
    
    const loanRecord = sourceSheet.getRange(taskRowNum, 1, 1, sourceSheet.getLastColumn()).getDisplayValues();
    this.moveToSheet(loanRecord, taskRowNum, sourceSheet, task.destSheet.approve);
    this.sendNotification(taskId, task.destSheet.approve);
  }

  reject({ taskId, comments }){
    const loan = this.getLoanByTaskId(taskId, App.SheetNames.PENDING_ACCEPT);
    const task = this.getTaskFromLoan(loan, taskId);
    if(!task) throw new Error(`${taskName} rejection error`);

    const sourceSheet = APP_CONSTS.SS.getSheetByName(task.sourceSheet);
    const { taskName, rowNum: taskRowNum, colNum: taskColNum } = task;
    delete task.rowNum;
    delete task.colNum;

    task.comments = comments;
    task.status = Task.Status.REJECTED;
    task.timestamp = new Date();
    sourceSheet.getRange(taskRowNum, taskColNum).setValue(JSON.stringify(task));
    sourceSheet.getRange(taskRowNum, loan.loanStatusColNum)
      .setValue(Loan.Status[
        taskName === Task.Names.BORROWER_ACCEPTANCE ?
          'Rejected' : 'Accepted'
      ][
        taskName === Task.Names.BORROWER_ACCEPTANCE ?
          'REJECTED_BORROWER' : 'ACCEPTED_BORROWER'
      ]);

    const loanRecord = sourceSheet.getRange(taskRowNum, 1, 1, sourceSheet.getLastColumn()).getDisplayValues();
    this.moveToSheet(loanRecord, taskRowNum, sourceSheet, task.destSheet.reject);
    this.sendNotification(taskId, task.destSheet.reject);
  }

  returnLoan(loanRow){
    const sourceSheet = APP_CONSTS.SS.getSheetByName(App.SheetNames.OUTSTANDING);
    const loan = new Loan(undefined, loanRow, App.SheetNames.OUTSTANDING);
    const task = loan.tasks[1];
    const taskColNum = task.colNum;
    delete task.rowNum;
    delete task.colNum;

    task.status = Task.Status.COMPLETED;
    task.timestamp = new Date();
    sourceSheet.getRange(loanRow, taskColNum).setValue(JSON.stringify(task));
    sourceSheet.getRange(loanRow, loan.loanStatusColNum)
      .setValue(Loan.Status.Accepted.ACCEPTED_ISSUER);

    const loanRecord = sourceSheet.getRange(loanRow, 1, 1, sourceSheet.getLastColumn()).getDisplayValues();
    this.moveToSheet(loanRecord, loanRow, sourceSheet, App.SheetNames.COMPLETED);
    this.sendNotification(task.taskId, App.SheetNames.COMPLETED);
  }
}


function approve({ taskId, comments }){
  const app = new App();
  app.approve({ taskId, comments });
}

function reject({ taskId, comments }){
  const app = new App();
  app.reject({ taskId, comments });
}

function returnLoan(loanRow){
  const app = new App();
  app.returnLoan(loanRow);
}

function _onFormSubmit(){
  const app = new App();
  app.onFormSubmit();
}

function resetAllIds(){
  App.resetAllIds();
}