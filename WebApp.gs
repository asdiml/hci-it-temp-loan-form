function include(fileName) {
  return HtmlService.createTemplateFromFile(fileName).evaluate().getContent();
}

function doGet(e){
  const webApp = new WebApp(e.parameter);
  return webApp.serve();
}

/**
 * The primary class controlling the functionality of the script as a web application
 */
class WebApp{

  /**
   * Webapp-related Enums
   * 
   * For ease of change and to prevent bugs, wherever possible, ALL literals meant to be constant 
   * during execution should be stored as immutable enums. 
   */
  static get HtmlDocNames(){
    return Object.freeze({
      INDEX_HTML: 'index',
      APPROVALPAGE_HTML: 'approval_page.html',
      NOTFOUND_HTML: '404.html', 
      CSS: 'css.html'
    });
  }
  static get Authorisation(){
    return Object.freeze({
      AUTHORISED_USERS_AT_INIT: [
        'edtech@hci.edu.sg',
        'yamhh@hci.edu.sg',
        'tants@hci.edu.sg',
        'alfie@hci.edu.sg',
        'geraldchoo@hci.edu.sg'
      ]
    });
  }

  /**
   * Accepts an e.parameter object and constructs an instance of Webapp that manages the deployed web app's HtmlTemplate and HtmlOutput objects. 
   * 
   * @param {object} e.parameter The object derived from the query string, which allows data to be passed back to server-side code
   */
  constructor({ taskId, borrowerEmail }){
    this.app = new App();
    this.taskId = taskId;
    this.borrowerEmail = borrowerEmail;
    this.clientEmail = Session.getActiveUser().getEmail();
  }

  /**
   * High level function for overall serving of HTML pages within the web application
   * 
   * @return the HtmlOutput object for Apps Script to serve the relevant page
   */
  serve(){
    const template = this.taskId ? this.createApprovalHtmlTemplate() : this.createIndexHtmlTemplate();

    template.title = this.app.title;
    template.url = this.app.url + (this.taskId ? `?taskId=${this.taskId}` : '');
    template.loanStatusEnums = Loan.Status;
    return this.evaluateTemplate(template);
  }

  /**
   * Creates the HtmlTemplate object required to serve the Approval Page. Assumes that taskId has been passed 
   * in the query string (i.e. this.taskId exists), thus such a check should be done BEFORE calling this function. 
   * 
   * @return the HtmlTemplate object for the Approval Page IF a task of taskID exists AND the current client's email matches 
   * that of the task's approver (i.e. task.approveEmail), and the 404 page HTML Template otherwise
   */
  createApprovalHtmlTemplate(){
    let task;
    let templateFile = WebApp.HtmlDocNames.NOTFOUND_HTML;
    const loan = this.app.getLoanByTaskId(this.taskId);
    if(loan){
      task = this.app.getTaskFromLoan(loan, this.taskId);
      if(this.clientEmail === task.approveEmail) templateFile = WebApp.HtmlDocNames.APPROVALPAGE_HTML;
    }

    const template = HtmlService.createTemplateFromFile(templateFile);
    if(templateFile === WebApp.HtmlDocNames.NOTFOUND_HTML) return template;
    template.loanDetails = loan.loanDetails;
    template.loanStatus = loan.loanStatus;
    template.task = task;
    return template;
  }

  /**
   * Creates the HtmlTemplate object required to serve the index page (that allows for the searching of loans by borrowerEmail). 
   * Page rendering is not affected if the borrower's email is not passed in the query string - it is assumed (many times correctly)
   * that the user has simply not attempted a search. 
   * 
   * @return the HtmlTemplate object for the index page. If borrowerEmail is specified (i.e. not undefined), the template will render with a list of
   * matching outstanding loans as the search results. This functionality is achieved with templated HTML in index.html
   */
  createIndexHtmlTemplate(){
    if(!this.retrieveAuthorisedEmails().includes(this.clientEmail)) 
      return HtmlService.createTemplateFromFile(WebApp.HtmlDocNames.NOTFOUND_HTML);

    const loans = this.app.getLoansByBorrowerEmail(this.borrowerEmail, App.SheetNames.OUTSTANDING);
    const template = HtmlService.createTemplateFromFile(WebApp.HtmlDocNames.INDEX_HTML);
    template.loans = loans;  // If the search has not been performed (when borrowerEmail is not specified in the URL), the loans object passed in will be null
    template.borrowerEmail = this.borrowerEmail; 
    return template;
  }

  /**
   * Retrieves the list of users (by email) authorised to access the index page of the web interface
   * 
   * @return an array of authorised email addresses
   */
  retrieveAuthorisedEmails(){
    const sheet = APP_CONSTS.SS.getSheetByName(App.SheetNames.AUTHORISED_USERS);
    const range = sheet.getRange(2, 1, sheet.getLastRow()-1, 1);
    return range.getValues().flat();
  }

  /**
   * Generates the HtmlOutput object to be returned (for the webpage to be served) from a HtmlTemplate, while configuring some settings
   * Read more about Templated HTML here: https://developers.google.com/apps-script/guides/html/templates
   * 
   * @param {htmlTemplate} The HtmlTemplate object to be served
   * @return the configured HtmlOutput object to be served
   */
  evaluateTemplate(htmlTemplate){
    const htmlOutput = htmlTemplate.evaluate();
    htmlOutput.setTitle(this.app.title)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    return htmlOutput;
  }
}