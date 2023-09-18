/**
 * Designates a frozen, constant object APP_CONSTS as the container for all app-side consts
 */
const APP_CONSTS = Object.freeze({

  /**
   * Describes the task flow of the script
   * Allows for workflow modification without (hopefully) changing a lot of code
   */
  FLOWS : {
    defaultFlow: [
      {
        taskName: "Borrower Acceptance",
        sourceSheet: App.SheetNames.PENDING_ACCEPT,
        destSheet: {
          approve: App.SheetNames.OUTSTANDING,
          reject: App.SheetNames.REJECTED_BY_BORROWER
        }
      },
      {
        taskName: "Issuer Check",
        sourceSheet: App.SheetNames.OUTSTANDING,
        destSheet: {
          approve: App.SheetNames.COMPLETED,
          reject: App.SheetNames.OUTSTANDING
        }
      }
    ]
  },

  /**
   * Returns the spreadsheet to which the active form is binded to, and if it has yet to be created, 
   * creates the sheet and moves it to the directory in which the form is stored. 
   * 
   * Is a global variable so that the global function parsedValues doesn't need to initialise any class for this to work
   */
  SS : (() => {
    let responseSs;
    try{
      const responseSsID = FormApp.getActiveForm().getDestinationId();
      responseSs = SpreadsheetApp.openById(responseSsID);
    } catch (e) {
      // If no linked spreadsheet created, create the response spreadsheet and link it to the form
      responseSs = SpreadsheetApp.create(this.title + ' (Responses)');
      const responseSsId = responseSs.getId();
      FormApp.getActiveForm().setDestination(FormApp.DestinationType.SPREADSHEET, responseSsId);

      //Move the spreadsheet into the form's folder
      const formFolder = DriveApp.getFileById(form.getId()).getParents().next();
      DriveApp.getFileById(responseSsId).moveTo(formFolder);
    }
    return responseSs;
  })()
});