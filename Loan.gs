/**
 * The reference class for all loan objects. Also contains loan-related enums
 */
class Loan{

  /**
   * Loan-related Enums
   * 
   * For ease of change and to prevent bugs, wherever possible, ALL literals meant to be constant 
   * during execution should be stored as immutable enums. 
   */
  static get Id(){
    return Object.freeze({
      HEADER: '_loanId',
      PREFIX: 'L-',
      LENGTH: 6
    });
  }
  static get Status(){
    return Object.freeze({
      HEADER: '_status',
      Pending: {
        PENDING_ACCEPT: 'Pending Acceptance',
        PENDING_RETURNACK: 'Pending Return Acknowledgement',
      },
      Accepted: {
        ACCEPTED_BORROWER: 'Outstanding',   // Once accepted by the borrower, the loan is now outstanding. 
        ACCEPTED_ISSUER: 'Completed'        // Once the issuer has completed checks on the equipment, the loan procedure will be completed.
      },
      Rejected: {
        REJECTED_BORROWER: 'Rejected by Borrower',
      }
    });
  }

  /**
   * Tracks and increments the loanId stored in PropertiesService
   * Generates the new Id string literal that includes the Id prefix
   * 
   * @return The loanId string literal
   */
  static createId(){
    const props = PropertiesService.getDocumentProperties();
    const loanIdEnumObj = Loan.Id;
    let loanId = Number(props.getProperty(loanIdEnumObj.HEADER));
    if(!loanId) loanId = 1;

    props.setProperty(loanIdEnumObj.HEADER, loanId + 1);
    return loanIdEnumObj.PREFIX + (loanId + 10 ** loanIdEnumObj.LENGTH).toString().slice(-loanIdEnumObj.LENGTH);
  }

  /**
   * Reset the PropertiesService value of Loan.Id.HEADER to 0, thereby resetting the loanId
   */
  static resetId(){
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(Loan.Id.HEADER);
  }
  
  /**
   * Constructs an instance of Loan by retrieving it based on either its id or loan record (i.e. row data) in a sheet. 
   * If its loan record and sheet name are specified, the constructor will choose to use those over loanId. 
   * For unused params, undefined must be passed as a filler. 
   * 
   * @param {string} [loanId] The id of the loan to get. If loanId is not specified (i.e. undefined), loanRecord AND sheetName must be specified. 
   * @param {number}: The row at which the loan is located in sheetName. Mandatory if loanId is not specified
   * @param {string} [sheetName] The name of the sheet within which to find the loan. Mandatory if loanId is not specified, and if only loanId is specified, optimises the search
   * @param {array} [loanRecord] 1d array of parsed data as it appears on the spreadsheet. If specified, cuts away the time required to parse loanRow in sheetName
   * @return returns an instance of the Loan class. If loan of loanId does not exist, returns null and logs an error to the console
   */
  constructor(loanId, loanRow, sheetName, loanRecord){
    // If loanId is not specified, BOTH loanRecord and sheetName should be specified. 
    if(!loanId && (!loanRow || !sheetName)) throw Error('Invalid parameter combination/input');

    // Retrieves/finds the loanRecord as well as the sheet's headers
    let headers;
    if(sheetName){
      const sheet = APP_CONSTS.SS.getSheetByName(sheetName);
      headers = parsedValues(sheetName, 1)[0];
      if(!loanRecord) loanRecord = parsedValues(sheetName, loanRow)[0];
    }
    else if(loanId){
      try {
        for(const key in App.SheetNames){
          const curSheetName = App.SheetNames[key];
          if(sheetName && sheetName !== curSheetName) continue;
          
          const values = parsedValues(curSheetName);
          loanRow = values.findIndex(value => value.some(cell => cell.taskId === taskId)) + 1;
          if(!loanRow) throw Error(`Loan of id ${loanId} does not exist`);
          loanRecord = values[loanRow-1];
          headers = values[0];
        }
      } catch(err){
        console.log(err);
        return null;
      }
    }
    else throw Error('Internal error during Loan construction');

    // Sets properties of Loan based on its record
    this.loanStatusColNum = headers.indexOf(Loan.Status.HEADER) + 1;
    this.loanDetails = loanRecord.slice(0, this.loanStatusColNum).map((loanDetail, i) => {
      return {
        label: headers[i], 
        value: loanDetail
      }
    });
    this.loanStatus = loanRecord[this.loanStatusColNum - 1];
    this.loanId = loanRecord[headers.indexOf(Loan.Id.HEADER)];
    this.loanRow = loanRow;
    
    // Adds additional data into each task object and inserts them into an array
    this.tasks = loanRecord.slice(this.loanStatusColNum);
    this.tasks.forEach((task, index) => {
      task.rowNum = loanRow;
      task.colNum = this.loanStatusColNum + index + 1;
    });
  }
}