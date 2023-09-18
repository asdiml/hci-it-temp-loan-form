/**
 * The reference class for all task objects. Also contains task-related enums
 */
class Task{

  /**
   * Task-related Enums
   * 
   * For ease of change and to prevent bugs, wherever possible, ALL literals meant to be constant 
   * during execution should be stored as immutable enums. 
   */
  static get Id(){
    return Object.freeze({
      HEADER: '_taskId',
      PREFIX: 'T-',
      LENGTH: 7
    });
  }
  static get Headers(){
    return Object.freeze([
      '_borrowerAcceptance',
      '_issuerCheck'
    ]);
  }
  static get Status(){
    return Object.freeze({
      PENDING: 'Pending',
      COMPLETED: 'Accepted',
      REJECTED: 'Rejected',
      WAITING: 'Waiting'
    });
  }
  static get Names(){
    return Object.freeze({
      BORROWER_ACCEPTANCE: APP_CONSTS.FLOWS.defaultFlow[0].taskName,
      ISSUER_CHECK: APP_CONSTS.FLOWS.defaultFlow[1].taskName
    });
  }

  /**
   * Tracks and increments the taskId stored in PropertiesService
   * Generates the new Id string literal that includes the Id prefix
   * 
   * @return The taskId string literal
   */
  static createId(){
    const props = PropertiesService.getDocumentProperties();
    const taskIdEnumObj = Task.Id;
    let taskId = Number(props.getProperty(taskIdEnumObj.HEADER));
    if(!taskId) taskId = 1;

    props.setProperty(taskIdEnumObj.HEADER, taskId + 1);
    return taskIdEnumObj.PREFIX + (taskId + 10 ** taskIdEnumObj.LENGTH).toString().slice(-taskIdEnumObj.LENGTH);
  }

  /**
   * Reset the PropertiesService value of Task.Id.HEADER to 0, thereby resetting the taskId
   */
  static resetId(){
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(Task.Id.HEADER);
  }
  
  /**
   * Constructs an instance of Task from the array element in the selected flow, and takes in additional parameters
   * 
   * @param {object} flowObj The array element of type object in the element flow that provides the object's base properties
   * @param {string} status The task status (either PENDING or WAITING)
   * @param {number} approveEmail The email address to which the acceptance request should be sent
   * @param {string} notifEmail The email address to which a notification should be sent upon acceptance/rejection
   */
  constructor(flowObj, status, approveEmail, notifEmail){
    
    // Specifies how the deep cloning of flowObj will occur
    this.deepClone = function(obj, ifRoot){
      const keyArray = Object.keys(obj);
      const returnObj = {};
      for(let i=0; i<keyArray.length; i++){
        const key = keyArray[i];
        const value = obj[key];
        const processedValue = typeof value === 'object' ? this.deepClone(value, 0) : value;

        if(ifRoot) this[key] = processedValue;
        else returnObj[key] = processedValue;
      }
      return returnObj;
    };

    // Executes the deep cloning and deletes the function after
    this.deepClone(flowObj, 1);
    delete this.deepClone;

    // Sets the other properties of Task
    this.status = status;
    this.approveEmail = approveEmail;
    this.notifEmail = notifEmail;
    this.comments = null;
    this.taskId = Task.createId();
    this.timestamp = new Date();
  }
}