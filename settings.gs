// This file will set all of the program variables. Initially the values will 
// be hard-coded into this script. In the future they will be saved to a file
// and loaded at boot-time.
//

/**
 * setProp
 * 
 * Sets a script property, only if it currently has no value
 * 
 * @param {string} key Name of the property to be set
 * @param {string} value Value to be assigned to property
 * @returns void
 */
function setProp(key, value) {

  if(key === "" || key === null || key === undefined) {
    throw new Error("setProp: Null key passed to function")
  }
  if (!PropertiesService.getScriptProperties().getProperty(key)) {
    PropertiesService.getScriptProperties().setProperty(key, value);
  }
}

/**
 * getProp
 * 
 * Returns the current value of a script property
 * 
 * @param {string} key Name of the property to be returned
 * @returns string ?
 */
function getProp(key) {
    if(key === "" || key === null || key === undefined) {
    throw new Error("getProp: Null key passed to function")
  }
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * initialize
 * 
 * This function sets the script-level properties that are used
 * as global variables for the program. All of the important 
 * script parameters are set here
 * 
 * @returns void
 */

function initialize() {

  setProp('REPLY_FROM', "Don Briggs <DonBriggsWork@gmail.com>");    //-- Address replys will be sent from
  setProp('PROCESS_LABEL',"AUTO_RESUME" );                          //-- Look for threads with this label to process
  setProp('DEST_LABEL', "Followup");                                //-- List of labels to add afer message is processed
  setProp('ATTACH_FILE', "Don_Briggs_2024.pdf");                    //-- Resume file to attach to response
  setProp('REPLY_FILE', "Std_Reply");                               //-- e-mail address to reply with
  setProp('ATTACH_FILE_ID', getFileId('Don_Briggs_2024.pdf'));      //-- gDrive id of file to be attached
  setProp('REPLY_FILE_ID', getFileId("Std_Reply"));                 //-- gDrive id of document for reply template
}


/**
 * getFileId
 * 
 * Returns the gDrive Object ID of the file whos name is 
 * passed as a parametner. If the file is not found this
 * function will return an error the program will exit
 * 
 * @param {string} filename Name of the file to fetch ID for
 * @returns integer ?
 */
function getFileId(fileName) {

  if(fileName  === "" || fileName === null || fileName === undefined) {
    throw new Error("getFileId: Null filename passed to function")
  }
  
  var oFiles = DriveApp.getFilesByName(fileName);
  if(oFiles.hasNext()) {
    return fileId = oFiles.next().getId();
  }
  else { 
    throw new Error("Did not find: " + fileName);
  }
}