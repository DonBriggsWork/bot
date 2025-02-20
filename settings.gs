// This file will set all of the program variables. Initially the values will 
// be hard-coded into this script. In the future they will be saved to a file
// and loaded at boot-time.


function setProp(key, value) {
  if(key == "") {then
    throw new Error("setProp: Null key passed to function")
  }
  PropertiesService.getScriptProperties().setProperty(key, value);
}


function getProp(key) {
    if(key == "") {then
    throw new Error("getProp: Null key passed to function")
  }
  return PropertiesService.getScriptProperties().getProperty(key);
}

  // This function will set the global variables that will
  // be used by the program. It is called on Startup. If
  // any required resource is not found, we will throw an
  // error and exit


function SetVars() {

  var replyFile  = "Std_Reply";  //-- File that provides text of the reply
  var attachFile = "Don_Briggs_Resume_2024a.pdf"; //-- Resume file to be attached to reply


  setProp('REPLY_FROM', "Don Briggs <DonBriggsWork@gmail.com>"); //-- Address replys will be sent from
  setProp('PROCESS_LABEL',"AUTO_RESUME" );                       //-- Look for threads with this label to process
  setProp('DEST_LABEL', "Followup");                             //-- List of labels to add afer message is processed
  setProp('ATTACH_FILE', attachFile);                            //-- Resume file to attach to response
  setProp('REPLY_FILE ', replyFile);                             //-- e-mail address to reply with
  setProp('attachFileId', getFileId(attachFile));
  setProp('replyFileId', getFileId(replyFile));

}


  // Search drive for a file, and return it's id

function getFileId(fileName) {

  oFiles = DriveApp.getFilesByName(fileName);
  if(oFiles.hasNext()) {
    return oFiles.next().getId();
  }
  else { 
    throw new Error("Did not find: " + fileName);
  }
}