// This file will set all of the program variables.
// Initially the values will be hard-coded into this
// script. In the future they will be saved to a file
// and loaded at boot-time.



function setProp(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

function getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

  // This function will set the global variables that will
  // be used by the program. It is called on Startup. If
  // any required resource is not found, we will throw an
  // error and exit

function SetVars() {

  var replyFile  = "Std Reply 2022"
  var attachFile = "Don_Briggs_Resume_2022.pdf";

  Logger.log("Setting up processing parameters");
  setProp('PROCESS_LABEL',"AUTO_RESUME" );            //-- Look for threads with this label to process
  setProp('DEST_LABEL', "Followup");                  //-- List of lables to add afer message is processed
  setProp('ATTACH_FILE', attachFile);                 //-- Resume file to attach to response
  setProp('REPLY_FILE ', replyFile);                  //-- e-mail remplage to reply with
  setProp('attachFileId', getFileId(attachFile));
  setProp('replyFileId', getFileId(replyFile));
}


  // Search drive for a file, and return it's id
function getFileId(fileName) {

  oFiles = DriveApp.getFilesByName(fileName);
  if(oFiles.hasNext()) {
    Logger.log("File found");
    return oFiles.next().getId();
  }
  else { 
    throw new Error("Did not find: " + fileName);
  }
}