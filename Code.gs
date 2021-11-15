// @ts-nocheck
/**
 * ResumeBot
 * 
 * @author DonBriggs <donBriggsWork@gmail.com>
 * @version 0.1
 * 
 * Search through inbox for messages with a certian label. Reply to
 * them with a standard template, attach resume, mark them as read, 
 * and move them to the "Jobs" label.
 */


function main() {

  const DEBUG = 0;

  //-- Job Configuration Blocks
  // const PROCESS_LABEL = "AUTO_TRACK";             //-- Look for threads with this label to process
  // const DEST_LABEL    = "Followup";               //-- Label to move threads to after processed
  // const ATTACH_FILE   = "Don!s Resume 2021c.pdf";
  // const REPLY_FILE    = "Submission Followup";

  const PROCESS_LABEL = "AUTO_RESUME";                //-- Look for threads with this label to process
  const DEST_LABEL    = "Followup";                   //-- Label to move threads to after processed
  const ATTACH_FILE   = "Don!s Resume 2021c.pdf";
  const REPLY_FILE    = "Job Lead Reply";

  var thread = null;
  var msgId = null;
  var threadId;
  var oMsg = null;
  var replyText = null;    //-- Text to be used for reply
  var oAttachment = null;  //-- FileObject to be attached to reply
  var replyId = null;

  //--- Open users mailbox, and get tagged threads to process

  var threads = getThreads(PROCESS_LABEL);
  if (threads.length == 0) {
    Logger.log("No threads to process. Program exiting.");
    return new Array();
  }
  Logger.log("Found Threads: " + threads.length);


  //--- Setup text to use in reply
  var oReplyFile = getFileObj(REPLY_FILE);
  if (oReplyFile == null) {
    Logger.log("ERROR: Could not open Reply text file. Program exiting.");
    return new Array();
  }
  else {
     Logger.log("Located response template: " + REPLY_FILE);
     replyText = oReplyFile.getBody().getText();
  }

  //-- Setup file to attach to reply
  oAttachment = getFileObj(ATTACH_FILE, MimeType.PDF);
  Logger.log("Attachment: " + ATTACH_FILE);

  //--- Begin main processing ---//

  var i = 1;
  for (let thread of threads) {
    Logger.log("  Thread: " +i + "/" + threads.length);
    thread.moveToArchive();
    thread = clearLabels(thread);
    setLabel(thread, DEST_LABEL);
    msgId = thread.getMessages()[0].getId();
    Logger.log("Subject:" + thread.getMessages()[0].getSubject());
    oMsg = GmailApp.getMessageById(msgId);
    reply = getReply(oMsg, replyText, oAttachment);

    if (DEBUG === 0) {
      reply.send()
    } else {
      if (i >= 3) {
        Logger.log("Breaking after 3htee messages for DEBUB mode");
        break;
      }
    }
    i++;
  }
}

  //-------------------------------------------------------------------------------------
  //-----------------------------------   FUNCTIONS   -----------------------------------
  //-------------------------------------------------------------------------------------

/**
 * getSheet
 * 
 * Gets a connection to the spreadsheet holding rules data. Note
 * that the rules spreadsheet should be in the same directory as
 * the project files.
 * 
 * @param {string} fileName Name of the spreadsheet file
 * @return {object} GoogleSheets Spreadsheet object
 */

function getSheet(fileName){
  
  var id = null;
  var files = DriveApp.getFiles();
  while (files.hasNext() && id == null) {
    var file = files.next();
    if (file.getName() == fileName) {
      id = file.getId();
    }
  }
  var ss = SpreadsheetApp.openById(id);
  return ss;
}

/**
 * getThreads
 * 
 * Retuns an array of Thread objects with a particular label
 * 
 * @param {string} strCriteria Label gmail search criteria string
 * @return array of Thread objects that match search criteria
*/

function getThreads(strLabel){
  Logger.log("Search: "  + strLabel);
  var threads = GmailApp.search('label: ' + strLabel);
  if (threads.length > 0) {
    return threads;
  } else {
    return new Array();
  };
}

/**
 * clearLabels
 * 
 * Removes all labels from a thread object, of an array of thread objects
 * 
 *@param {thread} thread - gmailApp thread object for label removal
 *@returns {thread} thread - gmailApp thread object with labels removed
 */

function clearLabels(thread){
  let labels = thread.getLabels();
    for (let label of labels) {
      thread.removeLabel(label);
    }
  thread.refresh();
  return thread;
}


/**
 * setLabel
 * 
 * Clear all existing labels for a thread, and set a new one
 * 
 * @param {object} thread - gmailApp Thread object to change labels for
 * @param {string} destLabel - New label to set for this thread
 * @returns void
 */

function setLabel(thread, destLabel) {

  var newLabel = GmailApp.getUserLabelByName(destLabel);
  if (newLabel !== null){
    thread.addLabel(newLabel).refresh();
  } else {
    Logger.log("ERROR: Could not add label: '" + destLabel + "', does not exist.");
  }
  return thread;
}


/**
 * getFileObj
 * 
 * open any file from g-Drive and return appropriate object
 * type, depending on mime type.
 * NOTE: This function still needs a lot of work
 * 
 * @param {string} fileName - Name of file to be opened
 * @param {string} getAsType - MimeType to return file as
 * @returns {object} Google object representong file or NULL if not found
 */

function getFileObj(fileName, getAsType = null){

  var oFile = null;
  var FileIterator = DriveApp.getFilesByName(fileName);

  while (FileIterator.hasNext())
  {
    var file = FileIterator.next();
    var tmpFilename = file.getName();

    if (file.getName() == fileName)
    {
      var id = file.getId();
      var mimeType = file.getMimeType();
      if (getAsType !== null) {
        //-- Get file as particular mime type
        //TODO Add error handling when 'getAs' fails
        oFile = file.getAs(getAsType);
      } 
      else {
        switch (mimeType)
          {
            //TODO Add support for all gdrive mime types
            //TODO Replace hardcoded mimetypes in case statement with Google Scripts 'mime' enum
            case "application/vnd.google-apps.spreadsheet":
              oFile = SpreadsheetApp.openById(id);
              break;
            case "application/vnd.google-apps.document":
              oFile = DocumentApp.openById(id);
              break;
            case "application/pdf":
              break;
            default:
              oFile = null;
              break;
          }
      }
      return oFile;
    }
  }
}


/**
 * replyToMsg
 * 
 * Create a gmailApp Draft object that will be used to reply
 * to the email messate
 * 
 * @param {object} oMsg -  The gmail message object to be replied to
 * @param {string} strReplyText - Text to be sent as the body of reply
 * @param {object} oAttachment - Attachment to be added to draft reply object
 * @returns {int} MessageId of the reply sent
 */
function getReply(oMsg, strReplyTxt, oAttachment){
  var sender = oMsg.getFrom();
  var firstName = sender.split(" ")[0];
  var subject = "RE: " + oMsg.getSubject();
  // var subject = thread.getFirstMessageSubject();
  var strBody = "Dear " + firstName + ":,\n\n" + strReplyTxt;

  var oDraft = oMsg.createDraftReply(strBody,{
    from: "Don Briggs <DonBriggsWork@gmail.com>",
    subject: subject,
    attachments: [oAttachment]
  });
  return oDraft;
}

