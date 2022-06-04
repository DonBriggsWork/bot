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
 * 
 * Branch: read_from_sheet
 */

/**
 * runThreads
 * 
 * Selects threads from the inbox, based on their label, automatically
 * sends a reply and begins tracking the thread
 */

function runThreads() {

  Logger.log("Beginning Run");
  SetVars();
  Logger.log("Searching for: " + getProp('PROCESS_LABEL'));

  var threads = getThreads(getProp('PROCESS_LABEL'));
     Logger.log("  - Found threads to process: " + threads.length);
  if (threads.length == 0) {
    Logger.log("No threads to process. Program exiting.");
    return new Array(); 
  }
  else {
    Logger.log("Found threads to process: " + threads.length);
  }

  var oReplyFile  = getFileObj(getProp('REPLY_FILE'));
  var replyText   = oReplyFile.getBody().getText();
  var oAttachment = getFileObj(getProp('ATTACH_FILE'), MimeType.PDF);
  Logger.log("Attachment: " + getProp('ATTACH_FILE'));

  var results = Array();

  for (let thread of threads) {
    var msgId = thread.getMessages()[0].getId();
    var oMsg = GmailApp.getMessageById(msgId);
    thread.moveToArchive();
    // thread = clearLabels(thread);
    thread.removeLabel(getProp('PROCESS_LABEL'));
    setLabel(thread, getProp('DEST_LABEL'));

    var oReply = getReply(oMsg, replyText, oAttachment);

    Logger.log("  - Processing Message ID: " + msgId);
    var result = {
        'msgId':   msgId,
        'subject': thread.getMessages()[0].getSubject(),
        'sender':  thread.getMessages()[0].getFrom(),
        'recDate': thread.getMessages()[0].getDate(),
        'replyDate': "xxx",
        'replyId': oReply.getId(),
    }    

    if (getProp('DEBUG') > 0) {
       Logger.log("  * FIRING");
      oReply.send()
    }
  }  
  Logger.log("Completed");
}


/**
 * getThreads
 * 
 * Retuns an array of Thread objects with a particular label
 * r
 * @param {string} strCriteria Label gmail search criteria string
 * @return array of Thread objects that match search criteria
*/

function getThreads(strLabel){
  var threads = new Array();
  Logger.log(" - GETTING THREADS FROM INBOX: " + strLabel);  
  Logger.log("Search: "  + strLabel);
  
  var threads = GmailApp.search('label: ' + strLabel);
  return threads;
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
  Logger.log("  - Clearing message labels");
  let labels = thread.getLabels();
    for (let label of labels) {
      thread.removeLabel(label);
    }
  thread.refresh();
  return thread;
}


/**
 * setLabels
 * 
 * Clear all existing labels for a thread, and set a new one
 * 
 * @param {object} thread - gmailApp Thread object to change labels for
 * @param {string} destLabel - New label to set for this thread
 * @returns void
 */

function setLabels(thread, destLabels) {

  var newLabel = GmailApp.getUserLabelByName(destLabel);
  if (newLabel !== null){
    thread.addLabel(newLabel).refresh();
  } else {
    throw new Error("ERROR: Could not add label: '" + destLabel + "', does not exist.");
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

  Logger.log("  - Getting File Object");
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
 * getReply
 * 
 * Create a gmailApp Draft object that will be used to reply
 * to the email, and return it.
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
  var strBody = "Dear " + firstName + ":,\n\n" + strReplyTxt;

  var oDraft = oMsg.createDraftReply(strBody,{
    from: "Don Briggs <DonBriggsWork@gmail.com>",
    subject: subject,
    attachments: [oAttachment]
  });
  return oDraft;
}