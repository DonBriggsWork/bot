// @TS -nocheck
/**
 * ResumeBot
 * 
 * @author DonBriggs <donBriggsWork@gmail.com>
 * @version 0.1
 * 
 * Search through inbox for messages marked with "PROCESS_LABEL" label. Reply to
 * them with a standard template, attach resume, mark them as read, 
 * and move them to the "DEST_LABEL" label. Program parameters are set in the
 * "settings.gs" file.
 */


  /**
   * main
   * 
   * Main function. Run this function to run the program.
   */
  
  function main() {

    SetVars();
    threads = getThreadsByLabel(getProp('PROCESS_LABEL'));
    if(!threads.length) {
      Logger.log("No threads  to process, Exiting");
      return;
    }
    runThreads(threads);
  }  


/**
 * runThreads
 * 
 * Selects threads from the inbox, based on their label, automatically
 * sends a reply and begins tracking the thread
 */

function runThreads(threads) {

  var replyText = getFileObj(getProp('REPLY_FILE')).getBody().getText();
  var oAttachment = getFileObj(getProp('ATTACH_FILE'), MimeType.PDF);
  oLabelRemove = GmailApp.getUserLabelByName(getProp('PROCESS_LABEL'));
  oLabelAdd    = GmailApp.getUserLabelByName(getProp('DEST_LABEL'));

  for (let thread of threads) {
    var msgId = thread.getMessages()[0].getId();
    var oMsg = GmailApp.getMessageById(msgId);
    thread.markRead
    thread.moveToArchive();
    thread.removeLabel(oLabelRemove);
    thread.addLabel(oLabelAdd);

    var oReply = getReply(oMsg, replyText, oAttachment);
    replyId = oReply.send();

    var result = {
      'msgId':   msgId,
      'subject': thread.getMessages()[0].getSubject(),
      'subject': thread.getMessages()[0].markRead(),
      'sender':  thread.getMessages()[0].getFrom(),
      'recDate': thread.getMessages()[0].getDate(),
      'replyDate': Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy"),
      'replyId': replyId,
    }
  }  

}


/**
 * getThreads
 * 
 * Retuns an array of Thread objects by search string
 * 
 * @param {string} strCriteria Label gmail search criteria string
 * @return array of Thread objects that match search criteria
*/

function getThreads(strSearch){

  threads = GmailApp.search('label: ' + strSearch);
  return threads;
}

/**
 * getThreadsByLabel
 * 
 * Returns all threads with a given label
 * 
 * @param {string} strLabel Label to search threads
 * @return array of Thread objects that match search criteria
 */
function getThreadsByLabel(strLabel) {

  if (strLabel == '') {
    throw new Error('getThreadsByLabel: Null label passed to function');
  }
  var label = GmailApp.getUserLabelByName(strLabel);
  if (label == '') {
    throw new Error('getThreadsByLabel: Label not found: ' + strLabel);
  }
  var threads = label.getThreads();
  return threads;
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
  throw new Error("File ojject not found: " + fileName);
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

  var oOptions = {from: getProp('REPLY_FROM'),
                  subject: subject,
                  attachments: oAttachment};

  var oDraft = oMsg.createDraftReply(strBody, oOptions);
  return oDraft;
}






