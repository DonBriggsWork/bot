// @TS -nocheck
/**
 * ResumeBot
 * 
 * @author DonBriggs <DonBriggsWork@gmail.com>
 * @version 0.11
 * 
 * Search through inbox for messages marked with PROCESS_LABEL label. Reply to
 * them with a standard template, attach resume, mark them as read, and move 
 * them to the "DEST_LABEL" label. Program parameters are set in the
 * "settings.gs" file.
 * 
 * 2024-03-23: Simplified code by removing comments, logger output and debug code.
 */


/**
 * main
 * 
 * returns void
 */
  
  function main() {

    initialize();
    const threads = getThreadsByLabel(getProp('PROCESS_LABEL'));
    if(!threads.length) {
      Logger.log("No threads to process");
      return;
    }
    runThreads(threads);
  }  


/**
 * runThreads
 * 
 * Selects threads from the inbox, based on search criteria, iterates
 * through the result set and automatically sends a reply.
 */

function runThreads(threads) {

  if(threads === "" || threads === null || threads === undefined) {
    throw new Error("runThreads: No thread collection passed to function");
  }

  const replyText = getFileObj(getProp('REPLY_FILE')).getBody().getText();
  const oAttachment = getFileObj(getProp('ATTACH_FILE'), MimeType.PDF);
  const oLabelRemove = GmailApp.getUserLabelByName(getProp('PROCESS_LABEL'));
  const oLabelAdd    = GmailApp.getUserLabelByName(getProp('DEST_LABEL'));

  for (let thread of threads) {
    var msgId = thread.getMessages()[0].getId();
    var oMsg = GmailApp.getMessageById(msgId);
    thread.markRead();
    thread.moveToArchive();
    thread.removeLabel(oLabelRemove);
    thread.addLabel(oLabelAdd);

    // Construct reply and send it.
    var oReply = getReply(oMsg, replyText, oAttachment);
    var replyId = oReply.send();
  }  
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
  if (strLabel === "" || strLabel === null || strLabel === undefined) {   
    throw new Error('getThreadsByLabel: Null label passed to function');
  }

  const label = GmailApp.getUserLabelByName(strLabel);
  if (label == '') {
    throw new Error('getThreadsByLabel: Label not found: ' + strLabel);
  }
  return label.getThreads();
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

  if (fileName === "" || fileName === null || fileName === undefined) {
    throw new Error('getFileObject: Null fileName passed to function');
  }

  var oFile = null;
  var FileIterator = DriveApp.getFilesByName(fileName);

  while (FileIterator.hasNext())
  {
    var file = FileIterator.next();
    var tmpFilename = file.getName();

    if (file.getName() == fileName)
    {
      const id = file.getId();
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