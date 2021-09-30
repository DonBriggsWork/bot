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

  //-- Job Configuration Blocks
  // const PROCESS_LABEL = "AUTO_TRACK";             //-- Look for threads with this label to process
  // const DEST_LABEL    = "Followup";               //-- Label to move threads to after processed
  // const ATTACH_FILE   = "Don!s Resume 2021c.pdf";
  // const REPLY_FILE    = "Submission Followup";

  const PROCESS_LABEL = "AUTO_RESUME";            //-- Look for threads with this label to process
  const DEST_LABEL    = "Jobs";                   //-- Label to move threads to after processed
  const ATTACH_FILE   = "Don!s Resume 2021c.pdf";
  const REPLY_FILE    = "Submission Followup";

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
     oReplyFile = null;  //-- Unset variable
   }

  //-- Setup file to attach to reply
  oAttachment = getFileObj(ATTACH_FILE, MimeType.PDF);
  Logger.log("Attachment: " + ATTACH_FILE);


  //--- Loop through list, process threads
  for (var i in threads){
    Logger.log("  Thread: " + i);
 
    setLabel(threads[i], DEST_LABEL);
    msgId = threads[i].getMessages()[0].getId();
    oMsg = GmailApp.getMessageById(msgId);  
    replyId = replyToMsg(oMsg, replyText, oAttachment);
    threads[i].moveToArchive();
  }
}

  //----------------------
  //---   FUNCTIONS   ----
  //----------------------


/**
 * getThreads
 * 
 * Retuns an array of Thread objects with a particular label
 * @param {string} strCriteria Label gmail search criteria string
 * @return array of Thread objects that match search criteria
*/

function getThreads(strLabel)
{
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
 * @param {threads[]} threads - array of gmailApp objects for label removal
* @returns {threads[] threads -
 */
function clearLabels(threads) {

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

  var labelName;
  var newLabel = GmailApp.getUserLabelByName(destLabel);
  var numLabels;
  var labels = new Array();

  labels = thread.getLabels();
  numLabels = labels.length;

  if(numLabels > 0) {
    for (var i = 0; i < numLabels; i++) {
      labelName = labels[i].getName();
      thread.removeLabel(labels[i]);
    }
  }
  thread.addLabel(newLabel);
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
function replyToMsg(oMsg, strReplyTxt, oAttachment) {

  var sender = oMsg.getFrom()
  // Split sender name on spaces to get first name
  var firstName = sender.split(" ")[0];
  var subject = "RE: " + oMsg.getSubject();
  var strBody = "Dear " + firstName + ",\n\n" + strReplyTxt;

  var oDraft = oMsg.createDraftReply(strBody,{
    from: "Don Briggs <DonBriggsWork@gmail.com>",
    subject: subject,
    attachments: [oAttachment]
  });

  //oDraft.send();
  return oDraft.getId();
}


function processAttachments() {
  var txtCriteria = "subject:Attachment Test";
  var threads = getThreadList(txtCriteria);
  var msg = null;
  if(threads.length > 0) {
    Logger.log("Found: "+threads.length +" threads with attachments");
    for (var i in threads) {
      var messages = threads[i].getMessages();
      for (var j in messages) {
       Logger.log(messages[j].getSubject());
       Logger.log(messages[j].getAttachments.length);
       for (var y in attachments) {
       var name = attachments[y].getName();
       Logger.log("Name: " + name);  
       }  
      }
    }
  }
}

/**
 * getThreads
 * 
 * Retuns an array of Thread objects with a particular label
 * @param {string} strCriteria Label gmail search criteria string
 * @return array of Thread objects that match search criteria
*/

function getThreads(strLabel)
{
  Logger.log("Search: "  + strLabel);
  var threads = GmailApp.search('label: ' + strLabel);
  if (threads.length > 0) {
    return threads;
  } else {
    return new Array();
  };
}


function getLog(strFilename) {
  Logger.log("Opening log file: "+strFilename);
  var files = DriveApp.getFilesByName(strFilename); // Get all files with name.
   while (files.hasNext()) {
      var file = files.next();
      var fileID = file.getId();
      logFile = DriveApp.getFileById(fileID);
      return logFile;
    }
  }

function writeLog(objFile, txtContent) {
  objFile.append(txtContent);
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

  var labelName;
  var newLabel = GmailApp.getUserLabelByName(destLabel);
  var numLabels;
  var labels = new Array();

  labels = thread.getLabels();
  numLabels = labels.length;

  if(numLabels > 0) {
    for (var i = 0; i < numLabels; i++) {
      labelName = labels[i].getName();
      thread.removeLabel(labels[i]);
    }
  }
  thread.addLabel(newLabel);
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
function replyToMsg(oMsg, strReplyTxt, oAttachment) {
  var sender = oMsg.getFrom()

  var firstName = sender.split(" ")[0];
  var subject = "RE: " + oMsg.getSubject();
  var strBody = "Dear " + firstName + ",\n\n" + strReplyTxt;

  var oDraft = oMsg.createDraftReply(strBody,{
    from: "Don Briggs <DonBriggsWork@gmail.com>",
    subject: subject,
    attachments: [oAttachment]
  });

  oDraft.send();
  return oDraft.getId();
}


function processAttachments() {
  var txtCriteria = "subject:Attachment Test";
  var threads = getThreadList(txtCriteria);
  var msg = null;
  if(threads.length > 0) {
    Logger.log("Found: "+threads.length +" threads with attachments");
    for (var i in threads) {
      var messages = threads[i].getMessages();
      for (var j in messages) {
       Logger.log(messages[j].getSubject());
       Logger.log(messages[j].getAttachments.length);
       for (var y in attachments) {
       var name = attachments[y].getName();
       Logger.log("Name: " + name);  
       }  
      }
    }
  }
}