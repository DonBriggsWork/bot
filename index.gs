/**
 * doGet
 * This function is called to process GET request when web URL is called
 * <return> {object} HtmlOUtputObject
 */

function doGet(e) {
  var params = JSON.stringify(e);
  Logger.log(params);
  var out =  HtmlService
            .createTemplateFromFile('main-layout')
            .evaluate();
  return out;
}

function include(fileName){
  // Logger.log("Running Include: " + fileName);
  var partial = HtmlService
            .createTemplateFromFile(fileName)
            .evaluate()
            .getContent();
  // partial = "\n" + partial + "\n";
  return "\n" + partial + "\n";
}

/**
 * threads
 * 
 * Switched GUI to 'threads' mode
 * 
 * <param> 
 */
function threads() {

}