/**
 * On opening of the document, add an item "Copy" which calls func go, onClick
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Copy Comments', 'go')
      .addToUi();
}

/**
 * On installation, call onOpen
 */
function onInstall(e) {
  onOpen(e);
}


/**
 * ------------------------------------------------------------------------------------- *
 */

/**
 * Runs the code
 * -Resets the URL property
 * -Calls function to create a copy of the document
 * -Calls function to get sets of 100 comments until all comments have been retrieved
 * -Calls function to append these comments and their replies
 * -Sets the URL property
 * -Calls the URL Dialog pop up
 */
function go(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteAllProperties();
  documentProperties.deleteProperty("URL");
  
  var id = getNewDocId();
  var oldId = DocumentApp.getActiveDocument().getId();
  var pageToken = "";
  
  var props = {};
  props["pageToken"] = "";
  props["id"] = id;
  props["oldId"] = oldId;
  
  documentProperties.setProperties(props);
  
  pickUpFromLeftOff();
}

/**
 * Returns comments in sets of 100
 * /Takes in the documentID and a reference to which set of 100 comments to retrieve
 * -Set optional arguments of the list of comments
 * -If there is a pageToken(reference to the next set of 100 comments), then set that argument
 */
function getComments(docId, prevToken) {
  var optionalArgs = {};
  optionalArgs["maxResults"] = 50;
  optionalArgs["includeDeleted"] = false;
  if(prevToken != ""){
    optionalArgs["pageToken"] = prevToken;
  }
  var comments = Drive.Comments.list(docId, optionalArgs);
  return comments;
}

function no(){
 var docId = DocumentApp.getActiveDocument().getId();
var docFile = DriveApp.getFileById(docId);
var newFile = docFile.makeCopy();
var newId = newFile.getId();
var commentRes = Drive.Comments.list(docId).items;
for( var each in commentRes){
   Drive.Comments.insert(commentRes[each], newId);
} 
}

/**
 * Makes a copy of the document and returns the ID of the copied document
 */
function getNewDocId(){ 
  var docApp = DocumentApp.getActiveDocument();
  var docId = docApp.getId();
  var docFile = DriveApp.getFileById(docId);
  var newFile = docFile.makeCopy();
  var newId = newFile.getId();
  return newId;
}

/**
 * Appends the comments
 * Then appends the replies to each comment
 * Comments made by other authors/people will be created as the user using this add-on
 *
 * Slicing the replyResource is necessary because objects are mutable
 */
function appendCommentsAndReplies(id, comments){
  var fileId = DriveApp.getFileById(id).getId();
  for(var commentRes in comments){
    var commentId = comments[commentRes].commentId;
    var replySave = comments[commentRes].replies.slice();
    comments[commentRes].replies = [];
    var newComment = Drive.Comments.insert(comments[commentRes], id);
    if(comments[commentRes].author["isAuthenticatedUser"] == false){
      var origContent = comments[commentRes].content;
      var authorName = comments[commentRes].author["displayName"];
      //var newContent = "<html>help</html>";
      var newContent = "\"" + authorName + "\"" + ": \n---------------------\n" + origContent;
      Drive.Comments.patch({'content':newContent}, id, newComment.commentId);
    }
    comments[commentRes].replies = replySave;
    if(replySave.length != 0){
      for(var reply in replySave){
        var newReply = Drive.Replies.insert(replySave[reply], fileId, newComment.commentId);
        if(replySave[reply].author["isAuthenticatedUser"] == false){
          var origContent = replySave[reply].content;
          var authorName = replySave[reply].author["displayName"];
          var newContent = "\"" + authorName + "\"" + ": \n---------------------\n" + origContent;
          Drive.Replies.patch({'content':newContent}, id, newComment.commentId, newReply.replyId)
        }
      }
    }
  }
}

var SECOND = 1000;
var MINUTE = 60*SECOND;
var MAX_RUNNING_TIME = 4*MINUTE;
var TIME_TO_WAIT = 1.2*MINUTE;

function pickUpFromLeftOff(){
  var documentProperties = PropertiesService.getDocumentProperties();
  var props = documentProperties.getProperties();
  var pageToken = props["pageToken"];
  var id = props["id"];
  var oldId = props["oldId"];
  var startTime = (new Date()).getTime();
  var comments = [];
  
  if(pageToken == "undefined"){
    appendCommentsAndReplies(id, comments);
    documentProperties.setProperty("URL", DocumentApp.openById(id).getUrl());
    urlDialog();
    return;
  }
  
  do{
    var commentRes = getComments(oldId, pageToken);
    comments = comments.concat(commentRes.items);
    pageToken = commentRes.nextPageToken;
    var currTime = (new Date()).getTime();
    if(currTime - startTime > MAX_RUNNING_TIME){
      var properties = {};
      if(pageToken == undefined){
        pageToken = "undefined";
      }
      properties["pageToken"] = pageToken;
      properties["newDocId"] = id;
      documentProperties.setProperties(properties);
      appendCommentsAndReplies(id, comments);
      var endTime = (new Date()).getTime();
      ScriptApp.newTrigger("pickUpFromLeftOff")
               .timeBased()
               .at(new Date(endTime+TIME_TO_WAIT))
               .create();
      return;
    }
  }while(pageToken != undefined);
  appendCommentsAndReplies(id, comments);
  documentProperties.setProperty("URL", DocumentApp.openById(id).getUrl());
  urlDialog();
}

/**
 * Pop ups the URL Dialog in the Google Document
 */
function urlDialog(){
  var html = HtmlService.createHtmlOutputFromFile('URLDialog')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(200)
      .setHeight(100);
  DocumentApp.getUi()
      .showModalDialog(html, "Copied Document");
}

/**
 * Attempts to retrieve URL
 * Returns non-URL if property does not exist
 * Returns URL if URL property exists/has been set
 */
function retrieveURL(){
  var documentProperties = PropertiesService.getDocumentProperties();
  var keys = documentProperties.getKeys();
  var isURL = false;
  var URL = null;
  for(var i = 0 ; i <keys.length ; i++){
    if(keys[i] == "URL"){
      URL = documentProperties.getProperty("URL");
      break
    }
  }
  return URL;
}





function goSave(){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty("URL");
  var id = getNewDocId();
  var oldId = DocumentApp.getActiveDocument().getId();
  var comments = [];
  var pageToken = "";
  var startTime = (new Date()).getTime();
  
  do{
    var commentRes = getComments(oldId, pageToken);
    comments = comments.concat(commentRes.items);
    pageToken = commentRes.nextPageToken;
    var currTime = (new Date()).getTime();
    if(currTime - startTime > MAX_RUNNING_TIME){
      var properties = {};
      properties["pageToken"] = pageToken;
      properties["newDocId"] = id;
      documentProperties.setProperties(properties);
      appendCommentsAndReplies(id, comments);
      var endTime = (new Date()).getTime();
      ScriptApp.newTrigger("pickOffFromLeftOff")
               .timeBased()
               .at(new Date(endTime+TIME_TO_WAIT))
               .create();
      return;
    }
  }while(pageToken != undefined);
  appendCommentsAndReplies(id, comments);
  documentProperties.setProperty("URL", DocumentApp.openById(id).getUrl());
  urlDialog();
}

