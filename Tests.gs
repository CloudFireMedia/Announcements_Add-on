function test() {
  var d = new Date(0)
  return
}

function test_listComments() {

  var comments = Drive.Comments.list(DocumentApp.getActiveDocument().getId());

  if (comments.items && comments.items.length > 0) {
    for (var i = 0; i < comments.items.length; i++) {
      var comment = comments.items[i];  
      Logger.log('Comment : %s, modified: %s', comment.content, comment.modifiedDate);
    }
  } else {
    Logger.log('No comment found.');
  }
}

function test_addHiddenComment() {
  var fileId = DocumentApp.getActiveDocument().getId()
  var resource = {content: 'Script last run: ' + new Date()}
  Drive.Comments.insert(resource, fileId)
}

function test_getLastRunTime() {
  var comments = Drive.Comments.list(DocumentApp.getActiveDocument().getId());
  if (comments.items && comments.items.length > 0) {
    for (var i = 0; i < comments.items.length; i++) {
      var content = comments.items[i].content
      if (content.indexOf('Script last run:') !== -1) {  
        var dateString = content.slice(16)
        Logger.log('Datetime : %s', dateString);        
        var datetime = new Date(dateString)
        Logger.log('Datetime : %s', datetime);
        if (datetime < new Date()) {
          Logger.log('old comment')
        }
      }
    }
  } else {
    Logger.log('No comments found.');
  }
}

function test_storeScriptRunTime() {
  var fileId = DocumentApp.getActiveDocument().getId()
  var comments = Drive.Comments.list(fileId);
  if (comments.items && comments.items.length > 0) {
    for (var i = 0; i < comments.items.length; i++) {
      var comment = comments.items[i]
      var content = comment.content
      if (content.indexOf('Script last run:') !== -1) {
        var resource = {content: 'Script last run: ' + new Date()}
        Drive.Comments.patch(resource, fileId, comment.commentId)
        test_listComments()
      }
    }
  } else {
    Logger.log('No comments found.');
  }  
}

