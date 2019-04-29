// Comments
// --------

var DOC_ID = '1muy_pMDEuW7ZMbBOY2JR9_g_PXN8P8yV_w6zR0SvSKw'

function test_createComment() {
  createComment('Test1')
  test_listComments()
}

function test_removeComments() {
  var comments = Drive.Comments.list(DOC_ID); 
  for (var i = 0; i < comments.items.length; i++) {   
    var nextComment = comments.items[i]  
    Logger.log('before - status: ' + nextComment.status + ', deleted: ' + nextComment.deleted + ', content: ' + nextComment.content); 
    Logger.log('Removing ' + nextComment.content)    
    Drive.Comments.remove(DOC_ID, nextComment.commentId)
  } 
  
  test_listComments()
} // test_removeComments()

function test_updateComment() {
  var commentId = 'AAAACkLLiH0'
  var resource = {status: 'resolved'}
  var result = Drive.Comments.update(resource, DOC_ID, commentId)
  return
}

function test_listComments() {

  var comments = Drive.Comments.list(DOC_ID); 
  
  if (comments.items[0] === undefined) {
    Logger.log('No comments')
    return
  }
  
  var numberOfComments = comments.items.length
  Logger.log('Number of comments: ' + numberOfComments)
  
  if (comments.items && numberOfComments > 0) { 
    for (var i = 0; i < numberOfComments; i++) { 
      var comment = comments.items[i]; 
      var modifiedDateString = comment.modifiedDate
      var modifiedDate = new Date(modifiedDateString)      
      Logger.log('content: ' + comment.content + ', status: ' + comment.status + ', deleted: ' + comment.deleted + ', modified: ' + modifiedDateString); 
      var lastWeek = new Date(new Date().getTime() - (7 * 24 * 60 * 60 * 1000))
      if (modifiedDate > lastWeek) {
//       Logger.log('last week')
      } else if (modifiedDate < lastWeek) {
//        Logger.log('Longer than a week')
      }
//      Logger.log('modified: ' + modifiedDateString)
    } 
  } else { 
    Logger.log('No comment found.'); 
  } 
}