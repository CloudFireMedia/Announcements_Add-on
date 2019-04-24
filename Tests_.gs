function test_misc() {
  while(true)
  var a = 1
}

function test_compareStrings() {
  var a = 'Title_03_17'
  var b = 'Title_03_17'
  var c = compareStrings(a,b)
  return
}

function test_Comments_getLastTimeScriptRun() {
  Comments_.getLastTimeScriptRun('1muy_pMDEuW7ZMbBOY2JR9_g_PXN8P8yV_w6zR0SvSKw')
  return
}

function test_uniq() {
  var a = [1,2,1,3,2,4]
  var b = arrayUnique(a)
  return
}

function test_modifyDatesInBody() {
  modifyDatesInBody_()
  return
}

function test_formatGDoc() {
  formatGDoc_()
  return
}

function test_short() {
  var text = '[ title | FN1 LN1 ] 01.01;'
  var re = /^\[[^\]\|]+\|[^\]]+\](\D+\d{1,2}\.\d{1,2}\W+)/gi;
  var match = new RegExp(re).exec(text);
  return
}

function test_reorderParagraphs() {
  reorderParagraphs_()
}

function test_copySlides() {
  copySlides_()
}

function test_inviteStaffSponsorsToComment() {
  inviteStaffSponsorsToComment_()
}

function test_rotateContent() {
  rotateContent_()
}

function test_getWeekTitles() {
  getWeekTitles()
}

function test_getChildren() {
  var body = DocumentApp.openById('11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58').getBody()
  var numberOfChildren = body.getNumChildren()
  for (var childIndex = 0; childIndex < numberOfChildren; childIndex++) {
    Logger.log(body.getChild(childIndex).getType())
  }
}

// Comments
// --------

function test_createComment() {
  createComment('Test1')
  test_listComments()
}

function test_removeComments() {
  var DOC_ID = '11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58'
//  var DOC_ID = '1L8myiog-o6puURPuWu7R6UOntPRp4-_xfHcdUTo3Yzc'
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
  var fileId = '11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58'
  var commentId = 'AAAACkLLiH0'
  var resource = {status: 'resolved'}
  var result = Drive.Comments.update(resource, fileId, commentId)
  return
}

function test_listComments() {

  var DOC_ID = '11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58'
//  var DOC_ID = '1L8myiog-o6puURPuWu7R6UOntPRp4-_xfHcdUTo3Yzc'

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