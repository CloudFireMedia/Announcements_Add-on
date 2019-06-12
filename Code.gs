var SCRIPT_NAME = "Announcements_Add-on"
var SCRIPT_VERSION = "v1.4.dev_ajr"

function onInstall(event) {
  onOpen(event)
}

function onOpen(event) {

  console.log(SCRIPT_VERSION)
  console.log(event)

  var ui = DocumentApp.getUi()
  
  if (event !== undefined && event.authMode !== ScriptApp.AuthMode.FULL) {
    ui.createAddonMenu().addItem('Start - Display full menu', 'getAuth').addToUi()
    return
  }
  
  var boundDoc = DocumentApp.getActiveDocument()
  
  if (boundDoc === null) {
    throw new Error('No bound doc')
  }
  
  var boundDocId = boundDoc.getId()
  var menu = ui.createMenu('CloudFire')

  if      (Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID') === boundDocId) { master()    }
  else if (Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID') === boundDocId) { zeroWeeks() }
  else if (Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID') === boundDocId)  { oneWeek()   }
  else if (Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID') === boundDocId) { twoWeeks()  }
  else if (Config.get('ANNOUNCEMENTS_ARCHIVE_ID') === boundDocId)       { archive()   }
  else { throw new Error('This GDoc is not supported by the add-on') }
  
  menu.addToUi()
  
  return
  
  // Private Functions
  // -----------------

  function master() {
  
    menu
      .addItem('Add Week', 'master_insertWeeks')
      .addItem('Add Month', 'master_insertMonth')
      .addSeparator()
      .addItem('Format Document', 'format_master')
      .addSubMenu(
        ui.createMenu('Instructions...')
        .addItem('Document Overview', 'showInstructions_Document')
        .addItem('Recurring Content', 'showInstructions_RecurringContent')
      )      
        
  } // onOpen.master()
  
  function zeroWeeks() {
  
    menu
      .addItem('Format', 'runAllFormattingFunctions_upcomingWeek')
      .addItem('Move Service Slides', 'copySlides')
      .addItem('Email Staff', 'emailStaff')
      
  } // onOpen.zeroWeeks()
  
  function oneWeek() {
  
    menu
      .addItem('Format All', 'runAllFormattingFunctions_oneWeekOut')
      .addSeparator()
      .addItem('Format Font', 'formatFont_oneWeek')
      .addItem('Format Paragraphs', 'format_removeEmptyParagraphs')
      
  } // onOpen.oneWeek()
  
  function twoWeeks() {
  
    var currentSubMenu = ui.createMenu('... for the current document')
      .addItem('Sort events chronologically', 'reorderParagraphs')
      .addItem('Remove "short start dates" from document', 'removeShortStartDates')
      .addItem('Format document styles', 'formatGDoc')
      //          .addItem('Update event descriptions', 'matchEvents')
      .addItem('Phrase "long start dates" in natural language', 'modifyDatesInBody')
      .addItem('Display tally of historical instances of each Live Announcement', 'countInstancesofLiveAnnouncement')
      .addItem('Remove tally of historical instances of each Live Announcement', 'cleanInstancesofLiveAnnouncement')

    var otherSubMenu = ui.createMenu('... for other CloudFire components')
      .addItem('Copy This Sunday\'s Service Slides to This Sunday\'s \'Live  Slides\' folder', 'copySlides')

    var manualSubMenu = ui.createMenu('Manually execute individual scripts...')
    manualSubMenu.addSubMenu(currentSubMenu)
    manualSubMenu.addSubMenu(otherSubMenu)

    menu
      .addItem('Rotate content', 'rotateContent')
      .addItem('Display tally of historical instances of each Live Announcement', 'countInstancesofLiveAnnouncement')
      .addItem('Invite Event Sponsors to comment', 'inviteStaffSponsorsToComment')      
      .addSeparator()      
      .addSubMenu(manualSubMenu)
                      
  } // onOpen.twoWeeks()
    
  function archive() {
  
    menu
      .addItem('Format', 'formatFont_archive')
      
  } // onOpen.archive()

} // onOpen()

function getAuth() {onOpen()}

// Master Sunday
// -------------

function master_insertWeeks()                {Announcements.master_insertWeeks()}
function master_insertMonth()                {Announcements.master_insertMonth()}
function format_master()                     {Announcements.format_master()}
function showInstructions_Document()         {Announcements.showInstructions_Document()}
function showInstructions_RecurringContent() {Announcements.showInstructions_RecurringContent()}

// 0 Weeks
// -------

// Menu
function runAllFormattingFunctions_upcomingWeek() {Announcements.runAllFormattingFunctions_upcomingWeek()}
function copySlides()                             {Announcements.copySlides()}
function emailStaff()                             {Announcements.emailStaff()}

// Client-side
function emailStaff_submit(executiveAssitant) {Announcements.emailStaff_submit(executiveAssitant)}

// 1 Week
// ------

function runAllFormattingFunctions_oneWeekOut() {Announcements.runAllFormattingFunctions_oneWeekOut()}
function formatFont_oneWeek()                   {Announcements.formatFont_oneWeek()}
function format_removeEmptyParagraphs()         {Announcements.format_removeEmptyParagraphs()}

// 2 Weeks
// -------

function inviteStaffSponsorsToComment()     {Announcements.inviteStaffSponsorsToComment()}
function rotateContent()                    {Announcements.rotateContent()}
// function copySlides()                       {Announcements.copySlides()} // In 0 Weeks
function reorderParagraphs()                {Announcements.reorderParagraphs()}
function removeShortStartDates()            {Announcements.removeShortStartDates()}
function formatGDoc()                       {Announcements.formatGDoc()}
function modifyDatesInBody()                {Announcements.modifyDatesInBody()}
function countInstancesofLiveAnnouncement() {Announcements.countInstancesofLiveAnnouncement()}
function cleanInstancesofLiveAnnouncement() {Announcements.cleanInstancesofLiveAnnouncement()}

// Archive
// -------

function formatFont_archive() {Announcements.formatFont_archive()}