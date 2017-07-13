// Respond to HipChat notifications, when you are offline and out of office.

var myEmail = 'YOUR@EMAIL.HE.RE'; // <<<<<<<< Put your email address here
var labelNewEmails = 'HipChat Ping'; // <<<<<<<< Where HipChat notifications will be found?
var labelEmailsToReadLater = 'HipChat Ping to Read'; // <<<<<<<< Label for those I will need to read later (Came when I was offline but not out of office).

var now = new Date();
var start = new Date(now.getTime() - 5 * 60 * 1000);
var end = new Date(now.getTime() + 5 * 60 * 1000);
var timesheetCalendar = CalendarApp.getCalendarsById();
var outOfOfficeEvents = timesheetCalendar.getEvents(now, end, {search: 'Out of Office'});
var outOfOffice = outOfOfficeEvents.length > 0;
  

function autoRespondHipChatEmails() {
  var threads = GmailApp.search('label:"' + labelNewEmails + '" label:unread');
  if (outOfOffice) {
    // Send response and mark as read.
    Logger.log('Sending replies to ' + threads.length + ' pings.');
    for (var i = 0; i < threads.length; i++) {
      var replyText = outOfOfficeEvents[0].getDescription();
      threads[i].reply('', {htmlBody: replyText, from: myEmail});
      threads[i].markRead();
    }
  } else {
    // Replace label, so that I can read it later.
    var oldLabel = GmailApp.getUserLabelByName(labelNewEmails);
    var newLabel = GmailApp.getUserLabelByName(labelEmailsToReadLater);
    Logger.log('There is no current out of office event, replacing label for ' + threads.length + ' pings.');
    for (var i = 0; i < threads.length; i++) {
      threads[i].removeLabel(oldLabel);
      threads[i].addLabel(newLabel);
    }
  }
}
