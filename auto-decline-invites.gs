// Sync your calendar and decline events you are invited for while you are out of office.

var invited = "INVITED";
var accepted = "YES";
var acceptedStatus = CalendarApp.GuestStatus.YES;
var rejectedStatus = CalendarApp.GuestStatus.NO;
var invitedStatus = CalendarApp.GuestStatus.INVITED;
var maybeStatus = CalendarApp.GuestStatus.MAYBE;
var dayInMillis = 24 * 60 * 60 * 1000;

var myName = 'BLAH'; // <<<<<<<< Put your name here
var myEmail = 'YOUR@EMAIL.HE.RE'; // <<<<<<<< Put your email address here
var calendar = CalendarApp.getCalendarById(myEmail);

function processInvites() {
  var start = new Date();
  var end = new Date(start.getTime() + 14 * dayInMillis);
  var events = calendar.getEvents(start, end, {statusFilters: [acceptedStatus, maybeStatus]});
  processEventList(events);
  var invites = calendar.getEvents(start, end, invited);
  processEventList(invites);
}

function processEventList(eventList) {
  for (var i = 0; i < eventList.length; i++) {
    var event = eventList[i];
    var title = event.getTitle();
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();
    var creators = event.getCreators();
    var status = event.getMyStatus();
    var isAllDay = event.isAllDayEvent();
    Logger.log("Processing: " + title + " from " + creators + " for " + startTime + " - " + endTime);
    if (isAllDay) {
      Logger.log("Not touching all day events.");
      continue;
    }
    var outOfOffices = timesheetCalendar.getEvents(startTime, endTime, {search: 'Out of Office'});
    if (outOfOffices.length > 0) {
      Logger.log("Found an out of office conflict for: " + title);
      if (status == invitedStatus || status == maybeStatus) {
        if (event.getGuestList().length > 20 || !event.guestsCanSeeGuests()) {
          Logger.log("Not likely to expect a response from me.");
          MailApp.sendEmail(myEmail, "Re: Invitation: " + title, "", {htmlBody: "Rejected an event from " + creators + " because I was out of office."});
        } else {
          var body =
              "Hi,<br><br>"+
                "You tried to book my time between <b>" + startTime + "</b> and <b>" + endTime + "</b>. " +
                  "Unfortunately I will be out of office at that time.<br><br>Regards,<br>Norman<br><br>" +
                    "P.S. This email is automatically generated, please let me know if you have received it by error.";
          MailApp.sendEmail(creators, "Re: Invitation: " + title, "", {htmlBody: body});
        }
        event.setMyStatus(rejectedStatus);
      } else if (status == acceptedStatus) {
        Logger.log("Rejecting " + title + " silently");
        event.setMyStatus(rejectedStatus);
        MailApp.sendEmail(myEmail, "Re: Invitation: " + title, "", {htmlBody: "Rejected an already accepted event from " + creators + " because I was out of office."});
      }
    }
  }
}
