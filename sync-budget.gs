// This script syncs data from a Google calendar into a spreadsheet, in order to forecast a budget.

const dayInMillis = 24 * 60 * 60 * 1000;

const reportPeriodInDays = 90;
const maxEventsToProcess = 1000;
const prefixIncome = '[Income] ';
const prefixReminder = '[Reminder] ';
const paymentsCalendar = CalendarApp.getCalendarById('CALENDAR ID HERE');

const spreadsheetDocument = SpreadsheetApp.openByUrl('SPREADSHEET URL HERE');
const eventsTab = spreadsheetDocument.getSheetByName('Events');
const columnNameDate = 'Date';
const columnNameAmount = 'Amount';
const columnNameDescription = 'Description';
const monthStartDescription = 'NEW MONTH'

function processPayments() {
  const start = new Date();
  const end = new Date(start.getTime() + reportPeriodInDays * dayInMillis);
  const events = paymentsCalendar.getEvents(start, end, {max: maxEventsToProcess});
  if (!events || events == null) {
    Logger.log('Couldn\'t read events from calendar.');
    return;
  }
  if (events.length == maxEventsToProcess) {
    Logger.log(`Couldn't read all the events in the date range, only processing ${maxEventsToProcess} events.`);
  }
  clearSpreadsheet();
  processEventList(events);
  formatSpreadsheet();
}

function clearSpreadsheet() {
  // Logger.log(`Clearing spreadsheet: ${spreadsheetDocument.getName()}`);
  eventsTab.clearContents();
  eventsTab.appendRow(['Date', 'Credit', 'Debit', 'Description']);
}

function formatSpreadsheet() {
  // Logger.log(`Formating spreadsheet: ${spreadsheetDocument.getName()}`);
  const lastRow = eventsTab.getMaxRows();
  // Format headers
  eventsTab.getRange('A1:D1').setFontWeight('bold');
  // Format dates and amounts
  eventsTab.getRange(`A2:A${lastRow}`).setNumberFormat('yyyy/MM/dd');
  eventsTab.getRange(`B2:B${lastRow}`).setNumberFormat('$0');
  eventsTab.getRange(`C2:C${lastRow}`).setNumberFormat('$0');
}

const localDateFormatter = new Intl.DateTimeFormat('en-US', {
  timeZone: 'Australia/Sydney',
  weekday: 'long',
  year: 'numeric',
  month: 'numeric',
  day: 'numeric' });

function getLocalDate(date) {
  const partValues = localDateFormatter.formatToParts(date).map(p => p.value);
  return `${partValues[6]}-${partValues[2].padStart(2, '0')}-${partValues[4].padStart(2, '0')}`;
}

function processEvent(event) {
  const title = event.getTitle();
  // Logger.log(`Processing event: ${title}`);
  if (!event.isAllDayEvent()) {
    Logger.log(`Ignoring event that is not all-day: ${title}`);
    return undefined;
  }
  const startTime = event.getStartTime();
  const endTime = event.getEndTime();
  if (endTime.getTime() - startTime.getTime() > dayInMillis) {
    Logger.log(`Ignoring multi-day evnet: ${title}`);
    return undefined;
  }
  const eventDate = getLocalDate(startTime);

  if (event.getMyStatus() !== CalendarApp.GuestStatus.OWNER) {
    Logger.log('Ignoring events without an "owner" status.');
    return undefined;
  }
  if (title.startsWith(prefixReminder)) {
    Logger.log(`Ignoring reminder: ${title}`);
    return undefined;
  }
  const income = (title.startsWith(prefixIncome) === true);
  const titleToParse = income ? title.slice(prefixIncome.length) : title;
  const titleParts = titleToParse.match(/^(\S+)\s+(.+)/);
  if (titleParts.length < 3) {
    Logger.log('Ignoring misformatted title.');
    return undefined;
  }
  const amount = parseInt(titleParts[1].replace(/\,/,"").replace(/\$/,""));
  const description = titleParts[2];
  return {
    date: eventDate,
    credit: income ? amount : undefined,
    debit: income ? undefined : amount,
    description
  };
}

function processEventList(eventList) {
  const validEvents = eventList.map(processEvent).filter(event => (event && event !== null));
  const monthsVisited = {};
  validEvents.forEach(event => {
    const monthStartDate = event.date.slice(0, 7);
    if (!monthsVisited.hasOwnProperty(monthStartDate)) {
      Logger.log(`Adding start of month event for [${monthStartDate}]`);
      eventsTab.appendRow([`${monthStartDate}-01`, undefined, undefined, monthStartDescription]);
      monthsVisited[monthStartDate] = true;
    }
    Logger.log(`Adding event on [${event.date}] - credit: ${event.credit} debit: ${event.debit} from ${event.description}`);
    eventsTab.appendRow([event.date, event.credit, event.debit, event.description]);
  });
}
