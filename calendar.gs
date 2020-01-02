function addEvent() {
  var SACal = CalendarApp.getCalendarsByName('SA Test')[0]; // TODO - Rename SA Test string
  var event = SACal.createEvent('Apollo 11 Landing',
    new Date('January 2, 2020 20:00:00 UTC'),
    new Date('January 2, 2020 21:00:00 UTC'),
    {location: 'The Moon'});
Logger.log('Event ID: ' + event.getId());
};