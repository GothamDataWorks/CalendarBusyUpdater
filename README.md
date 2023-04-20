# CalendarBusyUpdater
Takes a list of source calendars, pulls each event from them, then blocks out generic "busy" events on a target calendar based on those source events. Also, any generic-named events in the target that are no longer in the source calendars are removed from the target calendar. Any target calendar events with a non-generic name are KEPT. 

Example use would be making a generic calendar that a service like Calendly has access to. 

Code is in AppleScript. 

Note: the prefs are asked the 1st time this runs, saved in a plist file in ~/Library/Preferences. Currently, the only way to change those settings after the initial run is to edit or delete that plist file. 