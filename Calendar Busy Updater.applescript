-- Calendar Busy Updater

(*
	Take a list of source calendars, pulls each event from them, then blocks out generic "busy" events on a target calendar. Example use would be making a generic calendar that a service like Calendly has access to. 
	Note: the prefs are asked the 1st tie this runs, saved in a plist file in ~/Library/Preferences. To reset, edit or delete that plist file. 
	
HISTORY: 
	2023-04-20 ( danshockley ): First created. 
	
*)

-- internal prefs (maybe move to plist?):
property doRemoveOutdatedBusyEvents : true

-- Do not need to modify these properties:
property prefsDomain : "com.GothamDataWorks.CalendarBusyUpdater"

-- Define globals (that are not persistent properties)
global prefsPath
global calendarSourceNames
global calendarTargetName
global daysBack
global daysFuture
global genericBusyEventTitle


on run
	
	set prefsPath to ((path to preferences folder) as string) & prefsDomain & ".plist"
	
	if not initPrefs() then return false
	
	
	set searchDateRangeStart to (current date) - daysBack * days
	set searchDateRangeEnd to (current date) + daysFuture * days
	
	updateCalendar()
	
	return true
	
end run


on updateCalendar()
	
	set searchDateRangeStart to (current date) - daysBack * days
	set searchDateRangeEnd to (current date) + daysFuture * days
	
	tell application "Calendar"
		
		repeat with oneCalendarName in calendarSourceNames
			set oneCalendarName to contents of oneCalendarName
			
			set eventList to (every event of calendar oneCalendarName whose start date is greater than or equal to searchDateRangeStart and start date is less than or equal to searchDateRangeEnd)
			
			repeat with oneEvent in eventList
				
				set oneStartDate to start date of oneEvent
				set oneEndDate to end date of oneEvent
				tell calendar calendarTargetName
					make new event with properties {summary:genericBusyEventTitle, start date:oneStartDate, end date:oneEndDate}
				end tell
				
			end repeat
		end repeat
		
	end tell
	
	return true
	
end updateCalendar


on initPrefs()
	if testPathExists(prefsPath) then
		readPrefs()
		return true
	else
		askPrefs()
		if result then
			writePrefs()
		end if
		return true
	end if
end initPrefs


on readPrefs()
	
	set calendarSourceNames to plistRead(prefsPath, "calendarSourceNames")
	set calendarTargetName to plistRead(prefsPath, "calendarTargetName")
	set genericBusyEventTitle to plistRead(prefsPath, "genericBusyEventTitle")
	set daysBack to plistRead(prefsPath, "daysBack")
	set daysFuture to plistRead(prefsPath, "daysFuture")
	
	return true
	
end readPrefs


on writePrefs()
	
	plistWrite(prefsPath, "calendarSourceNames", calendarSourceNames)
	plistWrite(prefsPath, "calendarTargetName", calendarTargetName)
	plistWrite(prefsPath, "genericBusyEventTitle", genericBusyEventTitle)
	plistWrite(prefsPath, "daysBack", daysBack)
	plistWrite(prefsPath, "daysFuture", daysFuture)
	
	return true
	
end writePrefs


on askPrefs()
	
	set calendarSourceNames to text returned of (display dialog "Enter your 'source' calendar names (comma-delimited list):" default answer "" buttons {"Cancel", "Next"} default button "Next")
	set calendarSourceNames to replaceSimple({calendarSourceNames, "," & space, ","})
	set calendarSourceNames to parseChars({calendarSourceNames, ","})
	
	
	set calendarTargetName to text returned of (display dialog "Enter your 'target' calendar name (only ONE):" default answer "" buttons {"Cancel", "Next"} default button "Next")
	
	set genericBusyEventTitle to text returned of (display dialog "Enter the generic 'busy' event name:" default answer "BUSY Generic" buttons {"Cancel", "Next"} default button "Next")
	
	set daysBack to (text returned of (display dialog "Enter how many days back into the past this script should scan for events:" default answer 1 buttons {"Cancel", "Next"} default button "Next")) as number
	
	set daysFuture to (text returned of (display dialog "Enter how many days into the future this script should scan for events:" default answer 90 buttons {"Cancel", "Next"} default button "Next")) as number
	
	
	return true
	
end askPrefs


on plistRead(plistPath, plistItem)
	-- version 1.0, Daniel A. Shockley
	
	tell application "System Events"
		set plistFile to property list file plistPath
		return value of property list item plistItem of plistFile
	end tell
end plistRead


on plistWrite(plistPath, plistItemName, plistItemValue)
	-- version 1.1, Daniel A. Shockley
	
	-- 1.1 - rough work-around for Mavericks bug where using a list for property list item value wipes out data
	
	if class of plistItemValue is class of {"a", "b"} and AppleScript version of (system info) as number ³ 2.3 then
		-- Convert each list item into a string and escape it for the shell command:
		-- This will fail for any data types that AppleScript cannot coerce directly into a string.
		set plistItemValue_forShell to ""
		repeat with oneItem in plistItemValue
			set plistItemValue_forShell to plistItemValue_forShell & space & quoted form of (oneItem as string)
		end repeat
		set shellCommand to "defaults write " & quoted form of POSIX path of plistPath & space & plistItemName & space & "-array" & space & plistItemValue_forShell
		do shell script shellCommand
		return true
		
	else -- handle normally, since we aren't dealing with Mavericks list bug:
		
		tell application "System Events"
			-- create an empty property list dictionary item
			set the parent_dictionary to make new property list item with properties {kind:record}
			try
				set plistFile to property list file plistPath
			on error errMsg number errNum
				if errNum is -1728 then
					set plistFile to make new property list file with properties {contents:parent_dictionary, name:plistPath}
				else
					error errMsg number errNum
				end if
			end try
			tell plistFile
				try
					
					tell property list item plistItemName
						set value to plistItemValue
					end tell
				on error errMsg number errNum
					-- 
					if errNum is -10006 then
						make new property list item at Â
							end of property list items of contents of plistFile Â
							with properties Â
							{kind:class of plistItemValue, name:plistItemName, value:plistItemValue}
					else
						error errMsg number errNum
					end if
				end try
			end tell
			return true
		end tell
	end if
end plistWrite





on testPathExists(inputPath)
	-- version 1.5
	-- from Richard Morton, on applescript-users@lists.apple.com
	-- public domain, of course. :-)
	-- gets somewhat slower as nested-depth level goes over 10 nested folders
	if inputPath is not equal to "" then try
		get alias (inputPath as string) -- just in case inputPath was not string
		return true
	end try
	return false
end testPathExists

on parseChars(prefs)
	-- version 1.3
	
	set defaultPrefs to {considerCase:true}
	
	
	if class of prefs is list then
		if (count of prefs) is greater than 2 then
			-- get any parameters after the initial 3
			set prefs to {sourceTEXT:item 1 of prefs, parseString:item 2 of prefs, considerCase:item 3 of prefs}
		else
			set prefs to {sourceTEXT:item 1 of prefs, parseString:item 2 of prefs}
		end if
		
	else if class of prefs is not equal to (class of {someKey:3}) then
		-- Test by matching class to something that IS a record to avoid FileMaker namespace conflict with the term "record"
		
		error "The parameter for 'parseChars()' should be a record or at least a list. Wrap the parameter(s) in curly brackets for easy upgrade to 'parseChars() version 1.3. " number 1024
		
	end if
	
	
	set prefs to prefs & defaultPrefs
	
	
	set sourceTEXT to sourceTEXT of prefs
	set parseString to parseString of prefs
	set considerCase to considerCase of prefs
	
	
	set oldDelims to AppleScript's text item delimiters
	try
		set AppleScript's text item delimiters to the {parseString as string}
		
		if considerCase then
			considering case
				set the parsedList to every text item of sourceTEXT
			end considering
		else
			ignoring case
				set the parsedList to every text item of sourceTEXT
			end ignoring
		end if
		
		set AppleScript's text item delimiters to oldDelims
		return parsedList
	on error errMsg number errNum
		try
			set AppleScript's text item delimiters to oldDelims
		end try
		error "ERROR: parseChars() handler: " & errMsg number errNum
	end try
end parseChars

on replaceSimple(prefs)
	-- version 1.4
	
	set defaultPrefs to {considerCase:true}
	
	if class of prefs is list then
		if (count of prefs) is greater than 3 then
			-- get any parameters after the initial 3
			set prefs to {sourceTEXT:item 1 of prefs, oldChars:item 2 of prefs, newChars:item 3 of prefs, considerCase:item 4 of prefs}
		else
			set prefs to {sourceTEXT:item 1 of prefs, oldChars:item 2 of prefs, newChars:item 3 of prefs}
		end if
		
	else if class of prefs is not equal to (class of {someKey:3}) then
		-- Test by matching class to something that IS a record to avoid FileMaker namespace conflict with the term "record"
		
		error "The parameter for 'replaceSimple()' should be a record or at least a list. Wrap the parameter(s) in curly brackets for easy upgrade to 'replaceSimple() version 1.3. " number 1024
		
	end if
	
	
	set prefs to prefs & defaultPrefs
	
	
	set considerCase to considerCase of prefs
	set sourceTEXT to sourceTEXT of prefs
	set oldChars to oldChars of prefs
	set newChars to newChars of prefs
	
	set sourceTEXT to sourceTEXT as string
	
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to the oldChars
	if considerCase then
		considering case
			set the parsedList to every text item of sourceTEXT
			set AppleScript's text item delimiters to the {(newChars as string)}
			set the newText to the parsedList as string
		end considering
	else
		ignoring case
			set the parsedList to every text item of sourceTEXT
			set AppleScript's text item delimiters to the {(newChars as string)}
			set the newText to the parsedList as string
		end ignoring
	end if
	set AppleScript's text item delimiters to oldDelims
	return newText
	
end replaceSimple

