-- Calendar Busy Updater

(*
	Take a list of source calendars, pulls each event from them, then blocks out generic "busy" events on a target calendar. Example use would be making a generic calendar that a service like Calendly has access to. 
	Note: the prefs are asked the 1st time this runs, saved in a plist file in ~/Library/Preferences. Currently, the only way to change those settings after the initial run is to edit or delete that plist file. 
	
HISTORY: 
	2023-04-20 ( danshock ): time 1730: now deletes any generic events from the target calendar that are no longer confirmed by the source calendars. 
	2023-04-20 ( danshockley ): First created. 
	
*)

-- internal prefs (maybe move to plist?):
property doRemoveOutdatedBusyEvents : true

-- Do not need to modify these properties:
property prefsDomain : "com.GothamDataWorks.CalendarBusyUpdater"

property timestampFormat : "YYYY-MM-DD hh-mm-ss"

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
	
	return result
	
end run


on updateCalendar()
	
	(*
	LOGIC: 
	⁃	Get a list of existing generic busy events from the target. 
	⁃	Loop over events of source calendars, for each event: 
	⁃		If it IS in the old target list, add it to the "confirmed" list.
	⁃		If it is NOTin the old target list, just create it. 
	*)
	
	set searchDateRangeStart to (current date) - daysBack * days
	set searchDateRangeEnd to (current date) + daysFuture * days
	
	tell application "Calendar"
		
		-- get all the OLD target generic events (ignore any custom events with some non-generic summary/title):
		set oldTargetEventList to (every event of calendar calendarTargetName whose start date is greater than or equal to searchDateRangeStart and start date is less than or equal to searchDateRangeEnd and summary is genericBusyEventTitle)
		
		-- Now, build up a lists of OLD target event ranges, and of range/event pairs:
		set oldTargetRangeList to {}
		set oldTargetPairs to {}
		repeat with oneOldTargetEvent in oldTargetEventList
			set oldStart to start date of oneOldTargetEvent
			set oldEnd to end date of oneOldTargetEvent
			set oneOldRange to my dateAsCustomString(oldStart, timestampFormat) & (ASCII character 9) & my dateAsCustomString(oldEnd, timestampFormat)
			copy oneOldRange to end of oldTargetRangeList
			copy {oneOldRange, oneOldTargetEvent} to end of oldTargetPairs
		end repeat
		
		-- Loop over source events, either creating in target or adding to confirmedTargetRangeList
		set confirmedTargetRangeList to {}
		repeat with oneCalendarName in calendarSourceNames
			set oneCalendarName to contents of oneCalendarName
			
			set oneSourceEventList to (every event of calendar oneCalendarName whose start date is greater than or equal to searchDateRangeStart and start date is less than or equal to searchDateRangeEnd)
			repeat with oneEvent in oneSourceEventList
				
				set oneStartDate to start date of oneEvent
				set oneEndDate to end date of oneEvent
				set oneNewRange to my dateAsCustomString(oneStartDate, timestampFormat) & (ASCII character 9) & my dateAsCustomString(oneEndDate, timestampFormat)
				if oldTargetRangeList contains oneNewRange then
					copy oneNewRange to end of confirmedTargetRangeList
				else
					tell calendar calendarTargetName
						make new event with properties {summary:genericBusyEventTitle, start date:oneStartDate, end date:oneEndDate}
					end tell
				end if
				
			end repeat
		end repeat
		
		-- Now, compare the old targets to the new (good) targets, and delete any that are no longer relevant: 
		repeat with oneOldPair in oldTargetPairs
			set {oneOldRange, oneOldTargetEvent} to oneOldPair
			if confirmedTargetRangeList does not contain oneOldRange then
				delete oneOldTargetEvent
			end if
			
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
	
	if class of plistItemValue is class of {"a", "b"} and AppleScript version of (system info) as number ≥ 2.3 then
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
						make new property list item at ¬
							end of property list items of contents of plistFile ¬
							with properties ¬
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



on dateAsCustomString(incomingDate, stringFormat)
	-- version 1.2b, Daniel A. Shockley http://www.danshockley.com
	-- 1.2b added am/pm option, and date class checking, with a nod to Arthur J. Knapp
	-- NEEDS replaceSimple() handler
	
	-- takes any form of MM, DD, YYYY, YY
	-- AND any form of hh, mm, ss 
	--  (optional ap or AP, which gives am/pm or AM/PM})
	-- leaving off am/pm option coerces to military time
	-- MUST USE LOWER-CASE for TIME!!!! (avoids month/minute conflict)
	
	-- use single letters to allow single digits, where applicable
	
	-- textVars are each always 2 digits, whereas month and day 
	-- may be 1 digit, and year will normally be 4 digit
	
	if class of incomingDate is not date then
		try
			set incomingDate to date incomingDate
		on error
			set incomingDate to (current date)
		end try
	end if
	
	set numHours to (time of incomingDate) div hours
	set textHours to text -2 through -1 of ("0" & (numHours as string))
	
	set numMinutes to (time of incomingDate) mod hours div minutes
	set textMinutes to text -2 through -1 of ("0" & (numMinutes as string))
	
	set numSeconds to (time of incomingDate) mod minutes
	set textSeconds to text -2 through -1 of ("0" & (numSeconds as string))
	
	set numDay to day of incomingDate as number
	set textDay to text -2 through -1 of ("0" & (numDay as string))
	
	set numYear to year of incomingDate as number
	set textYear to text -2 through -1 of (numYear as string)
	
	-- Emmanuel Levy's Plain Vanilla get month number function
	copy incomingDate to b
	set the month of b to January
	set numMonth to (1 + (incomingDate - b + 1314864) div 2629728)
	set textMonth to text -2 through -1 of ("0" & (numMonth as string))
	
	set customDateString to stringFormat
	
	if numHours > 12 and (customDateString contains "ap" or customDateString contains "AP") then
		-- (afternoon) and requested am/pm
		set numHours to numHours - 12 -- pull off the military 12 hours for pm hours
		set textHours to text -2 through -1 of ("0" & (numHours as string))
	end if
	
	set customDateString to replaceSimple({customDateString, "MM", textMonth})
	set customDateString to replaceSimple({customDateString, "DD", textDay})
	set customDateString to replaceSimple({customDateString, "YYYY", numYear as string})
	set customDateString to replaceSimple({customDateString, "hh", textHours})
	
	set customDateString to replaceSimple({customDateString, "mm", textMinutes})
	set customDateString to replaceSimple({customDateString, "ss", textSeconds})
	
	
	-- shorter options
	set customDateString to replaceSimple({customDateString, "M", numMonth})
	set customDateString to replaceSimple({customDateString, "D", numDay})
	set customDateString to replaceSimple({customDateString, "YY", textYear})
	set customDateString to replaceSimple({customDateString, "h", numHours})
	set customDateString to replaceSimple({customDateString, "m", numMinutes})
	set customDateString to replaceSimple({customDateString, "s", numSeconds})
	
	-- AM/PM MUST be after Minutes/Month done, since it adds an M
	if (time of incomingDate) > (12 * hours) then
		-- afternoon
		set customDateString to replaceSimple({customDateString, "ap", "pm"})
		set customDateString to replaceSimple({customDateString, "AP", "PM"})
	else
		set customDateString to replaceSimple({customDateString, "ap", "am"})
		set customDateString to replaceSimple({customDateString, "AP", "AM"})
	end if
	
	return customDateString
	
end dateAsCustomString