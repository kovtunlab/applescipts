use scripting additions

-- Define the file path to save the CSV file
set savePath to ((path to desktop as string) & "40_23_shortlist.csv")

-- Initialize duplicate tracking
set seenEmails to {}
set processedCount to 0

-- Open the file for writing
set fileRef to open for access file savePath with write permission

try
	-- Write the header row to the CSV file
	write "Sender Name,Sender Email" & linefeed to fileRef
	
	tell application "Microsoft Outlook"
		-- Check if there is an email selected
		set selectedMessages to selection
		if selectedMessages is missing value or (count of selectedMessages) = 0 then
			display dialog "Please select one or more emails first!"
			close access fileRef
			return
		end if
		
		display dialog "Found " & (count of selectedMessages) & " selected items. Processing..." buttons {"OK"} default button "OK"
		
		-- Loop through each selected message
		repeat with theMessage in selectedMessages
			try
				-- Test if this is actually an email message by checking if it has a sender
				set messageSender to sender of theMessage
				
				-- If we get here, it's an email message
				set senderName to "Unknown"
				set senderEmail to "Unknown"
				
				if messageSender is not missing value then
					try
						set senderName to name of messageSender
					end try
					try
						set senderEmail to address of messageSender
					end try
				end if
				
				-- Skip if no email address found
				if senderEmail is "Unknown" or senderEmail is "" then
					log "Skipping item with no email address"
				else
					-- Check for duplicates
					if seenEmails does not contain senderEmail then
						-- Add to seen list
						set end of seenEmails to senderEmail
						
						-- Remove commas from the sender's name
						set senderName to my replace_chars(senderName, ",", "")
						
						-- Capitalize each word in the sender's name
						set senderName to my capitalize_words(senderName)
						
						-- Write the data to the CSV file as a comma-separated line
						write "\"" & senderName & "\",\"" & senderEmail & "\"" & linefeed to fileRef
						
						set processedCount to processedCount + 1
					else
						log "Skipping duplicate email: " & senderEmail
					end if
				end if
				
			on error errMsg
				-- This item is not an email message, skip it
				log "Skipping non-email item: " & errMsg
			end try
		end repeat
	end tell
	
on error errMsg
	display dialog "Error occurred: " & errMsg buttons {"OK"} default button "OK"
end try

-- Close the file
close access fileRef

if processedCount > 0 then
	display dialog "Export completed! Processed " & processedCount & " unique emails. File saved to desktop as 40_23_shortlist.csv." buttons {"OK"} default button "OK"
else
	display dialog "No valid email messages were processed. Check that you've selected actual email messages." buttons {"OK"} default button "OK"
end if

-- Function to replace characters in a string
on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the_items to text items of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the_items as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars

-- Function to capitalize the first letter of each word
on capitalize_words(this_text)
	try
		set AppleScript's text item delimiters to space
		set the_words to every text item of this_text
		set AppleScript's text item delimiters to ""
		set capitalized_text to ""
		repeat with a_word in the_words
			if length of (a_word as string) > 0 then
				set capitalized_text to capitalized_text & (do shell script "echo " & quoted form of a_word & " | awk '{print toupper(substr($0,1,1))tolower(substr($0,2))}'") & space
			end if
		end repeat
		if length of capitalized_text > 0 then
			return text 1 thru -2 of capitalized_text -- Remove trailing space
		else
			return this_text
		end if
	on error
		return this_text -- Return original if capitalization fails
	end try
end capitalize_words
