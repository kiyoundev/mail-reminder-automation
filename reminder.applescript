-- keywords to look for in the email body
property keyword_balance : {"statement balance"}
property keyword_due : {"due date", "payment due date"}
property keyword_minpayment : {"minimum payment", "minimum payment due"}

-- name of the reminder list the item should be added to
property defaultListName : "Finance"
-- prefix for reminder titles
property reminderPrefix : "Bill Payment Due"
-- priority for created reminders
property defaultPriority : 1
-- whether to flag created reminders
property defaultFlagged : true

(*
Mail rule entry point
This handler is called by Mail when the rule matches incoming messages
It processes each message and creates reminders for billing emails
*)
using terms from application "Mail"
	on perform mail action with messages theMessages for rule theRule
		repeat with aMessage in theMessages
			try
				my processBillingEmail(aMessage)
			on error errMsg number errNum
				do shell script "/usr/bin/logger " & quoted form of ("Mail rule error (" & errNum & "): " & errMsg)
			end try
		end repeat
	end perform mail action with messages
end using terms from

(*
Process an email message and create a reminder with the relevant properties.

Parameters:
aMessage - the email message to process

Returns:
none
*)
on processBillingEmail(aMessage)
	
	tell application "Mail"
		set senderField to extract address from (sender of aMessage)
		set subjectField to the subject of aMessage
		set bodyField to the content of aMessage
	end tell
	
	set parsedResult to my parse(senderField, subjectField, bodyField)
	if parsedResult is missing value then return
	set {bankName, balance, dateDue, minPayment} to {bankName, balance, dateDue, minPayment} of parsedResult
	
	my createReminderFor(bankName, dateDue, balance, minPayment)
	
end processBillingEmail


(*
Create a new reminder in the default list with the given title and properties.
If dateDue parsing fails, it will use the current date.

Parameters:
bankName - the name of the bank (optional)
dateDue - the due date of the payment
balance - the current balance
minPayment - the minimum payment due

Returns:
none
*)
on createReminderFor(bankName, dateDue, balance, minPayment)
	if bankName is missing value then
		set reminderTitle to reminderPrefix
	else
		set reminderTitle to bankName & ": " & reminderPrefix
	end if
	-- create a new reminder in the default list with the given title and properties. 
	-- If the default list does not exist, it will be created. 
	-- The reminder will have the given priority, flagged status, body, and due date.
	tell application "Reminders"
		if not (exists list defaultListName) then
			make new list with properties {name:defaultListName}
		end if
		set targetList to list defaultListName
		set newReminder to make new reminder at end of reminders of targetList with properties {name:reminderTitle}
		tell newReminder
			set priority to defaultPriority
			set flagged to defaultFlagged
			set body to ("Balance: " & balance & return & "Minimum Payment: " & minPayment)
			-- safely set the due date with fallback
			try
				set allday due date to date dateDue
			on error
				-- if date parsing fails, use a default current date
				set allday due date to (current date)
			end try
		end tell
	end tell
end createReminderFor

(*
Parse an email message and extract the relevant information.

Parameters:
senderField - the sender field of the email message
subjectField - the subject field of the email message
bodyField - the body of the email message

Returns:
A dictionary with the following keys:
bankName - the name of the bank, as extracted from the sender field
balance - the current balance of the account
dateDue - the due date of the payment
minPayment - the minimum payment due

If any of the above values are missing, missing value is returned.
*)
on parse(senderField, subjectField, bodyField)
	set cleanBody to my normalizeLineEndings(bodyField)
	set balance to my firstValueForKeywords(cleanBody, keyword_balance)
	set dateDue to my firstValueForKeywords(cleanBody, keyword_due)
	set minPayment to my firstValueForKeywords(cleanBody, keyword_minpayment)
	
	if balance is missing value or dateDue is missing value or minPayment is missing value then return missing value
	
	-- extract bank name from sender email address
	set bankName to my extractBankName(senderField)
	
	return {bankName:bankName, balance:balance, dateDue:date dateDue, minPayment:minPayment}
end parse


(*
Extract the name of the bank from the sender field of an email message.

Parameters:
senderField - the sender field of the email message

Returns:
The name of the bank, as a string. If the sender field does not contain a valid email address, "Unknown Bank" is returned.
*)
on extractBankName(senderField)
	-- extract email address from sender field
	tell application "Mail"
		set emailAddress to extract address from senderField
	end tell
	
	-- extract domain from email address
	set AppleScript's text item delimiters to "@"
	set parts to text items of emailAddress
	set AppleScript's text item delimiters to ""
	
	if (count parts) > 1 then
		set domain to item 2 of parts
		-- extract bank name from domain (before first dot)
		set AppleScript's text item delimiters to "."
		set domainParts to text items of domain
		set AppleScript's text item delimiters to ""
		set bankName to item 1 of domainParts
		
		-- capitalize first letter
		set firstChar to character 1 of bankName
		set restChars to characters 2 thru -1 of bankName as text
		set bankName to my uppercaseText(bankName)
		
		return bankName
	end if
	
	return "Unknown Bank"
end extractBankName

(*
Searches a given body text for the first value associated with any of the given keywords.

Parameters:
bodyText - the text to search
keywordList - a list of keywords to search for

Returns:
The first value associated with any of the given keywords, as a string. If no value is found, missing value is returned.

Notes:
The search is case-insensitive.
The value associated with a keyword is the text following the keyword on the same line, up to the next newline or line separator.
If a keyword is found on a line by itself, the associated value is the text on the following line, up to the next newline or line separator.
If a keyword is not found, the search continues with the next keyword in the list.
*)
on firstValueForKeywords(bodyText, keywordList)
	repeat with keywordText in keywordList
		
		set colonHit to my valueAfterColon(bodyText, keywordText)
		if colonHit is not missing value then return colonHit
		
		set blankHit to my valueAfterEmptyLine(bodyText, keywordText)
		if blankHit is not missing value then return blankHit
	end repeat
	return missing value
end firstValueForKeywords

(*
Searches a given body text for the first value associated with any of the given keywords.

Parameters:
bodyText - the text to search
marker - the keyword to search for

Returns:
The first value associated with the given keyword, as a string. If no value is found, missing value is returned.

Notes:
The search is case-insensitive.
The value associated with a keyword is the text following the keyword on the same line, up to the next newline or line separator.
If a keyword is found on a line by itself, the associated value is the text on the following line, up to the next newline or line separator.
*)
on valueAfterEmptyLine(bodyText, marker)
	set markerLower to my lowercaseText(marker)
	repeat with idx from 1 to (count paragraphs of bodyText) - 1
		set thisLine to paragraph idx of bodyText
		set nextLine to paragraph (idx + 1) of bodyText
		if my trimText(thisLine) is equal to markerLower then return my trimText(nextLine)
	end repeat
	return missing value
end valueAfterEmptyLine

(*
Searches a given body text for the first value associated with a given keyword.

Parameters:
bodyText - the text to search
marker - the keyword to search for

Returns:
The first value associated with the given keyword, as a string. If no value is found, missing value is returned.

Notes:
The search is case-insensitive.
The value associated with a keyword is the text following the keyword on the same line, up to the next newline or line separator.
*)
on valueAfterColon(bodyText, marker)
	set markerLower to my lowercaseText(marker)
	
	repeat with lineText in paragraphs of my lowercaseText(bodyText)
		set trimmedLineLower to my trimText(lineText)
		if trimmedLineLower contains markerLower then
			set AppleScript's text item delimiters to ":"
			set parts to text items of trimmedLineLower
			set AppleScript's text item delimiters to ""
			if (count parts) > 1 then return my trimText(item 2 of parts)
		end if
	end repeat
	return missing value
end valueAfterColon

(*
Normalize line endings in a given text
Replaces all occurrences of return, line separator, and paragraph separator with linefeed
Returns the modified text
*)
on normalizeLineEndings(t)
	if t is missing value then return t
	set cleaned to my replaceText(t, return, linefeed)
	set cleaned to my replaceText(cleaned, (character id 8232), linefeed) -- LINE SEPARATOR
	set cleaned to my replaceText(cleaned, (character id 8233), linefeed) -- PARAGRAPH SEPARATOR
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to linefeed
	set parts to text items of cleaned
	set normalized to parts as text
	set AppleScript's text item delimiters to oldDelims
	return normalized
end normalizeLineEndings

(*
Replace all occurrences of targetText in sourceText with replacementText
Returns the modified text
*)
on replaceText(sourceText, targetText, replacementText)
	if sourceText is missing value then return sourceText
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to targetText
	set parts to text items of sourceText
	set AppleScript's text item delimiters to replacementText
	set newText to parts as text
	set AppleScript's text item delimiters to oldDelims
	return newText
end replaceText

(*
Converts a given text to lowercase
Returns the modified text
*)
on lowercaseText(t)
	return do shell script "/bin/echo " & quoted form of t & " | /usr/bin/tr '[:upper:]' '[:lower:]'"
end lowercaseText

(*
Converts a given text to uppercase
Returns the modified text
*)
on uppercaseText(t)
	if t is missing value then return t
	return do shell script "/bin/echo " & quoted form of t & " | /usr/bin/tr '[:lower:]' '[:upper:]'"
end uppercaseText

(*
Trim whitespace from the beginning and end of a given text
Returns the modified text
*)
on trimText(t)
	return do shell script "/bin/echo " & quoted form of t & " | /usr/bin/sed -E 's/^[[:space:]]+//; s/[[:space:]]+$//'"
end trimText