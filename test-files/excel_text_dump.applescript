-- Prints the text Excel displays in every cell of each sheet's used range (first 300 rows), as
-- JSON lines {"sheet", "ref", "text"}, for comparing the viewer's output with Excel's:
--
--     osascript test-files/excel_text_dump.applescript "$PWD/test-files/excel-sample.xlsx" > out.jsonl

on run argv
	set path_ to item 1 of argv
	set out to {}
	tell application "Microsoft Excel"
		open (POSIX file path_)
		set wb to active workbook
		repeat with k from 1 to count of worksheets of wb
			set ws to worksheet k of wb
			set sheetName to name of ws
			set ur to used range of ws
			set r0 to first row index of ur
			set c0 to first column index of ur
			set nr to count of rows of ur
			if nr > 300 then set nr to 300
			set nc to count of columns of ur
			repeat with i from 0 to nr - 1
				repeat with j from 0 to nc - 1
					set cel to range (my colName(c0 + j) & (r0 + i)) of ws
					set txt to string value of cel
					if txt is not "" then
						set end of out to "{\"sheet\":" & my q(sheetName) & ",\"r\":" & (r0 + i - 1) & ",\"c\":" & (c0 + j - 1) & ",\"text\":" & my q(txt) & "}"
					end if
				end repeat
			end repeat
		end repeat
		close active workbook saving no
	end tell
	set AppleScript's text item delimiters to linefeed
	return out as text
end run

on q(s)
	set s to s as text
	set AppleScript's text item delimiters to "\\"
	set parts to text items of s
	set AppleScript's text item delimiters to "\\\\"
	set s to parts as text
	set AppleScript's text item delimiters to "\""
	set parts to text items of s
	set AppleScript's text item delimiters to "\\\""
	set s to parts as text
	set AppleScript's text item delimiters to ""
	return "\"" & s & "\""
end q

on colName(n)
	set s to ""
	repeat while n > 0
		set m to (n - 1) mod 26
		set s to (character (m + 1) of "ABCDEFGHIJKLMNOPQRSTUVWXYZ") & s
		set n to (n - 1) div 26
	end repeat
	return s
end colName
