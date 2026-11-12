-- Builds a deck in PowerPoint with its own default theme (16:9, Aptos), then saves it and
-- PowerPoint's PDF of it to ~/Downloads:
--
--     osascript test-files/make_pptx_native.applescript
--     mv ~/Downloads/pptx-native.pptx ~/Downloads/pptx-native.pdf test-files/

set outDir to POSIX path of (path to downloads folder)
tell application "Microsoft PowerPoint"
	set p to make new presentation
	set s to make new slide at end of p with properties {layout:slide layout title slide}
	set content of text range of text frame of shape 1 of s to "Native PowerPoint deck"
	set content of text range of text frame of shape 2 of s to "Default theme, made by AppleScript"

	set s to make new slide at end of p with properties {layout:slide layout text slide}
	set content of text range of text frame of shape 1 of s to "Agenda"
	set content of text range of text frame of shape 2 of s to "Where we are" & return & "What changed" & return & "Decisions needed" & return & "Timeline"

	set s to make new slide at end of p with properties {layout:slide layout two column text}
	set content of text range of text frame of shape 1 of s to "Pros and cons"
	set content of text range of text frame of shape 2 of s to "Works offline" & return & "Free" & return & "Small download"
	set content of text range of text frame of shape 3 of s to "No editing yet" & return & "Charts not drawn"

	set s to make new slide at end of p with properties {layout:slide layout title only}
	set content of text range of text frame of shape 1 of s to "Shapes in the default theme"
	set sh to make new shape at end of s with properties {auto shape type:autoshape rounded rectangle, left position:60, top:150, width:250, height:120}
	set content of text range of text frame of sh to "Theme style"
	set sh to make new shape at end of s with properties {auto shape type:autoshape oval, left position:350, top:150, width:200, height:200}
	set content of text range of text frame of sh to "Oval"
	set sh to make new shape at end of s with properties {auto shape type:autoshape chevron, left position:600, top:170, width:250, height:100}
	set sh to make new shape at end of s with properties {auto shape type:autoshape flowchart decision, left position:100, top:320, width:200, height:150}
	set sh to make new shape at end of s with properties {auto shape type:autoshape cloud callout, left position:600, top:320, width:250, height:150}
	set content of text range of text frame of sh to "Callout"

	set s to make new slide at end of p with properties {layout:slide layout title slide}
	set content of text range of text frame of shape 1 of s to "Thank you"
	set content of text range of text frame of shape 2 of s to "Questions?"

	save p in (POSIX file (outDir & "pptx-native.pptx")) as save as Open XML presentation
	save active presentation in (POSIX file (outDir & "pptx-native.pdf")) as save as PDF
	close active presentation saving no
end tell
