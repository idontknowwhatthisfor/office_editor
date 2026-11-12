-- Re-saves a deck in PowerPoint and exports PowerPoint's PDF of the slides (the reference):
--
--     osascript test-files/pptx_resave.applescript "$PWD/test-files/in.pptx" name
--     mv ~/Downloads/name.pptx test-files/ && mv ~/Downloads/name.pdf test-files/
--
-- PowerPoint is sandboxed and may not write into this folder, so it writes to ~/Downloads.

on run argv
	set inPath to item 1 of argv
	set baseName to item 2 of argv
	set outDir to POSIX path of (path to downloads folder)
	tell application "Microsoft PowerPoint"
		open (POSIX file inPath)
		set p to active presentation
		save p in (POSIX file (outDir & baseName & ".pptx")) as save as Open XML presentation
		save active presentation in (POSIX file (outDir & baseName & ".pdf")) as save as PDF
		close active presentation saving no
	end tell
end run
