-- Builds xlsx_chart.xlsx with Microsoft Excel itself: a data sheet with a frozen header row and
-- embedded charts (clustered column, line with markers, pie, stacked area, scatter), and a chart
-- sheet (stacked bar). Also saves Excel's PDF of it, the reference rendering:
--
--     osascript test-files/make_xlsx_charts.applescript && mv ~/Downloads/xlsx_chart.* test-files/
--
-- Excel is sandboxed and may not write into this folder, so the files go to ~/Downloads first.

set outDir to POSIX path of (path to downloads folder)

tell application "Microsoft Excel"
	set display alerts to false
	set wb to make new workbook
	set ws to worksheet 1 of wb
	set name of ws to "Sales"
	set ws to worksheet "Sales" of wb
	set value of range "A1:D7" of ws to {{"Month", "North", "South", "West"}, ¬
		{"Jan", 120, 80, 45}, {"Feb", 135, 95, 60}, {"Mar", 128, 110, 72}, ¬
		{"Apr", 150, 105, 90}, {"May", 162, 98, 110}, {"Jun", 171, 120, 125}}
	set value of range "F1:G6" of ws to {{"Hours", "Score"}, {1, 52}, {2.5, 61}, {3, 70}, {4.5, 74}, {6, 88}}
	set bold of font object of range "A1:G1" of ws to true
	set number format of range "B2:D7" of ws to "#,##0"

	set co to make new chart object at ws with properties {left position:20, top:130, width:380, height:220}
	set ch to chart of co
	set chart type of ch to column clustered
	chart wizard ch source (range "A1:D7" of ws)

	set co to make new chart object at ws with properties {left position:420, top:130, width:380, height:220}
	set ch to chart of co
	set chart type of ch to line markers
	chart wizard ch source (range "A1:D7" of ws)

	set co to make new chart object at ws with properties {left position:20, top:370, width:300, height:220}
	set ch to chart of co
	set chart type of ch to pie chart
	chart wizard ch source (range "A1:B7" of ws)

	set co to make new chart object at ws with properties {left position:340, top:370, width:300, height:220}
	set ch to chart of co
	set chart type of ch to area stacked
	chart wizard ch source (range "A1:D7" of ws)

	set co to make new chart object at ws with properties {left position:20, top:610, width:380, height:220}
	set ch to chart of co
	set chart type of ch to xy scatter lines
	chart wizard ch source (range "F1:G6" of ws)

	set cs to make new chart sheet at end of active workbook
	set csChart to chart of cs
	chart wizard csChart source (range "A1:D7" of ws)
	set chart type of csChart to bar stacked
	set name of cs to "Totals"

	activate object ws
	select range "A2" of ws
	set freeze panes of active window to true
	save workbook as wb filename (outDir & "xlsx_chart.xlsx") file format Excel XML file format
	save workbook as active workbook filename (outDir & "xlsx_chart.pdf") file format PDF file format
	close active workbook saving no
end tell
