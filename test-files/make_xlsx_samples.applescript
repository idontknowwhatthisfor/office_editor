-- Builds XLSX test workbooks with Microsoft Excel itself, so they are saved exactly as Excel
-- saves them (shared formulas, cached values, styles, theme colours):
--
--     osascript test-files/make_xlsx_samples.applescript && mv ~/Downloads/excel-*.xlsx test-files/
--
-- Excel is sandboxed and may not write into this folder, so the files go to ~/Downloads first.
-- It may ask once for access to that folder.

set outDir to POSIX path of (path to downloads folder)

tell application "Microsoft Excel"
	set display alerts to false

	-- ---------------------------------------------------------------- excel-sample.xlsx
	set wb to make new workbook
	set ws to worksheet 1 of wb
	set name of ws to "Budget"
	set ws to worksheet "Budget" of wb
	set value of range "A1:G1" of ws to {{"Date", "Item", "Category", "Price", "Qty", "Total", "Note"}}
	set formula of range "A2:A201" of ws to "=DATE(2026,1,1)+ROW()*3"
	set formula of range "B2:B201" of ws to "=\"Item \"&(ROW()-1)"
	set formula of range "C2:C201" of ws to "=CHOOSE(MOD(ROW(),4)+1,\"Food\",\"Rent\",\"Travel\",\"Other\")"
	set formula of range "D2:D201" of ws to "=ROUND(MOD(ROW()*37.3,500),2)"
	set formula of range "E2:E201" of ws to "=MOD(ROW(),7)+1"
	-- Turn A:E into plain values; F stays formulas (Excel writes them as shared formulas).
	set value of range "A2:E201" of ws to value of range "A2:E201" of ws
	set formula of range "F2:F201" of ws to "=D2*E2"
	set value of range "G3" of ws to "This note is long enough to spill over the empty cells to its right."
	set value of range "G5" of ws to "Short"
	set value of range "H5" of ws to "Blocks spill"
	set value of range "G4" of ws to "This note is cut off by H4"
	set value of range "H4" of ws to "x"
	set number format of range "A2:A201" of ws to "d mmm yyyy"
	set number format of range "D2:D201" of ws to "#,##0.00"
	set number format of range "F2:F201" of ws to "$#,##0.00;[Red]-$#,##0.00"
	set bold of font object of range "A1:G1" of ws to true
	set color of interior object of range "A1:G1" of ws to {31, 78, 121}
	set color of font object of range "A1:G1" of ws to {255, 255, 255}
	set value of range "A203" of ws to "Total"
	set bold of font object of range "A203:F203" of ws to true
	set formula of range "F203" of ws to "=SUM(F2:F201)"
	set number format of range "F203" of ws to "$#,##0.00"
	set value of range "A204" of ws to "Average price"
	set formula of range "D204" of ws to "=AVERAGE(D2:D201)"
	set value of range "A205" of ws to "Rows over 1,000"
	set formula of range "F205" of ws to "=COUNTIF(F2:F201,\">1000\")"
	set value of range "A206" of ws to "Big spend?"
	set formula of range "F206" of ws to "=IF(F203>100000,\"Yes\",\"No\")"
	set value of range "A207" of ws to "Price of Item 7"
	set formula of range "D207" of ws to "=VLOOKUP(\"Item 7\",B2:D201,3,FALSE)"
	set value of range "A208" of ws to "Refund"
	set value of range "F208" of ws to -125.5
	set number format of range "F208" of ws to "$#,##0.00;[Red]-$#,##0.00"
	set column width of range "A:A" of ws to 14
	set column width of range "B:B" of ws to 12
	set column width of range "G:G" of ws to 10
	activate object ws
	select range "B2" of ws
	set freeze panes of active window to true

	-- Formats: number formats, fonts, fills, borders, alignment, merges.
	set ws2 to make new worksheet at end of wb
	set name of ws2 to "Formats"
	set ws2 to worksheet "Formats" of wb
	set value of range "A1" of ws2 to "Number formats and styles"
	merge range "A1:D1" of ws2
	set horizontal alignment of range "A1:D1" of ws2 to horizontal align center
	set font size of font object of range "A1" of ws2 to 16
	set bold of font object of range "A1" of ws2 to true
	set value of range "A3:B3" of ws2 to {{"Format", "Value"}}
	set bold of font object of range "A3:B3" of ws2 to true
	set fmts to {"General", "General", "General", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??", "m/d/yyyy", "d-mmm-yy", "dddd, mmmm d, yyyy", "h:mm AM/PM", "h:mm:ss", "[h]:mm", "yyyy-mm-dd hh:mm", "$#,##0.00", "#,##0.00 [$€-x-euro1]", "[Blue]#,##0;[Red]-#,##0;[Green]0", "[Blue]#,##0;[Red]-#,##0;[Green]0", "[Blue]#,##0;[Red]-#,##0;[Green]0", "@", "0.000", "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"}
	set vals to {1234.5678, 0.000123, 123456789012, 1234.5678, 1234.5678, 1234567.891, -1234567.891, 0.256, 0.256, 123456789, 1.25, 3.14159, 46337.75, 46337.75, 46337.75, 0.6875, 0.6875, 1.75, 46337.75, -42, 1234.5, 1500, -1500, 0, "00123", 2.5, -1234.5}
	repeat with i from 1 to count of fmts
		set r to i + 3
		set value of range ("A" & r) of ws2 to "'" & (item i of fmts)
		set number format of range ("B" & r) of ws2 to item i of fmts
		set value of range ("B" & r) of ws2 to item i of vals
	end repeat
	set value of range "A32" of ws2 to "Boolean"
	set formula of range "B32" of ws2 to "=1<2"
	set value of range "A33" of ws2 to "Error"
	set formula of range "B33" of ws2 to "=1/0"
	set value of range "A34" of ws2 to "Text formula"
	set formula of range "B34" of ws2 to "=UPPER(\"done\")&\" \"&TEXT(0.5,\"0%\")"

	set value of range "D3" of ws2 to "Fonts"
	set bold of font object of range "D3" of ws2 to true
	set value of range "D4" of ws2 to "Arial 10"
	set name of font object of range "D4" of ws2 to "Arial"
	set font size of font object of range "D4" of ws2 to 10
	set value of range "D5" of ws2 to "Times New Roman 14 italic"
	set name of font object of range "D5" of ws2 to "Times New Roman"
	set font size of font object of range "D5" of ws2 to 14
	set italic of font object of range "D5" of ws2 to true
	set value of range "D6" of ws2 to "Underlined red"
	set underline of font object of range "D6" of ws2 to underline style single
	set color of font object of range "D6" of ws2 to {192, 0, 0}
	set value of range "D7" of ws2 to "Struck through"
	set strikethrough of font object of range "D7" of ws2 to true
	set value of range "D8" of ws2 to "Theme accent fill"
	set theme color of interior object of range "D8" of ws2 to first accent theme color
	set tint and shade of interior object of range "D8" of ws2 to 0.6
	set value of range "D9" of ws2 to "Yellow fill"
	set color of interior object of range "D9" of ws2 to {255, 255, 0}

	set value of range "D11" of ws2 to "Alignment"
	set bold of font object of range "D11" of ws2 to true
	set value of range "D12" of ws2 to "Right"
	set horizontal alignment of range "D12" of ws2 to horizontal align right
	set value of range "D13" of ws2 to "Center"
	set horizontal alignment of range "D13" of ws2 to horizontal align center
	set value of range "D14" of ws2 to "Wrapped text that runs over several lines inside one cell"
	set wrap text of range "D14" of ws2 to true
	set value of range "D15" of ws2 to "Top aligned"
	set vertical alignment of range "D15" of ws2 to vertical alignment top
	set row height of range "15:15" of ws2 to 36
	set value of range "D16" of ws2 to "Plain"

	set value of range "D18" of ws2 to "Merged block"
	merge range "D18:E20" of ws2
	set horizontal alignment of range "D18:E20" of ws2 to horizontal align center
	set vertical alignment of range "D18:E20" of ws2 to vertical alignment center
	set color of interior object of range "D18:E20" of ws2 to {226, 239, 218}

	set value of range "D22" of ws2 to "Borders"
	set bold of font object of range "D22" of ws2 to true
	set value of range "D23:F25" of ws2 to {{"thin", "medium", "thick"}, {"1", "2", "3"}, {"dashed", "double", "dotted"}}
	repeat with e in {edge left, edge top, edge right, edge bottom, inside vertical, inside horizontal}
		set weight of (get border (range "D23:D24" of ws2) which border e) to border weight thin
		set weight of (get border (range "E23:E24" of ws2) which border e) to border weight medium
		set weight of (get border (range "F23:F24" of ws2) which border e) to border weight thick
	end repeat
	set line style of (get border (range "D25" of ws2) which border edge bottom) to dash
	set line style of (get border (range "E25" of ws2) which border edge bottom) to double
	set line style of (get border (range "F25" of ws2) which border edge bottom) to dot
	set color of (get border (range "F23:F24" of ws2) which border edge bottom) to {0, 112, 192}

	set column width of range "A:A" of ws2 to 34
	set column width of range "B:B" of ws2 to 22
	set column width of range "D:D" of ws2 to 24
	set column width of range "C:C" of ws2 to 3

	set ws3 to make new worksheet at end of wb
	set name of ws3 to "Hidden"
	set ws3 to worksheet "Hidden" of wb
	set value of range "A1" of ws3 to "This sheet is hidden"
	set visible of ws3 to sheet hidden

	set ws4 to make new worksheet at end of wb
	set name of ws4 to "Wide & tall"
	set ws4 to worksheet "Wide & tall" of wb
	set value of range "A1" of ws4 to "Row 1 and column A are frozen; columns C:D are hidden, row 4 is hidden."
	set formula of range "B2:Z300" of ws4 to "=ADDRESS(ROW(),COLUMN(),4)"
	set formula of range "A2:A300" of ws4 to "=ROW()"
	set column width of range "C:D" of ws4 to 0
	set row height of range "4:4" of ws4 to 0
	activate object ws4
	select range "B2" of ws4
	set freeze panes of active window to true
	set display gridlines of active window to false

	activate object ws
	save workbook as wb filename (outDir & "excel-sample.xlsx") file format Excel XML file format
	close active workbook saving no

	-- ---------------------------------------------------------------- excel-5000-rows.xlsx
	set wb to make new workbook
	set ws to worksheet 1 of wb
	set name of ws to "Data"
	set ws to worksheet "Data" of wb
	set value of range "A1:J1" of ws to {{"ID", "Region", "Product", "Units", "Price", "Revenue", "Cost", "Profit", "Margin", "Date"}}
	set bold of font object of range "A1:J1" of ws to true
	set formula of range "A2:A5001" of ws to "=ROW()-1"
	set formula of range "B2:B5001" of ws to "=CHOOSE(MOD(ROW(),5)+1,\"North\",\"South\",\"East\",\"West\",\"Central\")"
	set formula of range "C2:C5001" of ws to "=\"Product \"&CHAR(65+MOD(ROW()*7,26))"
	set formula of range "D2:D5001" of ws to "=MOD(ROW()*13,97)+1"
	set formula of range "E2:E5001" of ws to "=ROUND(5+MOD(ROW()*3.7,95),2)"
	set formula of range "J2:J5001" of ws to "=DATE(2025,1,1)+MOD(ROW()*11,365)"
	set value of range "A2:E5001" of ws to value of range "A2:E5001" of ws
	set value of range "J2:J5001" of ws to value of range "J2:J5001" of ws
	set formula of range "F2:F5001" of ws to "=D2*E2"
	set formula of range "G2:G5001" of ws to "=ROUND(F2*0.62,2)"
	set formula of range "H2:H5001" of ws to "=F2-G2"
	set formula of range "I2:I5001" of ws to "=IF(F2=0,0,H2/F2)"
	set number format of range "E2:H5001" of ws to "#,##0.00"
	set number format of range "I2:I5001" of ws to "0.0%"
	set number format of range "J2:J5001" of ws to "yyyy-mm-dd"
	set column width of range "J:J" of ws to 11
	activate object ws
	select range "A2" of ws
	set freeze panes of active window to true
	save workbook as wb filename (outDir & "excel-5000-rows.xlsx") file format Excel XML file format
	close active workbook saving no
end tell
