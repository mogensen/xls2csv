## Use linux line feeds
set LF to {ASCII character 10}
tell application "Microsoft Excel"
	activate
	
	set outPath to (path of active workbook)
	set fileName to (name of active workbook)
	
	## Remove file extension
	set fileName to (do shell script "echo " & fileName & " | sed 's/\\.[^.]*$//' ")
	set outFile to (outPath & ":" & fileName & ".csv")
	
	## Loop params
	set lastCol to count of columns of used range of active sheet
	set lastRow to count of rows of used range of active sheet
	
	## Get the cols that are integer
	set headerCells to {}
	repeat with i from 1 to lastCol
		set headerCells to headerCells & (value of cell 1 of column i of active sheet)
	end repeat
	
	set number_cells to (choose from list headerCells with prompt "Choose integer fields:" with multiple selections allowed and empty selection allowed)
	if number_cells is false then
		return "User cancled action"
	end if
	
	## Open file for write
	set openFile to open for access file outFile with write permission
	set eof openFile to 0
	
	set rowStr to "\"" & (value of cell 1 of column 1 of active sheet)
	repeat with c from 2 to lastCol
		set cellVal to (value of cell 1 of column c of active sheet)
		set rowStr to rowStr & "\",\"" & cellVal
	end repeat
	set rowStr to rowStr & "\""
	write rowStr & LF to openFile as Çclass utf8È
	
	repeat with r from 2 to lastRow
		
		# Handle first column
		if number_cells contains (value of cell 1 of column 1 of active sheet) then
			set rowStr to "\"" & ((value of cell r of column 1 of active sheet) as integer)
		else
			set rowStr to "\"" & (value of cell r of column 1 of active sheet)
		end if
		
		# All the res off the colums
		repeat with c from 2 to lastCol
			if number_cells contains (value of cell 1 of column c of active sheet) then
				set cellVal to ((value of cell r of column c of active sheet) as integer)
			else
				set cellVal to (value of cell r of column c of active sheet)
			end if
			set rowStr to rowStr & "\",\"" & cellVal
		end repeat
		
		set rowStr to rowStr & "\""
		write rowStr & LF to openFile as Çclass utf8È
	end repeat
	close access openFile
end tell

## Choose encoding for target file
set encoding to (choose from list {"utf-8", "utf-16", "latin1"} default items "utf-16" with prompt "Choose target encoding:")
if encoding is false then
	return "No encoding choosen"
end if

## POSIX path of output file
set p to POSIX path of outFile
set quotedPath to quoted form of p

## POSIX path of encoded output file
set outFileUTF16 to (outPath & ":" & fileName & "-" & encoding & ".csv")
set q to POSIX path of outFileUTF16
set outputPath to quoted form of q

## Remove windows newlines
do shell script "perl -pi -e 's#\\r#\\\\n#g' " & quotedPath

## Convert encoding to the user defined
do shell script "iconv -f utf-8 -t " & encoding & " " & quotedPath & " > " & outputPath

## Delete temp file
do shell script "rm -f " & quotedPath