' setup
filename = "All Text.txt"
wordPerDay = 10
wordRepeat = 5
wordBatch = 50

folderDaily = "daily"
folderBatch = "batch"

runBackup = false

' run
Set en = CreateObject("System.Collections.ArrayList")
Set vi = CreateObject("System.Collections.ArrayList")
Set tmp = CreateObject("System.Collections.ArrayList")

Set fso = CreateObject("Scripting.FileSystemObject")
Set reader = fso.OpenTextFile(filename)
Dim writer

' backup old folder
if runBackup = true then
	timenow = Now
	if fso.FolderExists(folderDaily) then
		fso.GetFolder(folderDaily).Name = "Backup " & FormatDateTime(timenow, 1) & "_" & folderDaily 
	end if
	if fso.FolderExists(folderBatch) then
		fso.GetFolder(folderBatch).Name = "Backup " & FormatDateTime(timenow, 1) & "_" & folderBatch
	end if
else
	if fso.FolderExists(folderDaily) then
		fso.DeleteFolder(folderDaily)
	end if
	if fso.FolderExists(folderBatch) then
		fso.DeleteFolder(folderBatch)
	end if
end if
' create new folder
fso.CreateFolder folderDaily
fso.CreateFolder folderBatch

i = 0
line = ""
linemod = 0
wordCount = 0
fileCount = 1
fileContent = ""
batchContent = ""

endOfFile = false

Do Until endOfFile
	if not reader.AtEndOfStream Then
		' read file content
		linemod = i mod 3
		line = reader.ReadLine
		select case linemod
			case 0:
				if len(line) = 0 then
					Msgbox "Error vn line: " & (i + 1)
					Wscript.Quit
				end if
				vi.Add(line)
			case 1:
				if len(line) = 0 then
					Msgbox "Error en line: " & (i + 1)
					Wscript.Quit
				end if
				en.Add(line)
				batchContent = batchContent & line & vbCrlf
				fileContent = fileContent & Replace(space(wordRepeat), " ", line & vbCrlf) & vbCrlf
				wordCount = wordCount + 1
			case 2:
				tmp.Add(line)
		end select
		
		i = i + 1
	else
		endOfFile = true
	end if
	
	' write batch
	if (wordCount mod wordBatch = 0) or (endOfFile = true) Then
		if len(batchContent) > 0 then
			
			Set writer = fso.CreateTextFile(folderBatch & "\Batch_" & (wordCount - 49) & "-" & wordCount & ".txt",True)
			writer.Write batchContent
			writer.Close
			
			batchContent = ""
			fileCount = fileCount + 1
			
		end if
	end if
	
	' write file
	if (wordCount mod wordPerDay = 0) or (endOfFile = true) Then
		if len(fileContent) > 0 then
			
			Set writer = fso.CreateTextFile(folderDaily & "\Day_" & fileCount & ".txt",True)
			writer.Write fileContent
			writer.Close
			
			fileContent = ""
			fileCount = fileCount + 1
			
		end if
	end if
	
Loop

reader.Close

Msgbox "Completed!"

Wscript.Quit
' debug
Msgbox "En length: " & en.Count
Msgbox "Vi length: " & vi.Count
Msgbox "Tmp length: " & tmp.Count
