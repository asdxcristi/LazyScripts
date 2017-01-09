'File shits
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")

'Change this if you have an already written file
outputFile="command.sql"

If (fso.FileExists(outputFile)) Then 'append(8) to file
   Set f = fso.OpenTextFile(outputFile, 8, True)
   
Else ' file doesn't exist => we create it and write (2) to it
   Set f = fso.CreateTextFile(outputFile, 2)
End If


While (1)

	insertCommand="INSERT INTO "
	
	tableNameAndShit = InputBox("Input the TABLE_NAME and colums Ex:wf_world_regions (region_id, region_name)")
	insertCommand = insertCommand & tableNameAndShit & vbNewLine ' vbCrLf aka newline aka "\n"


	insertCommand = insertCommand & "VALUES "
	values=InputBox("Input the values Ex:(151,'Eastern Europe')")
	insertCommand = insertCommand & values & ";"

'write the complete command to file
	f.WriteLine insertCommand

Wend

f.Close
WScript.Quit