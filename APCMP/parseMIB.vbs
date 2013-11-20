'Parsing Text with SNMP Traps


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("U:\GitHub\cbc_winteam\dev\opsmgr\apc-ups\MIBParsing\traps.txt", 1)

Set objLogFile = objFSO.CreateTextFile("U:\GitHub\cbc_winteam\dev\opsmgr\apc-ups\MIBParsing\traps.csv", 2, True)


objLogFile.Write "Name,"
objLogFile.Write "Description,"
objLogFile.Write "Type,"
objLogFile.Write "Summary,"
objLogFile.Write "Severity,"
objLogFile.Write "State,"
objLogFile.Write "OID"
objLogFile.Writeline


do while not objFile.AtEndOfStream
    curLine =  objFile.ReadLine()
	
	'Finds Name
	if InStr(curLine, "TRAP-TYPE") <> 0 Then
			name = Replace(curLine, "TRAP-TYPE", "")
			name = Trim(name)
			name = name + ","
			objLogFile.Write name
	end if
	
	'Description
	if InStr(curLine, "DESCRIPTION") <> 0 Then
		description = ""
		curLine =  objFile.ReadLine()
		do while InStr(curLine, "--#TYPE") = 0 
				curLine = Trim(curLine)
				description = description + " " + curLine
				curLine =  objFile.ReadLine()
		loop
		description = Trim(description)
		description = description + ","
		objLogFile.Write description
	end if
	
	'Type
	if InStr(curLine, "--#TYPE") <> 0 Then
			typeTrap = Replace(curLine, "--#TYPE", "")
			typeTrap = Trim(typeTrap)
			typeTrap = typeTrap + ","
			objLogFile.Write typeTrap
	end if
	
	'Summary
	if InStr(curLine, "--#SUMMARY") <> 0 Then
			summary = Replace(curLine, "--#SUMMARY", "")
			summary = Trim(summary)
			summary = summary + ","
			objLogFile.Write summary
	end if
	
	'Severity
	if InStr(curLine, "--#SEVERITY") <> 0 Then
			severity = Replace(curLine, "--#SEVERITY ", "")
			severity = Trim(severity)
			severity = severity + ","
			objLogFile.Write severity
	end if
	'State
	if InStr(curLine, "--#STATE") <> 0 Then
			state = Replace(curLine, "--#STATE", "")
			state = Trim(state)
			state = state + ","
			objLogFile.Write state
			curLine  = objFile.ReadLine()
	end if
	
	'OID
	if InStr(curLine, "::=") <> 0 Then
			oid = Replace(curLine, "::=", "")
			oid = Trim(oid)
			oid = "1.3.6.1.4.1.318.0." + oid
			objLogFile.Write oid
			objLogFile.Writeline
	end if
loop