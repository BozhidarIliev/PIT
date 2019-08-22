option explicit

dim items,collection,MemoryPercentage,CPUPercentage,cacheCreated,cacheIsOld,query,objRecordSet,objConnection,count,currentIPAddress,cacheTextFile, sumOfCPUInfo,cachefile, cacheFilePath, cacheFileName,outputFile, index, confFile, user, password, outputFileName, objLogFile, logFile, confFileName, confFilePath ,fileitem, mainFolderPath, inputFileName, logFilePath, logFileName, result, ipAddress, currentIPNum, colComputer, totalPhysicalMemory, usedPhysicalMemory, inputFilePath, shell, return, exec, results, oFSO, textStream, output, line, command, objWMIService, CPUInfo, objOutputFile, continued, item, MemoryInfo, objItem, strComputer, objRefresher, colItems, outputFilePath, objSWbemLocator, objSWbemServices,connection, sql, resultSet
set oFSO = CreateObject("Scripting.FileSystemObject")
set shell = CreateObject("wscript.shell")
set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set connection = CreateObject("ADODB.Connection")
set resultSet = CreateObject("ADODB.recordset")

const ForReading = 1
const ForWriting = 2
const ForAppending = 8
const FormattedOutput = 1
const UnformattedOutput = 2

mainFolderPath = oFSO.GetParentFolderName(WScript.ScriptFullName)

inputFilePath = mainFolderPath & "\input"
outputFilePath = mainFolderPath & "\output"
logFilePath = mainFolderPath & "\logs"
confFilePath = mainFolderPath & "\conf"
cacheFilePath = mainFolderPath & "\cache"

cacheCreated = false
cacheIsOld = false

function setFilesNames(path, ByRef fileName)
	if oFSO.GetFolder(path).files.count = 0 then
		if path = logFilePath then
			set objLogFile = oFSO.CreateTextfile(logFilePath & "\logs.log", ForWriting)
			objLogFile.writeline currentTime & "Creating log file"
		elseif path = outputFilePath then
			set outputFile = oFSO.CreateTextfile(outputFilePath & "\output.log", ForWriting)
			objLogFile.writeline currentTime & "Creating output file"
		elseif path = cacheFilePath then
			set cacheTextFile = oFSO.CreateTextFile(cacheFilePath & "\cache.txt", ForWriting)
			objLogFile.writeline currentTime & "Creating cache file"
			cacheIsOld = false
			cacheCreated = true
		else 
			objLogFile.WriteLine currentTime & "Missing file in " & path
			objLogFile.WriteLine currentTime & "Script terminated."
			wscript.Quit
		end if
	else
		for each fileItem in oFSO.GetFolder(path).files
			fileName = fileItem.name
			exit for
		next
		if path = cacheFilePath then
			set cacheFile = oFSO.GetFile(cacheFilePath & "\" & cacheFileName)
			if dateDiff("N", cacheFile.DateLastModified, now) > 10 then
				cacheIsOld = true
				cacheCreated = false
			else 
				cacheIsOld = false
				cacheCreated = false
			end if
		end if

	end if 

	for each fileItem in oFSO.GetFolder(path).files
		fileName = fileItem.name
		exit for
	next
end function
call setFilesNames(logFilePath, logFileName)
Set objLogFile = oFSO.OpenTextFile(logFilePath & "\" & logFileName, ForWriting)

objlogfile.writeline currentTime & "Starting script"
objlogfile.writeline currentTime & "Checking if all files exist"

call setFilesNames(inputFilepath, inputFileName)
call setFilesNames(outputFilePath, outputFileName)
call setFilesNames(confFilePath, confFileName)
call setFilesNames(cacheFilePath, cacheFileName)

objlogfile.writeline currentTime & "All files exist" 

Set objOutputFile = oFSO.OpenTextFile(outputFilePath & "\" & outputFileName, ForWriting)
set confFile =  oFSO.OpenTextFile(confFilePath & "\" & confFileName, ForReading)

For index = 0 To 2
	if index = 0 then 
		user = confFile.ReadLine
	elseif index = 1 then 
		password = confFile.ReadLine
	elseif index = 2
		connection.ConnectionString = confFile.ReadLine
	end if
Next
'connection.open

function currentTime
	currentTime = "<" & time & "> - "
end function

function InsertAv(byval FQDN, byval IP, byval State)
' if result contains "UP" - return "UP"
	sql = "insert into MON_AV_NT(FQDN,IP,State)" & _ 
			" values('" & FQDN & "','" & IP & "','" & State & "')"
	set resultSet = connect.Execute(sql)
	resultSet.Close
end function

function InsertPerf(byval FQDN, byval IP, byval result)
'you will make an array results where before that there will be splitted value of result
sql = "insert into MON_PERF_NT(FQDN,IP,CPU_Usage,Memory_Usage)" & _
			" values('" & domain & "','" & host & "','" & results(0) & "','" & results(1) & "')"
		set resultSet = connect.Execute(sql)
end function

function monWinAv(host, formatType)
	Set exec = shell.Exec("ping -n 2 -w 1000 " & host)
    results = LCase(exec.StdOut.ReadAll)

		if InStr(results, "ping request could not find") > 0 then
			monWinAv = "Ping request could not find host " & host & "."
		elseif InStr(results, "received = 2") > 0 then
			monWinAv = "server: " & host & ", status: UP"
		elseif InStr(results, "received = 2") = 0 then
			monWinAv = "server: " & host & ", status: DOWN"
		end if

		if formatType = UnformattedOutput then
			select case monWinAv
			case "server: " & host & ", status: UP"
				monWinAv = "UP"
			case "server: " & host & ", status: DOWN"
				monWinAv = "DOWN"
			case else
				monWinAv = "ERR"
			end select
		end if
end function

function monWinPerf(strComputer,formatType)
	output = ""

	CPUPercentage = 0
	MemoryPercentage = 0
	for index = 0 to 4
		on error resume next 
		Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "Root\CIMv2", user, password)
		if err.number <> 0 then 
			objLogFile.WriteLine currentTime & "server: " & strComputer & " -  Error Description: """ & err.Description & """"
		else 												
			Set CPUInfo = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor",,48) 
			sumOfCPUInfo = 0
			count = 0
			For Each Item in CPUInfo
				sumOfCPUInfo = sumOfCPUInfo + Item.PercentProcessorTime
				count = count + 1
			Next
			CPUPercentage = CPUPercentage + round(sumOfCPUInfo/count)

			Set colComputer = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
			for each objItem in colComputer
				usedPhysicalMemory = objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory
				totalPhysicalMemory = objItem.TotalVisibleMemorySize
				MemoryPercentage = MemoryPercentage + (usedPhysicalMemory/totalPhysicalMemory)*100
				exit for
			Next
		end if
	next

	if formatType = UnformattedOutput then
		monWinPerf = round(CPUPercentage/5) & "%, " & round(MemoryPercentage/5) & "%"
	elseif formatType = FormattedOutput then
		monWinPerf = "CPU: " & round(CPUPercentage/5) & "% MEM: " & round(MemoryPercentage/5) & "%"
	end if
end function

function monWinAdr(inputFilePath)
	set textStream = oFSO.OpenTextFile(inputFilePath & "\" & inputFileName, ForReading)
	objLogFile.writeLine currentTime & "Starting procedure"

	do until textStream.AtEndOfStream
		line = textStream.ReadLine
		collection = split(line, ",")

		host = collection(0)
		domain = collection(1)

		if cacheIsOld = false and cacheCreated = false then
			set cacheTextFile = oFSO.OpenTextFile(cacheFilePath & "\" & cacheFileName, ForReading)		
			objlogfile.writeline currentTime & "Using cache file"
			objlogfile.writeline currentTime & "Retrieving performance data from servers"
			do until cacheTextFile.AtEndOfStream
				item = cacheTextFile.ReadLine
				items = Split(item,"- ")
				result = monWinPerf(host)
				if instr(item, "UP") > 0 then
					objOutputFile.WriteLine currentTime & items(1) & " - " & result
				else 
					objOutputFile.writeline currentTime & items(1)
				end if
				exit do
			loop
			objlogfile.writeline currentTime &  "End of procedure"
		elseif cacheIsOld = true and cacheCreated = false then
			objlogfile.writeline currentTime & "Overriding cache"
			objlogfile.writeline currentTime & "Running PING"
			objlogfile.writeline currentTime & "Retrieving performance data from servers"

			result = monWinAv(host)
			results = InStr(result, "UP")
			
			if results > 0 then
				objOutputFile.WriteLine result & " - " & monWinPerf(host)
				cacheTextfile.writeLine currentTime & "server: " & host & ", status: UP" 	
			else
				objOutputFile.WriteLine result
				cacheTextFile.writeLine currentTime & "server: " & host & ", status: DOWN" 	
			end if
			objlogfile.writeline currentTime & "End of procedure"
		elseif cacheIsOld = false and cacheCreated = true then
			objLogFile.writeline currentTime & "Cache not found"							
			objlogfile.writeline currentTime & "Running PING"
			objLogFile.writeline currentTime & "Retrieving performance data from servers"
			objlogfile.writeline currentTime & "Writing the output"

			result = monWinAv(host)
			results = InStr(result, "UP")
			
			if results > 0 then
				objOutputFile.WriteLine result & " - " & monWinPerf(host)
				cacheTextfile.writeLine currentTime & "server: " & host & ", status: UP" 	
			else
				objOutputFile.WriteLine result
				cacheTextFile.writeLine currentTime & "server: " & host & ", status: DOWN" 	
			end if
			objlogfile.writeline currentTime & "End of procedure"
		end if


	objlogfile.writeline currentTime & "Task completed"
	loop
end function
monWinAdr(inputFilePath)
objLogFile.writeline currentTime & "Script ended"
wscript.echo "Script ended"