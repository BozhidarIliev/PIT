option explicit

dim query,objRecordSet,objConnection,count,currentIPAddress,cacheTextFile, sumOfCPUInfo ,j,i,cachefileOpened,cacheExists,cachefile, cacheFilePath, cacheFileName,outputFile, index, confFile, user, password, outputFileName, objLogFile, logFile, confFileName, confFilePath ,fileitem, mainFolderPath, inputFileName, logFilePath, logFileName, tryPingResult, ipAddress, currentIPNum, colComputer, totalPhysicalMemory, usedPhysicalMemory, inputFilePath, shell, return, exec, results, oFSO, textStream, output, line, command, objWMIService, CPUInfo, objOutputFile, continued, item, MemoryInfo, objItem, strComputer, objRefresher, colItems, outputFilePath, objSWbemLocator, objSWbemServices
set oFSO = CreateObject("Scripting.FileSystemObject")
set shell = CreateObject("wscript.shell")
set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

const ForReading = 1
const ForWriting = 2
const ForAppending = 8

mainFolderPath = oFSO.GetParentFolderName(WScript.ScriptFullName)

inputFilePath = mainFolderPath & "\input"
outputFilePath = mainFolderPath & "\output"
logFilePath = mainFolderPath & "\logs"
confFilePath = mainFolderPath & "\conf"
cacheFilePath = mainFolderPath & "\cache"

cacheExists = false

if oFSO.GetFolder(cacheFilePath).files.count > 0 then 
	cacheExists = true
end if 

function setFilesNames(path, ByRef fileName)
	if oFSO.GetFolder(path).files.count = 0 then
			if path = logFilePath then
				set logFile = oFSO.CreateTextfile(logFilePath & "\logs.log", True)
			elseif path = outputFilePath then
				set outputFile = oFSO.CreateTextfile(outputFilePath & "\output.log", True)
			else 
				objLogFile.WriteLine "<" & time & "> - " & "Missing file in " & path
				objLogFile.WriteLine "<" & time & "> - " & "Script terminated."
				wscript.Quit
			end if
	end if

	for each fileItem in oFSO.GetFolder(path).files
		fileName = fileItem.name
		exit for
	next

end function
call setFilesNames(logFilePath, logFileName)
Set objLogFile = oFSO.OpenTextFile(logFilePath & "\" & logFileName, ForWriting)

objlogfile.writeline "<" & time & "> - " & "Starting script"
objlogfile.writeline "<" & time & "> - " & "Checking if all files exist"

call setFilesNames(inputFilepath, inputFileName)
call setFilesNames(outputFilePath, outputFileName)
call setFilesNames(confFilePath, confFileName)
call setFilesNames(cacheFilePath, cacheFileName)

objlogfile.writeline "<" & time & "> - " & "All files exist" 


Set objOutputFile = oFSO.OpenTextFile(outputFilePath & "\" & outputFileName, ForWriting)

function monWinAv(host)
	Set exec = shell.Exec("ping -n 2 -w 1000 " & host)
    results = LCase(exec.StdOut.ReadAll)
    
	if InStr(results, "ping request could not find") > 0 then
	   	monWinAv = "Ping request could not find host " & host & "."
	elseif InStr(results, "received = 2") > 0 then
		monWinAv = "server: " & host & ", status: UP"
	elseif InStr(results, "received = 2") = 0 then
		monWinAv = "server: " & host & ", status: DOWN"
	end if
end function

function monWinPerf(strComputer)
	output = ""

	set confFile =  oFSO.OpenTextFile(confFilePath & "\" & confFileName, ForReading)
	For index = 0 To 1
		if index = 0 then 
			user = confFile.ReadLine
		elseif index = 1 then 
			password = confFile.ReadLine
			exit for
		end if
	Next

	Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "Root\CIMv2", user, password)
	on error resume next 
	if err.number <> 0 then objLogFile.WriteLine "<" & Time & "> - server: " & strComputer & " -  Error Description: """ & err.Description & """" end if
													 
	Set CPUInfo = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor",,48) 
	sumOfCPUInfo = 0
	count = 0
	For Each Item in CPUInfo
		sumOfCPUInfo = sumOfCPUInfo + Item.PercentProcessorTime
		count = count + 1
	Next
    output = output & "CPU: " & round(sumOfCPUInfo/count) & "% "

	Set colComputer = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48) 
	
	for each objItem in colComputer
		usedPhysicalMemory = objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory
		totalPhysicalMemory = objItem.TotalVisibleMemorySize
		output = output & " MEM: " & FormatPercent(usedPhysicalMemory/totalPhysicalMemory,0)
		exit for
	Next

	monWinPerf = output
end function

function monWinAdr(inputFilePath)
	set textStream = oFSO.OpenTextFile(inputFilePath & "\" & inputFileName, ForReading)
	objLogFile.writeLine "<" & time & "> - " & "Starting procedure"
	objLogFile.writeLine "<" & time & "> - " & "Checking if cache exists"

	if cacheExists then  
		objLogFile.writeline "<" & time & "> - " & "Cache found"
		set cacheTextFile = oFSO.OpenTextFile(cacheFilePath & "\" & cacheFileName, ForReading)
		set cacheFile = oFSO.GetFile(cacheFilePath & "\" & cacheFileName)
		objLogFile.writeline "<" & time & "> - " & "Checking if cache file is more than 10 minutes old"
		if dateDiff("N", cacheFile.DateLastModified, now) > 10 then
			objLogFile.writeline "<" & time & "> - " & "Cache file is more than 10 minutes old" 
			cacheExists = false
			set cacheTextFile = oFSO.OpenTextFile(cacheFilePath & "\" & cacheFileName, ForWriting)
		else 
			objLogFile.writeline "<" & time & "> - " & "Cache file is not more than 10 minutes old"
		end if
		
		if cacheExists then
			objlogfile.writeline "<" & time & "> - " & "Using cache file"
			objlogfile.writeline "<" & time & "> - " & "Retrieving performance data from servers"
			i = 0
			do until textStream.AtEndOfStream
				line = textStream.ReadLine

				for each ipAddress in Split(line, ",") 'add collection and use like this addresses(0)
					currentIPAddress = ipAddress
					exit for
				next

				do until cacheTextFile.AtEndOfStream
					item = cacheTextFile.ReadLine
						if instr(item, "UP") > 0 then
							objOutputFile.WriteLine item & " - " & monWinPerf(currentIPAddress)
						else 
							objOutputFile.writeline item
						end if
						exit do
				loop
			loop
			objlogfile.writeline "End of procedure"
		else
			objlogfile.writeline "<" & time & "> - " & "Overriding cache"
			objlogfile.writeline "<" & time & "> - " & "Running PING"
			objlogfile.writeline "<" & time & "> - " & "Retrieving performance data from servers"
			do until textStream.AtEndOfStream
				line = textStream.ReadLine
				
				for each ipAddress in Split(line, ",")
					
					tryPingResult = monWinAv(ipAddress)
					results = InStr(tryPingResult, "UP")
					
					if results > 0 then
						objOutputFile.WriteLine tryPingResult & " - " & monWinPerf(ipAddress)
						cacheTextfile.writeLine "server: " & ipAddress & ", status: UP" 	
					else
						objOutputFile.WriteLine tryPingResult
						cacheTextFile.writeLine "server: " & ipAddress & ", status: DOWN" 	
					end if
					exit for
				next
			loop

		end if
	else 'cache doesnt exist
		objLogFile.writeline "<" & time & "> - " & "Cache found"
		set cacheTextFile = oFSO.CreateTextFile(cacheFilePath & "\" & cacheFileName, ForWriting)					
		objlogfile.writeline "<" & time & "> - " & "Running PING"
		objLogFile.writeline "<" & time & "> - " & "Retrieving performance data from servers"
		objlogfile.writeline "<" & time & "> - " & "Writing the output"
		do until textStream.AtEndOfStream
			line = textStream.ReadLine
			
			for each ipAddress in Split(line, ",")
				
				tryPingResult = monWinAv(ipAddress)
				results = InStr(tryPingResult, "UP")
				
				if results > 0 then
					objOutputFile.WriteLine tryPingResult & " - " & monWinPerf(ipAddress)
					cacheTextfile.writeLine "server: " & ipAddress & ", status: UP" 	
				else
					objOutputFile.WriteLine tryPingResult
					cacheTextFile.writeLine "server: " & ipAddress & ", status: DOWN" 	
				end if
				exit for
			next
		loop
		objlogfile.writeline "<" & time & "> - " & "End of procedure"
	end if
	objlogfile.writeline "<" & time & "> - " & "Task completed"
end function

monWinAdr(inputFilePath)
objLogFile.writeline "<" & time & "> - " & "Script ended"
