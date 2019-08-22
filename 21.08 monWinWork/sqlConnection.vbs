Dim connect, sql, resultSet, pth, txt, oFSO
set oFSO = CreateObject("Scripting.FileSystemObject")
set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
set shell = CreateObject("wscript.shell")
Set connect = CreateObject("ADODB.Connection")
set resultSet = CreateObject("ADODB.recordset")
connect.ConnectionString = "Provider=SQLOLEDB;Server=172.16.42.106;Database=labmon;Trusted_Connection=True;User ID=sa;Password=QAZ123wsx"
connect.Open

dim objWMIService,totalPhysicalMemory,usedPhysicalMemory,colComputer,count,sumOfCPUInfo,CPUInfo,shell,host, exec, results, ip, collection, result, textStream, output, confFile, index,strComputer, user,password
function monWinAv(host)
	Set exec = shell.Exec("ping -n 2 -w 1000 " & host)
    results = LCase(exec.StdOut.ReadAll)
    
	if InStr(results, "ping request could not find") > 0 then
	   	monWinAv = "Error"
	elseif InStr(results, "received = 2") > 0 then
		monWinAv = "UP"
	elseif InStr(results, "received = 2") = 0 then
		monWinAv = "DOWN"
	end if
end function

set textStream = oFSO.OpenTextFile("C:\Users\bojid\Desktop\monWin\input\server_list", ForReading)

function monWinPerf(strComputer)
	output = ""

	set confFile =  oFSO.OpenTextFile("C:\Users\bojid\Desktop\monWin\conf\conf.properties", ForReading)
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
    output = output & "" & round(sumOfCPUInfo/count) & "% "

	Set colComputer = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48) 
	
	for each objItem in colComputer
		usedPhysicalMemory = objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory
		totalPhysicalMemory = objItem.TotalVisibleMemorySize
		output = output & FormatPercent(usedPhysicalMemory/totalPhysicalMemory,0)
		exit for
	Next

	monWinPerf = output
end function

do until textStream.AtEndOfStream
	input = textStream.ReadLine
	collection = split(input, ",")

	host = collection(0)
	domain = collection(1)

	result = monWinAv(host) 'result = UP or DOWN, maybe implement some error handling
		sql = "insert into MON_AV_NT(FQDN,IP,State)" & _ 
			" values('" & domain & "','" & host & "','" & result & "')"
	set resultSet = connect.Execute(sql)

	if result = "UP" then
		result = Split(monWinPerf(host), " ") 'example monWinPerf output: "27% 66%" 
		sql = "insert into MON_PERF_NT(FQDN,IP,CPU_Usage,Memory_Usage)" & _
			" values('" & domain & "','" & host & "','" & result(0) & "','" & result(1) & "')"
		set resultSet = connect.Execute(sql)
	else
		'I still can't decide what to insert in the table PERF when the server is DOWN
	end if 
loop

sql="SELECT * FROM MON_AV_NT"

Set resultSet = connect.Execute(sql)

pth = "C:\Users\bojid\Desktop\test.csv"
Set txt = oFSO.CreateTextFile(pth, True)

On Error Resume Next
While Not resultSet.eof
  txt.WriteLine(resultSet("FQDN"))
  resultSet.MoveNext
wend

resultSet.Close
connect.Close
Set connect = Nothing