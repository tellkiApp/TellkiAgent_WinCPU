'###################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								              ##
'##																													              ##
'## December, 2014																									              ##
'##																													              ##
'## Version 1.0																										              ##
'##																													              ##
'## DESCRIPTION: Monitor CPU utilization																			              ##
'##																													              ##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "CPU.vbs" <HOST> <METRIC_STATE> <USERNAME> <PASSWORD> <DOMAIN>             ##
'##																													              ##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "CPU.vbs" "10.10.10.1" "1,1,1,0,0,0" "user" "pwd" "domain"	              ##
'##																													              ##
'## README:	<METRIC_STATE> is generated internally by Tellki and its only used by Tellki default monitors. 						  ##
'##         1 - metric is on ; 0 - metric is off					              												  ##
'## 																												              ##
'## 	    <USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this ##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   ##
'##			pass them to the script.																						      ##
'## 																												              ##
'###################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 5 Then 
	CALL ShowError(3, 0) 
End If
'Set Culture - en-us
SetLocale(1033)

'METRIC_ID
Const CPUUtilization = "222:% CPU Utilization:6"
Const CPUUserTime = "200:% User Time:6"
Const CPUPrivilegedTime = "39:% Privileged Time:6"
Const CPUInterruptTime = "105:% Interrupt Time:6"
Const CPUInterruptsSec = "208:Interrupts/Sec:4"
Const CPUQueueLength = "51:Processor Queue Length:4"


'INPUTS
Dim Host, MetricState, Username, Password, Domain
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
Username = WScript.Arguments(2)
Password = WScript.Arguments(3)
Domain = WScript.Arguments(4)


Dim arrMetrics, arrProcesses, top
arrMetrics = Split(MetricState,",")
Dim objSWbemLocator, objSWbemServices, colItems
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim Counter, objItem, FullUserName, OS
Dim arrProcess(100)
Counter = 0

	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)
	
	if Err.Number = 0 Then
		objSWbemServices.Security_.ImpersonationLevel = 3
		OS = GetOSVersion(objSWbemServices)
		if OS >= 3000 Then
			Set colItems = objSWbemServices.ExecQuery("select Name, PercentProcessorTime FROM Win32_PerfFormattedData_PerfProc_Process WHERE Name<>'_Total' And Name<>'Idle'",,16) 
			If colItems.Count <> 0 Then
				For Each objItem in colItems 
					arrProcesses = arrProcesses & objItem.Name & ":" & objItem.PercentProcessorTime + "|"
				Next
				top = GetTop(arrProcesses,5)
			End If
		End If
		if Len(top) = 0 Then top = "-"
		if OS >= 3000 Then
			Set colItems = objSWbemServices.ExecQuery("select PercentProcessorTime,PercentUserTime,PercentPrivilegedTime,PercentInterruptTime,InterruptsPerSec from Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'",,16) 
			If colItems.Count <> 0 Then
				For Each objItem in colItems
					'% CPU Utilization
					If arrMetrics(0)=1 Then CALL Output(CPUUtilization,FormatNumber(objItem.PercentProcessorTime),"",top)
					'% User Time
					If arrMetrics(1)=1 Then CALL Output(CPUUserTime,FormatNumber(objItem.PercentUserTime),"",top)
					'% Privileged Time
					If arrMetrics(2)=1 Then CALL Output(CPUPrivilegedTime,FormatNumber(objItem.PercentPrivilegedTime),"",top)
					'% Interrupt Time
					If arrMetrics(3)=1 Then CALL Output(CPUInterruptTime,FormatNumber(objItem.PercentInterruptTime),"",top)
					'Interrupts per second
					If arrMetrics(4)=1 Then CALL Output(CPUInterruptsSec,FormatNumber(objItem.InterruptsPerSec),"",top)
				Next
			Else
				'If there is no response in WMI query
				CALL ShowError(5, Host)
			End If
		Else
			Dim sumPCT, i, N1, D1, N2, D2, objInstance1, perf_instance2, PercentProcessorTime, cpuPCT
			sumPCT=0
			For i = 1 to 5
				Set objInstance1 = objSWbemServices.Get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
				N1 = objInstance1.PercentProcessorTime
				D1 = objInstance1.TimeStamp_Sys100NS
				WScript.Sleep(1000)
				Set perf_instance2 = objSWbemServices.get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
				N2 = perf_instance2.PercentProcessorTime
				D2 = perf_instance2.TimeStamp_Sys100NS
				PercentProcessorTime = (1 - ((N2 - N1)/(D2-D1)))*100
				sumPCT=abs(PercentProcessorTime)+sumPCT
			Next
			cpuPCT=Round((sumPCT/10),2)
			If arrMetrics(0)=1 Then CALL Output(CPUUtilization,FormatNumber(cpuPCT),"",top)
		End if
		Set colItems = objSWbemServices.ExecQuery("select ProcessorQueueLength from Win32_PerfFormattedData_PerfOS_System",,16) 
		If colItems.Count <> 0 Then
			For Each objItem in colItems 
				'ProcessorQueueLength
				If arrMetrics(5)=1 Then CALL Output(CPUQueueLength,FormatNumber(objItem.ProcessorQueueLength),"",top)
			Next
		Else
			'If there is no response in WMI query
			CALL ShowError(5, Host)
		End If
        If Err.number <> 0 Then
            CALL ShowError(5, Host)
         	Err.Clear
        End If
	End If


If Err Then 
	CALL ShowError(1, 0)
Else
	WScript.Quit(0)
End If

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg 
	WScript.Quit(ErrorCode)
End Sub

Sub Output(MetricID, MetricValue, MetricObject, MetricData)
	if MetricData = "" Then 
		MetricData = "-"
	End If
	If MetricObject <> "" Then
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|"  & MetricValue & "|" & MetricObject & "=" & MetricData & "|"
		Else
			CALL ShowError(5, Host) 
		End If
	Else
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "=" & MetricData & "|"
		Else
			CALL ShowError(5, Host)
		End If
	End If
End Sub

Function GetOSVersion(SWbem)
	Dim colItems, objItem
	Set colItems = SWbem.ExecQuery("select BuildVersion from Win32_WMISetting",,16)
	For Each objItem in colItems
		GetOSVersion = CInt(objItem.BuildVersion)
	Next
End Function

Function GetTop(ValueList, TotalRecords)
	Dim Val, rs, Counter, out, exists
	Set rs = CreateObject("ADODB.RECORDSET")
	rs.Fields.append "Property", 200, 255
	rs.Fields.append "Value", 20, 25
	rs.CursorType = 3
	rs.Open
	exists = 0
	For Each objItem in Split(ValueList,"|") 
		If (objItem<>"") Then
			Val = Split(objItem,":")
			rs.AddNew
			rs.Fields("Property") = Val(0)
			rs.Fields("Value") = Val(1)
			rs.Update
			exists = 1
		End if
	Next
	rs.Sort = "Value DESC, Property"
	if exists = 1 and not rs.EOF then
		rs.MoveFirst
		Counter = 0
		Do Until rs.EOF OR Counter = TotalRecords
			If out = "" Then
				out = rs.Fields(0) & ":" & CLng(rs.Fields(1))
			Else
				out = out & ";" & rs.Fields(0) & ":" & CLng(rs.Fields(1))
			End If
			rs.MoveNext
			Counter = Counter + 1
		Loop
		GetTop = out
	Else
		GetTop = ""
	End If
	Set rs = Nothing
End Function


