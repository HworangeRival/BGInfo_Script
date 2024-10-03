Option Explicit

' Costanti
Const WARNING_THRESHOLD = 80
Const CRITICAL_THRESHOLD = 90

' Funzione principale
Function GetSystemStatus
    Dim status, message
    
    status = "GOOD"
    message = "System Status:" & vbCrLf
    
    ' CPU
    message = message & EvaluateMetric("CPU", GetCPUUsage(), status)
    
    ' Memoria
    message = message & EvaluateMetric("Memory", GetMemoryUsage(), status)
    
    ' Disco C:
    message = message & EvaluateMetric("Disk C:", GetDiskUsage("C:"), status)
    
    ' Processi critici
    Dim criticalProcesses
    criticalProcesses = CheckCriticalProcesses(Array("svchost.exe", "lsass.exe", "csrss.exe"))
    If criticalProcesses <> "" Then
        status = "BAD"
        message = message & criticalProcesses
    End If
    
    ' Aggiunta dello stato finale al messaggio
    message = "OVERALL " & status & vbCrLf & message
    
    GetSystemStatus = message
End Function

' Funzioni di supporto per il monitoraggio
Function GetCPUUsage
    Dim objWMIService, colItems, objItem
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("SELECT LoadPercentage FROM Win32_Processor")
    For Each objItem in colItems
        GetCPUUsage = objItem.LoadPercentage
        Exit Function
    Next
    GetCPUUsage = 0
    On Error GoTo 0
End Function

Function GetMemoryUsage
    Dim objWMIService, colItems, objItem
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("SELECT FreePhysicalMemory, TotalVisibleMemorySize FROM Win32_OperatingSystem")
    For Each objItem in colItems
        GetMemoryUsage = 100 - (objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize * 100)
        Exit Function
    Next
    GetMemoryUsage = 0
    On Error GoTo 0
End Function

Function GetDiskUsage(driveLetter)
    Dim objWMIService, colItems, objItem
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("SELECT FreeSpace, Size FROM Win32_LogicalDisk WHERE DeviceID='" & driveLetter & ":'")
    For Each objItem in colItems
        GetDiskUsage = 100 - (objItem.FreeSpace / objItem.Size * 100)
        Exit Function
    Next
    GetDiskUsage = 0
    On Error GoTo 0
End Function

Function EvaluateMetric(metricName, currentValue, ByRef status)
    Dim message
    
    message = metricName & ": " & FormatNumber(currentValue, 2) & "% "
    
    If currentValue > CRITICAL_THRESHOLD Then
        status = "BAD"
        message = message & "CRITICAL"
    ElseIf currentValue > WARNING_THRESHOLD Then
        If status <> "BAD" Then status = "WARNING"
        message = message & "WARNING"
    Else
        message = message & "OK"
    End If
    
    EvaluateMetric = message & vbCrLf
End Function

Function CheckCriticalProcesses(processes)
    Dim process, objWMIService, colItems, message
    message = ""
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    For Each process In processes
        Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & process & "'")
        If colItems.Count = 0 Then
            message = message & "Critical process " & process & " not running" & vbCrLf
        End If
    Next
    CheckCriticalProcesses = message
    On Error GoTo 0
End Function

' Output per BGInfo
Echo GetSystemStatus()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
