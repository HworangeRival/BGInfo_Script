Option Explicit

Function CheckProcessorStatus()
    Dim objWMIService, colProcessors, objProcessor
    Dim strComputer, strStatus, strResult
    
    strComputer = "."
    strStatus = "GOOD"
    strResult = ""
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colProcessors = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
    
    For Each objProcessor in colProcessors
        If objProcessor.Status <> "OK" Then
            strStatus = "BAD"
            strResult = strResult & "Processor " & objProcessor.DeviceID & ": BAD (" & objProcessor.Status & "); "
        Else
            strResult = strResult & "Processor " & objProcessor.DeviceID & ": GOOD; "
        End If
    Next
    
    If strStatus = "GOOD" Then
        CheckProcessorStatus = "All processors: Good"
    Else
        CheckProcessorStatus = Left(strResult, Len(strResult) - 2) ' Removes the last "; "
    End If
End Function

Echo CheckProcessorStatus()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
