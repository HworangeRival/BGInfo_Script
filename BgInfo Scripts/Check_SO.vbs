Option Explicit

Function CheckOperatingSystemStatus()
    Dim objWMIService, colOperatingSystems, objOS
    Dim strComputer, strStatus, strResult
    
    strComputer = "."
    strStatus = "GOOD"
    strResult = ""
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colOperatingSystems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    
    For Each objOS in colOperatingSystems
        If objOS.Status <> "OK" Then
            strStatus = "BAD"
            strResult = strResult & "System " & objOS.Caption & ": BAD (" & objOS.Status & "); "
        Else
            strResult = strResult & "System " & objOS.Caption & ": GOOD; "
        End If
    Next
    
    If strStatus = "GOOD" Then
        CheckOperatingSystemStatus = "System: Good"
    Else
        CheckOperatingSystemStatus = Left(strResult, Len(strResult) - 2) ' Rimuove l'ultimo "; "
    End If
End Function

Echo CheckOperatingSystemStatus()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
