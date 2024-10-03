Option Explicit

Function CheckVideoControllerStatus()
    Dim objWMIService, colVideoControllers, objVideoController
    Dim strComputer, strStatus, strResult
    
    strComputer = "."
    strStatus = "GOOD"
    strResult = ""
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colVideoControllers = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
    
    For Each objVideoController in colVideoControllers
        If objVideoController.Status <> "OK" Then
            strStatus = "BAD"
            strResult = strResult &   objVideoController.Name & ": BAD (" & objVideoController.Status & "); "
        Else
            strResult = strResult &   objVideoController.Name & ": GOOD; "
        End If
    Next
    
    If strStatus = "GOOD" Then
        CheckVideoControllerStatus = "All Card: Good"
    Else
        CheckVideoControllerStatus = Left(strResult, Len(strResult) - 2) ' Rimuove l'ultimo "; "
    End If
End Function

Echo CheckVideoControllerStatus()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
