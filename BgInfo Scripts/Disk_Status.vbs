Option Explicit

Function CheckDiskStatus()
    Dim objWMIService, colDisks, objDisk
    Dim strComputer, strStatus, strResult
    
    strComputer = "."
    strStatus = "GOOD"
    strResult = ""
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colDisks = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
    
    For Each objDisk in colDisks
        If objDisk.Status <> "OK" Then
            strStatus = "BAD"
            strResult = strResult & objDisk.Caption & ": BAD; "
        Else
            strResult = strResult & objDisk.Caption & ": GOOD; "
        End If
    Next
    
    If strStatus = "GOOD" Then
        CheckDiskStatus = "All disk: Good"
    Else
        CheckDiskStatus = Left(strResult, Len(strResult) - 2) ' Rimuove l'ultimo "; "
    End If
End Function

Echo CheckDiskStatus()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
