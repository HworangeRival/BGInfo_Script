Option Explicit

Function CheckNetworkAdapterStatus()
    Dim objWMIService, colNetAdapters, objNetAdapter
    Dim strComputer, strStatus, strResult
    
    strComputer = "."
    strStatus = "GOOD"
    strResult = ""
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colNetAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE PhysicalAdapter = True")
    
    For Each objNetAdapter in colNetAdapters
        If objNetAdapter.NetConnectionStatus <> 2 Then ' 2 significa "Connected"
            strStatus = "BAD"
            strResult = strResult &  objNetAdapter.Name & ": BAD (Not Connected); "
        Else
            strResult = strResult &  objNetAdapter.Name & ": GOOD; "
        End If
    Next
    
    If strStatus = "GOOD" Then
        CheckNetworkAdapterStatus = "All adapters: Good"
    Else
        CheckNetworkAdapterStatus = Left(strResult, Len(strResult) - 2) ' Rimuove l'ultimo "; "
    End If
End Function

Echo CheckNetworkAdapterStatus()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
