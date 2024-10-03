Option Explicit

Function CheckMissingDrivers()
    Dim objWMIService, colItems, objItem
    Dim strComputer, strMissingDrivers
    
    strComputer = "."
    strMissingDrivers = ""
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PNPEntity WHERE ConfigManagerErrorCode <> 0")
    
    For Each objItem in colItems
        If objItem.ConfigManagerErrorCode = 28 Then ' 28 indica un driver mancante
            strMissingDrivers = strMissingDrivers & objItem.Name & "!: Missing Driver!; "
        End If
    Next
    
    If strMissingDrivers = "" Then
        CheckMissingDrivers = "All Driver OK!"
    Else
        CheckMissingDrivers = Left(strMissingDrivers, Len(strMissingDrivers) - 2) ' Rimuove l'ultimo "; "
    End If
End Function

Echo CheckMissingDrivers()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
