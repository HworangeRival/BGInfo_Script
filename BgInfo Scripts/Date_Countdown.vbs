Option Explicit

Const TARGET_DATE = "22-11-2024 09:00:00"  
Const DATE_FORMAT = "dd/MM/yyyy"  
Const TIME_FORMAT = "HH:mm:ss"   

Function GetCountdown()
    On Error Resume Next
    Dim targetDate, currentDate, timeDiff
    Dim days, hours, minutes, seconds
    Dim formattedTarget, result
    targetDate = CDate(TARGET_DATE)
    If Err.Number <> 0 Then
        GetCountdown = "Error: Date ivalid"
        Exit Function
    End If
    currentDate = Now()
    timeDiff = targetDate - currentDate
    If timeDiff < 0 Then
        GetCountdown = "Date OUT!!!"
        Exit Function
    End If
    days = Int(timeDiff)
    hours = Hour(timeDiff)
    minutes = Minute(timeDiff)
    seconds = Second(timeDiff)
    result = FormatNumber(days, 0) & " Days, " & _
             PadZero(hours) & ":" & _
             PadZero(minutes) & ":" & _
             PadZero(seconds)
  formattedTarget = FormatDateTime(targetDate, vbLongDate) & " " & FormatDateTime(targetDate, vbLongTime)
  GetCountdown =   result 'formattedTarget & ":" & vbNewLine & result
End Function
Function PadZero(num)' Funzione per aggiungere zeri iniziali
    PadZero = Right("0" & num, 2)
End Function

' Output principale
Echo GetCountdown()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
