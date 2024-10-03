' Hacker Journal Trimestrale Release Checker for BGInfo
' Controlla l'uscita del nuovo numero trimestrale di Hacker Journal e mostra l'ultima uscita

Option Explicit

' Funzione per ottenere le date di rilascio trimestrali
Function GetQuarterlyReleaseDates(yearx)
    GetQuarterlyReleaseDates = Array( _
        DateSerial(yearx, 3, 10), _
        DateSerial(yearx, 6, 10), _
        DateSerial(yearx, 9, 10), _
        DateSerial(yearx, 12, 10) _
    )
End Function

' Funzione per ottenere la data del prossimo numero trimestrale e dell'ultimo uscito
Function GetReleaseInfo()
    Dim dtmToday, arrReleaseDates, dtmNextRelease, dtmLastRelease
    Dim i, yearx
    
    dtmToday = Date
    yearx = Year(dtmToday)
    arrReleaseDates = GetQuarterlyReleaseDates(yearx)
    
    ' Trova la prossima data di uscita e l'ultima uscita
    For i = 0 To UBound(arrReleaseDates)
        If arrReleaseDates(i) > dtmToday Then
            dtmNextRelease = arrReleaseDates(i)
            If i > 0 Then
                dtmLastRelease = arrReleaseDates(i - 1)
            Else
                dtmLastRelease = GetQuarterlyReleaseDates(yearx - 1)(3) ' Ultimo dell'anno precedente
            End If
            Exit For
        End If
    Next
    
    ' Se tutte le date sono passate, prendi la prima dell'anno successivo e l'ultima dell'anno corrente
    If IsEmpty(dtmNextRelease) Then
        dtmNextRelease = GetQuarterlyReleaseDates(year + 1)(0)
        dtmLastRelease = arrReleaseDates(3)
    End If
    
    Dim arrResult(1)
    arrResult(0) = dtmNextRelease
    arrResult(1) = dtmLastRelease
    GetReleaseInfo = arrResult
End Function

' Funzione principale
Function GetHackerJournalInfo
    Dim arrReleaseInfo, dtmNextRelease, dtmLastRelease, dtmToday, intDaysUntilRelease
    
    arrReleaseInfo = GetReleaseInfo()
    dtmNextRelease = arrReleaseInfo(0)
    dtmLastRelease = arrReleaseInfo(1)
    dtmToday = Date
    
    intDaysUntilRelease = DateDiff("d", dtmToday, dtmNextRelease)
    
    Dim strResult
    strResult = "Last edition: " & FormatDateTime(dtmLastRelease, 1) & vbNewLine & vbTab
    
    If intDaysUntilRelease = 0 Then
        strResult = strResult & "Hacker Journal TODAY!"
    ElseIf intDaysUntilRelease = 1 Then
        strResult = strResult & "Hacker Journal TOMORROW!"
    Else
       ' strResult = strResult & "Prossimo numero in uscita il: " & FormatDateTime(dtmNextRelease, 1) & vbCrLf
        strResult = strResult & "Next edition days Left: " & intDaysUntilRelease
    End If
    
    GetHackerJournalInfo = strResult
End Function

' Chiamata alla funzione principale per BGInfo
Echo GetHackerJournalInfo
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
