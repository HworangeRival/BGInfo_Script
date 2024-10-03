Option Explicit

Function GetWitchesCalendarDay()
    Dim currentDate, yearz, monthx, dayx
    Dim sabbat, nextSabbat, daysUntilNext
    
    currentDate = Now()
    yearz = CInt(Year(currentDate))
    monthx = Month(currentDate)
    dayx = Day(currentDate)
    
    sabbat = ""
    nextSabbat = ""
    daysUntilNext = 0
    
    Select Case True
        Case (monthx = 12 And dayx >= 21) Or (monthx = 1 And dayx <= 1)
            sabbat = "Yule (Solstizio d'Inverno)"
            nextSabbat = "Imbolc"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz + 1, 2, 1))
        Case (monthx = 1 And dayx > 1) Or (monthx = 2 And dayx < 2)
            sabbat = "Periodo tra Yule e Imbolc"
            nextSabbat = "Imbolc"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 2, 1))
        Case monthx = 2 And dayx >= 2 And dayx <= 7
            sabbat = "Imbolc"
            nextSabbat = "Ostara"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 3, 21))
        Case (monthx = 2 And dayx > 7) Or (monthx = 3 And dayx < 21)
            sabbat = "Periodo tra Imbolc e Ostara"
            nextSabbat = "Ostara"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 3, 21))
        Case monthx = 3 And dayx >= 21 And dayx <= 25
            sabbat = "Ostara (Equinozio di Primavera)"
            nextSabbat = "Beltane"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 5, 1))
        Case (monthx = 3 And dayx > 25) Or monthx = 4
            sabbat = "Periodo tra Ostara e Beltane"
            nextSabbat = "Beltane"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 5, 1))
        Case monthx = 5 And dayx >= 1 And dayx <= 5
            sabbat = "Beltane"
            nextSabbat = "Litha"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 6, 21))
        Case (monthx = 5 And dayx > 5) Or (monthx = 6 And dayx < 21)
            sabbat = "Periodo tra Beltane e Litha"
            nextSabbat = "Litha"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 6, 21))
        Case monthx = 6 And dayx >= 21 And dayx <= 25
            sabbat = "Litha (Solstizio d'Estate)"
            nextSabbat = "Lammas"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 8, 1))
        Case (monthx = 6 And dayx > 25) Or monthx = 7
            sabbat = "Periodo tra Litha e Lammas"
            nextSabbat = "Lammas"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 8, 1))
        Case monthx = 8 And dayx >= 1 And dayx <= 7
            sabbat = "Lammas (Lughnasadh)"
            nextSabbat = "Mabon"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 9, 21))
        Case (monthx = 8 And dayx > 7) Or (monthx = 9 And dayx < 21)
            sabbat = "Periodo tra Lammas e Mabon"
            nextSabbat = "Mabon"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 9, 21))
        Case monthx = 9 And dayx >= 21 And dayx <= 25
            sabbat = "Mabon (Equinozio d'Autunno)"
            nextSabbat = "Samhain"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 10, 31))
        Case (monthx = 9 And dayx > 25) Or monthx = 10
            sabbat = "Periodo tra Mabon e Samhain"
            nextSabbat = "Samhain"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 10, 31))
        Case monthx = 11 And dayx >= 1 And dayx <= 7
            sabbat = "Samhain"
            nextSabbat = "Yule"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 12, 21))
        Case Else
            sabbat = "Periodo tra Samhain e Yule"
            nextSabbat = "Yule"
            daysUntilNext = DateDiff("d", currentDate, DateSerial(yearz, 12, 21))
    End Select
    
    GetWitchesCalendarDay = "Today: " & sabbat & vbNewLine & vbTab & _
                            "Next Sabbat: " & nextSabbat & " (After " & daysUntilNext & " days)"
End Function

Echo GetWitchesCalendarDay()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
