Option Explicit

' Funzione per ottenere l'altezza delle onde
Function GetWaveHeight(latitude, longitude)
    Dim objHTTP, strURL, strResponse, jsonObject
    
    ' Sostituisci 'TUA_API_KEY' con la tua chiave API di Stormglass
    Const API_KEY = "TUA_API_KEY"
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    strURL = "https://api.stormglass.io/v2/weather/point?lat=" & latitude & "&lng=" & longitude & "&params=waveHeight"
    
    objHTTP.Open "GET", strURL, False
    objHTTP.setRequestHeader "Authorization", API_KEY
    objHTTP.Send
    
    If objHTTP.Status = 200 Then
        strResponse = objHTTP.responseText
        ' Parsing JSON semplificato (in una situazione reale, usare un parser JSON appropriato)
        Set jsonObject = CreateObject("ScriptControl")
        jsonObject.Language = "JScript"
        jsonObject.AddCode "var jsonObj = " & strResponse
        
        ' Ottieni il primo valore di altezza d'onda
        GetWaveHeight = Round(jsonObject.Eval("jsonObj.hours[0].waveHeight[0].value"), 2)
    Else
        GetWaveHeight = "Errore nel recupero dei dati"
    End If
    
    Set objHTTP = Nothing
End Function

' Coordinate di esempio (sostituisci con le coordinate del luogo desiderato)
Const LATITUDE = "41.9028" ' Esempio: Roma
Const LONGITUDE = "12.4964"

' Imposta il testo per BGInfo
Echo "Altezza onde a Roma: " & GetWaveHeight(LATITUDE, LONGITUDE) & " metri"

'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
