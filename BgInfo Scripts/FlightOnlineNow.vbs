Option Explicit

' Funzione per ottenere il numero di voli
Function GetFlightCount()
    Dim objHTTP, strURL, strResponse
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    strURL = "https://opensky-network.org/api/states/all"
    
    objHTTP.Open "GET", strURL, False
    objHTTP.Send
    
    If objHTTP.Status = 200 Then
        strResponse = objHTTP.responseText
        ' Analizza la risposta JSON per estrarre il conteggio dei voli
        ' Nota: questa è una semplificazione e potrebbe richiedere una parsing JSON più robusto
        GetFlightCount = Len(Split(strResponse, """"icao24"""":")(1))
    Else
        GetFlightCount = "Errore nel recupero dei dati"
    End If
    
    Set objHTTP = Nothing
End Function

' Imposta il testo per BGInfo
Echo "Flight online now: " & GetFlightCount()

'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
