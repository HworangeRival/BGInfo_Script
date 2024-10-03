' BGInfo Script: Ricetta del Giorno usando TheMealDB API
Option Explicit

' Funzione per ottenere il contenuto da un URL
Function HTTPGet(strURL)
    Dim objHTTP
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", strURL, False
    objHTTP.Send
    HTTPGet = objHTTP.ResponseText
End Function

' Funzione per estrarre un valore da una stringa JSON
Function ExtractJSONValue(jsonString, key)
    Dim regex, matches
    Set regex = New RegExp
    regex.Pattern = """" & key & """:""([^""]+)"""
    regex.Global = False
    Set matches = regex.Execute(jsonString)
    If matches.Count > 0 Then
        ExtractJSONValue = matches(0).SubMatches(0)
    Else
        ExtractJSONValue = ""
    End If
End Function

' Ottieni una ricetta casuale dall'API
Dim apiUrl, jsonResponse
apiUrl = "https://www.themealdb.com/api/json/v1/1/random.php"
jsonResponse = HTTPGet(apiUrl)

' Estrai le informazioni rilevanti
Dim nomePiatto, categoriaPiatto, areaPiatto, linkIstruzioni
nomePiatto = ExtractJSONValue(jsonResponse, "strMeal")
categoriaPiatto = ExtractJSONValue(jsonResponse, "strCategory")
areaPiatto = ExtractJSONValue(jsonResponse, "strArea")
linkIstruzioni = ExtractJSONValue(jsonResponse, "strYoutube")
linkIstruzioni = Replace(linkIstruzioni, "\", "")
' Formatta l'output per BGInfo
'Echo "Ricetta del Giorno:"
Echo nomePiatto & vbNewLine & vbTab & " Video: " & linkIstruzioni

'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
