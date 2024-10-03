' Script per visualizzare l'ultima vulnerabilità da Circl.lu
' ATTENZIONE: Questo script accede a dati reali su vulnerabilità.
' Usare con cautela e solo per scopi legittimi di ricerca sulla sicurezza.

Option Explicit

' Funzione per ottenere dati dall'API Circl.lu
Function GetCirclData()
    On Error Resume Next
    Dim xhr
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        GetCirclData = "Error: " & Err.Description
        Exit Function
    End If
    
    xhr.Open "GET", "https://cve.circl.lu/api/last", False
    xhr.Send
    If Err.Number <> 0 Then
        GetCirclData = "API Access Error" '& Err.Description
    ElseIf xhr.Status = 200 Then
        GetCirclData = xhr.ResponseText
    Else
        GetCirclData = "Error: State " & xhr.Status
    End If
    On Error GoTo 0
End Function

' Funzione per estrarre informazioni JSON
Function ExtractJsonValue(jsonString, key)
    Dim regex, matches
    Set regex = New RegExp
    regex.Pattern = """" & key & """\s*:\s*""?([^"",\}]+)""?"
    regex.Global = False
    Set matches = regex.Execute(jsonString)
    If matches.Count > 0 Then
        ExtractJsonValue = matches(0).SubMatches(0)
    Else
        ExtractJsonValue = "Not found"
    End If
End Function

' Funzione per estrarre array JSON
Function ExtractJsonArray(jsonString, key)
    Dim regex, matches
    Set regex = New RegExp
    regex.Pattern = """" & key & """\s*:\s*\[(.*?)\]"
    regex.Global = False
    Set matches = regex.Execute(jsonString)
    If matches.Count > 0 Then
        ExtractJsonArray = matches(0).SubMatches(0)
    Else
        ExtractJsonArray = "Not found"
    End If
End Function

' Ottieni i dati dall'API
Dim rawData
rawData = GetCirclData()

' Estrai le informazioni rilevanti
Dim cveId, summary, vulnName, lastModified, references
If InStr(rawData, "Error") = 0 Then
    cveId = ExtractJsonValue(rawData, "id")
    summary = ExtractJsonValue(rawData, "summary")
    lastModified = ExtractJsonValue(rawData, "Modified")
    references = ExtractJsonArray(rawData, "references")
    
    ' Estrai il nome della vulnerabilità dal sommario
    Dim nameRegex, nameMatches
    Set nameRegex = New RegExp
    nameRegex.Pattern = """([^""]+)""(\s+vulnerability|\s+Vulnerability)"
    nameRegex.Global = False
    Set nameMatches = nameRegex.Execute(summary)
    If nameMatches.Count > 0 Then
        vulnName = nameMatches(0).SubMatches(0)
    Else
        vulnName = "Unkonown"
    End If
Else
    cveId = "N/A"
    summary = rawData
    vulnName = "N/A"
    lastModified = "N/A"
    references = "N/A"
End If

' Prepara l'informazione per BGInfo
Dim vulnerabilityInfo
vulnerabilityInfo = "ID: " & cveId &  " Name: " & vulnName & vbNewLine & vbTab & _
                    "Last Edit: " & lastModified & vbNewLine & vbTab & _
                    "Description: " & FormatLongText(summary, 70) & vbNewLine & vbTab & _
                    "References: " & FormatLongText(references, 70)

Function FormatLongText(text, lineLength)
    Dim formattedText, i, currentLine
    formattedText = ""
    currentLine = ""
    
    For i = 1 To Len(text)
        currentLine = currentLine & Mid(text, i, 1)
        If Len(currentLine) >= lineLength Then
            formattedText = formattedText & currentLine & vbNewLine & vbTab
            currentLine = ""
        End If
    Next
    
    If Len(currentLine) > 0 Then
        formattedText = formattedText & currentLine
    End If
    
    FormatLongText = formattedText
End Function

' Visualizza con BGInfo o con WScript.Echo
On Error Resume Next
BGInfo.AddCustomField "VulnerabilityInfo", vulnerabilityInfo
If Err.Number <> 0 Then
   ' Echo "Errore nell'aggiungere il campo a BGInfo: " & Err.Description
   ' Echo "Risultato:" & vbNewLine & vulnerabilityInfo
End If
On Error GoTo 0
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'


' Output per debug
Echo vulnerabilityInfo