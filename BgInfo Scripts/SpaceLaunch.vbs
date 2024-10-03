Option Explicit

Function GetUpcomingSpaceLaunch()
    Dim objHTTP, strURL, strResponse
    Dim jsonObject, launchInfo, launchDate
    
    strURL = "https://ll.thespacedevs.com/2.2.0/launch/?limit=1&offset=0"
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    On Error Resume Next
    objHTTP.open "GET", strURL, False
    objHTTP.setRequestHeader "User-Agent", "BGInfo Script (Educational Purpose)"
    objHTTP.send
    
    If Err.Number <> 0 Then
        GetUpcomingSpaceLaunch = "Errore nel recupero delle informazioni sul lancio. Controlla la connessione internet."
        Exit Function
    End If
    On Error Goto 0
    
    If objHTTP.Status <> 200 Then
        GetUpcomingSpaceLaunch = "Errore nel recupero dei dati. Codice di stato: " & objHTTP.Status
        Exit Function
    End If
    
    strResponse = objHTTP.responseText
    Echo strResponse
    Set jsonObject = ParseJson(strResponse)
    
    If Not jsonObject.Exists("results") Or jsonObject("results").Count = 0 Then
        GetUpcomingSpaceLaunch = "Nessun lancio trovato nei risultati."
        Exit Function
    End If
    
    Set launchInfo = jsonObject("results")(0)
    launchDate = SafeGetValue(launchInfo, "net")
    If launchDate <> "" Then
        launchDate = ConvertToValidDate(launchDate)
    Else
        launchDate = "Data non disponibile"
    End If
    
    GetUpcomingSpaceLaunch = "Prossimo Lancio Spaziale:" & vbNewLine & _
                             "Missione: " & SafeGetValue(launchInfo, "name") & vbNewLine & _
                             "Data: " & launchDate & vbNewLine & _
                             "Agenzia: " & SafeGetNestedValue(launchInfo, "launch_service_provider,name") & vbNewLine & _
                             "Razzo: " & SafeGetNestedValue(launchInfo, "rocket,configuration,name") & vbNewLine & _
                             "Luogo: " & SafeGetNestedValue(launchInfo, "pad,name") & ", " & SafeGetNestedValue(launchInfo, "pad,location,name")
End Function

Function ConvertToValidDate(dateString)
    Dim convertedDate
    On Error Resume Next
    convertedDate = CDate(dateString)
    If Err.Number = 0 Then
        ConvertToValidDate = FormatDateTime(convertedDate, vbShortDate) & " " & FormatDateTime(convertedDate, vbShortTime) & " UTC"
    Else
        ConvertToValidDate = dateString
    End If
    On Error Goto 0
End Function

Function ParseJson(jsonString)
    Set ParseJson = CreateObject("Scripting.Dictionary")
    
    ' Rimuovi spazi bianchi non necessari e parentesi esterne
    jsonString = Trim(jsonString)
    jsonString = Mid(jsonString, 2, Len(jsonString) - 2)
    
    Dim pairs, pair, key, value
    pairs = Split(jsonString, ",")
    
    For Each pair In pairs
        pair = Trim(pair)
        key = Left(pair, InStr(pair, ":") - 1)
        value = Mid(pair, InStr(pair, ":") + 1)
        
        ' Rimuovi le virgolette dalle chiavi e dai valori
        key = Replace(Trim(key), """", "")
        value = Trim(value)
        
        If Left(value, 1) = "{" Then
            ' Oggetto nidificato
            Set value = ParseJson(value)
        ElseIf Left(value, 1) = "[" Then
            ' Array
            Set value = ParseJsonArray(value)
        Else
            ' Valore semplice
            value = Replace(value, """", "")
        End If
        
        If Not ParseJson.Exists(key) Then
            ParseJson.Add key, value
        End If
    Next
End Function

Function ParseJsonArray(jsonArray)
    Set ParseJsonArray = CreateObject("Scripting.Dictionary")
    
    ' Rimuovi le parentesi quadre
    jsonArray = Mid(jsonArray, 2, Len(jsonArray) - 2)
    
    Dim items, item, index
    items = Split(jsonArray, ",")
    index = 0
    
    For Each item In items
        item = Trim(item)
        If Left(item, 1) = "{" Then
            Set item = ParseJson(item)
        Else
            item = Replace(item, """", "")
        End If
        ParseJsonArray.Add index, item
        index = index + 1
    Next
End Function

Function SafeGetValue(obj, key)
    If obj Is Nothing Then
        SafeGetValue = ""
    ElseIf TypeName(obj) = "Dictionary" Then
        If obj.Exists(key) Then
            If IsObject(obj(key)) Then
                Set SafeGetValue = obj(key)
            Else
                SafeGetValue = obj(key)
            End If
        Else
            SafeGetValue = ""
        End If
    Else
        SafeGetValue = ""
    End If
End Function

Function SafeGetNestedValue(obj, keys)
    Dim arrKeys, i, value
    arrKeys = Split(keys, ",")
    Set value = obj
    For i = 0 To UBound(arrKeys)
        If TypeName(value) = "Dictionary" Then
            If value.Exists(arrKeys(i)) Then
                If IsObject(value(arrKeys(i))) Then
                    Set value = value(arrKeys(i))
                Else
                    value = value(arrKeys(i))
                End If
            Else
                SafeGetNestedValue = ""
                Exit Function
            End If
        Else
            SafeGetNestedValue = ""
            Exit Function
        End If
    Next
    If IsObject(value) Then
        SafeGetNestedValue = ""
    Else
        SafeGetNestedValue = value
    End If
End Function

Function CacheAndGetLaunch()
    Dim fso, file, cachedLaunch, currentDate, cachedDate
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim cacheFile : cacheFile = fso.GetSpecialFolder(2) & "\space_launch_cache.txt"
    currentDate = Date()
    
    If fso.FileExists(cacheFile) Then
        Set file = fso.OpenTextFile(cacheFile, 1)
        cachedDate = file.ReadLine()
        cachedLaunch = file.ReadAll()
        file.Close()
        
        If CDate(cachedDate) = currentDate Then
            CacheAndGetLaunch = cachedLaunch
            Exit Function
        End If
    End If
    
    Dim newLaunch : newLaunch = GetUpcomingSpaceLaunch()
    
    Set file = fso.CreateTextFile(cacheFile, True)
    file.WriteLine currentDate
    file.Write newLaunch
    file.Close()
    
    CacheAndGetLaunch = newLaunch
End Function

' Esecuzione dello script
Echo CacheAndGetLaunch()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
