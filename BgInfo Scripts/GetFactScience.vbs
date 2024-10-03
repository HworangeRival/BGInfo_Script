Option Explicit

Function GetDailyScienceFact()
    Dim objHTTP, strURL, strResponse
    Dim currentDate, dayOfYear
    
    ' Ottieni il giorno dell'anno corrente (1-366)
    currentDate = Now()
    dayOfYear = DateDiff("d", DateSerial(Year(currentDate), 1, 0), currentDate)
    
    ' URL dell'API di Numbers API per un fatto sul numero del giorno dell'anno
    strURL = "http://numbersapi.com/" & dayOfYear & "/math"
    
    ' Crea l'oggetto per la richiesta HTTP
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    On Error Resume Next
    objHTTP.open "GET", strURL, False
    objHTTP.send
    
    If Err.Number <> 0 Then
        GetDailyScienceFact = "Error check internet."
        Exit Function
    End If
    On Error Goto 0
    
    ' Ottieni la risposta
    strResponse = objHTTP.responseText
    
    ' Formatta la risposta
    GetDailyScienceFact = dayOfYear & "' day of Year" & vbNewLine & vbTab & _
                          strResponse
End Function

Function CacheAndGetFact()
    Dim fso, file, cachedFact, currentDate, cachedDate
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Percorso del file di cache
    Dim cacheFile : cacheFile = fso.GetSpecialFolder(2) & "\daily_science_fact_cache.txt"
    
    currentDate = Date()
    
    ' Controlla se il file di cache esiste e se è aggiornato
    If fso.FileExists(cacheFile) Then
        Set file = fso.OpenTextFile(cacheFile, 1)
        cachedDate = file.ReadLine()
        cachedFact = file.ReadAll()
        file.Close()
        
        ' Se la data nel cache è quella di oggi, usa il fatto in cache
        If CDate(cachedDate) = currentDate Then
            CacheAndGetFact = cachedFact
            Exit Function
        End If
    End If
    
    ' Se non c'è un cache valido, ottieni un nuovo fatto
    Dim newFact : newFact = GetDailyScienceFact()
    
    ' Salva il nuovo fatto nel cache
    Set file = fso.CreateTextFile(cacheFile, True)
    file.WriteLine currentDate
    file.Write newFact
    file.Close()
    
    CacheAndGetFact = newFact
End Function

Echo CacheAndGetFact()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
