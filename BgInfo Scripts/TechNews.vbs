Option Explicit

Function GetImportantTechNews()
    Dim objHTTP, strURL, strResponse
    Dim i, topStoryId, maxScore, currentScore
    Dim title, score, url
    
    strURL = "https://hacker-news.firebaseio.com/v0/topstories.json?print=pretty&orderBy=""$key""&limitToFirst=10"
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    On Error Resume Next
    objHTTP.Open "GET", strURL, False
    objHTTP.Send
    If Err.Number <> 0 Then
        GetImportantTechNews = "Error retrieving news: " & Err.Description
        Exit Function
    End If
    On Error GoTo 0
    
    If objHTTP.Status <> 200 Then
        GetImportantTechNews = "Error retrieving news. HTTP status: " & objHTTP.Status
        Exit Function
    End If
    
    strResponse = objHTTP.responseText
    strResponse = Mid(strResponse, 2, Len(strResponse) - 2) ' Remove [ and ]
    
    maxScore = 0
    topStoryId = ""
    
    For Each i In Split(strResponse, ",")
        i = Trim(i)
        i = Replace(i, "[", "")
        i = Replace(i, "]", "")
        
        strURL = "https://hacker-news.firebaseio.com/v0/item/" & i & ".json"
        objHTTP.Open "GET", strURL, False
        objHTTP.Send
        
        If objHTTP.Status = 200 Then
            strResponse = objHTTP.responseText
            currentScore = CLng(ExtractValue(strResponse, "score"))
            
            If currentScore > maxScore Then
                maxScore = currentScore
                topStoryId = i
            End If
        End If
    Next
    
    If topStoryId <> "" Then
        strURL = "https://hacker-news.firebaseio.com/v0/item/" & topStoryId & ".json"
        objHTTP.Open "GET", strURL, False
        objHTTP.Send
        
        If objHTTP.Status = 200 Then
            strResponse = objHTTP.responseText
            
            title = ExtractValue(strResponse, "title")
            score = ExtractValue(strResponse, "score")
            url = ExtractValue(strResponse, "url")
            
            GetImportantTechNews =  title & " Score: " & score & vbNewLine & vbTab & _
                                    "URL: " & url
        Else
            GetImportantTechNews = "Error retrieving news details."
        End If
    Else
        GetImportantTechNews = "No stories found."
    End If
End Function

Function ExtractValue(jsonString, key)
    Dim startPos, endPos
    startPos = InStr(jsonString, """" & key & """:")
    If startPos > 0 Then
        startPos = startPos + Len(key) + 3
        endPos = InStr(startPos, jsonString, ",")
        If endPos = 0 Then endPos = InStr(startPos, jsonString, "}")
        If endPos > 0 Then
            ExtractValue = Mid(jsonString, startPos, endPos - startPos)
            ExtractValue = Replace(ExtractValue, """", "")
        End If
    End If
End Function

' Execute the script
Echo GetImportantTechNews()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
