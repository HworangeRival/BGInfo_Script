Option Explicit

Dim emailToCheck
emailToCheck = "hworangerival@gmail.com"

Function CheckEmailReputation(email)
    Dim objHTTP, url, response, jsonResponse
    url = "https://emailrep.io/" & email
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    On Error Resume Next
    objHTTP.Open "GET", url, False
    objHTTP.setRequestHeader "User-Agent", "BGInfo Email Check Script"
    objHTTP.send ""
    If Err.Number <> 0 Then
        CheckEmailReputation = "Connection Error" '& Err.Description
    ElseIf objHTTP.Status = 429 Then
        CheckEmailReputation = "Too many Request."
    ElseIf objHTTP.Status = 200 Then
        response = objHTTP.responseText
        Set jsonResponse = ParseJson(response)
      If Not jsonResponse Is Nothing Then
            Dim reputation, suspiciousScore
            reputation = jsonResponse.Item("reputation")
            suspiciousScore = CDbl(jsonResponse.Item("suspicious"))
            
            If reputation = "high" Then
                CheckEmailReputation = "HIGH! ***"
            ElseIf reputation = "medium" Then
                CheckEmailReputation = "Medium! **"
            ElseIf reputation = "low" Then
                CheckEmailReputation = "Low! *"
            Else
                CheckEmailReputation = "Unknown Score! " & suspiciousScore
            End If
        Else
            CheckEmailReputation = "Error JSON Answer"
        End If
    Else
        CheckEmailReputation = "Error Service: " & objHTTP.Status
    End If
    On Error GoTo 0
    Set objHTTP = Nothing
End Function
 
Function ParseJson(jsonString)
    Set ParseJson = CreateObject("Scripting.Dictionary")

    If InStr(jsonString, """reputation"":") > 0 Then
        ParseJson.Add "reputation", GetJsonValue(jsonString, "reputation")
    End If
    If InStr(jsonString, """suspicious"":") > 0 Then
        ParseJson.Add "suspicious", GetJsonValue(jsonString, "suspicious")
    End If
End Function

Function GetJsonValue(jsonString, key)
    Dim regex, matches
    Set regex = New RegExp
    regex.Pattern = """" & key & """\s*:\s*""?([^,}""]+)""?"
    regex.Global = False
   Set matches = regex.Execute(jsonString)
    If matches.Count > 0 Then
        GetJsonValue = matches(0).SubMatches(0)
    Else
        GetJsonValue = ""
    End If
End Function

Echo  CheckEmailReputation(emailToCheck)
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
