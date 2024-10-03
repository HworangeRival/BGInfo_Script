Option Explicit

Const API_KEY = "f8ddaa7099fceed72f8fb08a5d9a9aa5" 'APIKEY
Const CITY = "Rotterdam" 'CITY

Function GetWeather()
    Dim objHTTP, jsonText, weatherData
    
    Dim strURL
    strURL = "https://api.openweathermap.org/data/2.5/weather?q=" & CITY & "&appid=" & API_KEY & "&units=metric"
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", strURL, False
    objHTTP.Send
    
    If objHTTP.Status = 200 Then
        jsonText = objHTTP.responseText
        
        Dim temp, weatherMain
        temp = ExtractValue(jsonText, """temp"":")
        weatherMain = ExtractValue(jsonText, """main"":")
        
        weatherData = CITY & ", " & weatherMain & ", " & temp & "'C"
        
        weatherData = Replace(weatherData, "[Â]", "")
        GetWeather = weatherData
    Else
        GetWeather = "Weather Unavailable"
    End If
End Function

Function ExtractValue(jsonText, key)
    Dim startPos, endPos
    startPos = InStr(jsonText, key)
    If startPos > 0 Then
        startPos = startPos + Len(key)
        If Left(Mid(jsonText, startPos, 1), 1) = """" Then
            startPos = startPos + 1
            endPos = InStr(startPos, jsonText, """")
        Else
            endPos = InStr(startPos, jsonText, ",")
            If endPos = 0 Then endPos = InStr(startPos, jsonText, "}")
        End If
        If endPos > startPos Then
            ExtractValue = Mid(jsonText, startPos, endPos - startPos)
        End If
    End If
End Function

Echo GetWeather()

'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
