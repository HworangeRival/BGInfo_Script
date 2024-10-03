Option Explicit


Dim objHTTP, jsonData, latitude, longitude
Function FetchData(url)
    On Error Resume Next
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", url, False
    objHTTP.Send
    If objHTTP.Status = 200 Then
        FetchData = objHTTP.responseText
    Else
        FetchData = "Error: Web Request Error"
    End If
    Set objHTTP = Nothing
    On Error Goto 0
End Function
jsonData = FetchData("http://api.open-notify.org/iss-now.json")
If InStr(jsonData, "iss_position") > 0 Then
    Dim latPart, latValue
    latPart = Split(jsonData, """latitude"":")
    If UBound(latPart) >= 1 Then
        latValue = Split(latPart(1), ",")
        If UBound(latValue) >= 0 Then
            latitude = Replace(latValue(0), """", "")
            latitude = Trim(latitude)
        End If
    End If
    Dim lonPart, lonValue
    lonPart = Split(jsonData, """longitude"":")
    If UBound(lonPart) >= 1 Then
        lonValue = Split(lonPart(1), "}")
        If UBound(lonValue) >= 0 Then
            longitude = Replace(lonValue(0), """", "")
            longitude = Trim(longitude)
        End If
    End If
    Dim output
    output = "Latitude: " & latitude & vbCrLf & vbtab & "Longitude: " & longitude
   Echo output
Else
   Echo "Unable to fetch ISS Data."
End If

'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
