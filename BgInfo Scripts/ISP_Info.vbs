Option Explicit

Function GetISPInfo()
    Dim objHTTP, strURL
    strURL = "https://ipapi.co/json/"
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.open "GET", strURL, False
    objHTTP.send
    If objHTTP.Status = 200 Then
        GetISPInfo = Trim(objHTTP.responseText)
    Else
        GetISPInfo = "Error: " & objHTTP.Status
    End If
    Set objHTTP = Nothing
End function
    
Function EXtractInfos(strResponse)
    Dim strISP, strCountry, strRegion, strCity, strASN
    strISP = ExtractValue(strResponse, "org")
    strCountry = ExtractValue(strResponse, "country_name")
    strRegion = ExtractValue(strResponse, "region")
    strCity = ExtractValue(strResponse, "city")
    strASN = ExtractValue(strResponse, "asn")
    EXtractInfos = "ISP: " & strISP & ", City: " & strCity'& vbNewLine & _
                  ' "Country: " & strCountry & ", ASN: " & strASN & ", Region: " & strRegion & ", City: " & strCity
End Function

Function ExtractValue(jsonStr, key)
        Dim regex, matches
        Set regex = New RegExp
        regex.Pattern = """" & key & """\s*:\s*""([^""]*)""|""" & key & """\s*:\s*(\d+)"
        regex.Global = False
        Set matches = regex.Execute(jsonStr)
        If matches.Count > 0 Then
            If matches(0).SubMatches(0) <> "" Then
                ExtractValue = matches(0).SubMatches(0)
            Else
                ExtractValue = matches(0).SubMatches(1)
            End If
        Else
            ExtractValue = ""
        End If
End Function

Echo EXtractInfos(GetISPInfo)
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
