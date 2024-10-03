' BGInfo Crypto Price Checker
' Questo script recupera i prezzi in tempo reale di multiple criptovalute

Option Explicit

' Funzione per effettuare richieste HTTP
Function HttpGet(url)
    Dim http
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send
    If http.Status = 200 Then
        HttpGet = http.ResponseText
    Else
        HttpGet = "Errore: " & http.Status
    End If
End Function

' Funzione per ottenere il prezzo di una criptovaluta
Function GetCryptoPrice(symbol)
    Dim url, response, price
    url = "https://api.coingecko.com/api/v3/simple/price?ids=" & symbol & "&vs_currencies=usd,eur"
    response = HttpGet(url)
    
    If InStr(response, """usd""") > 0 And InStr(response, """eur""") > 0 Then
        Dim usdPrice, eurPrice
        usdPrice = Split(Split(response, """usd"":")(1), ",")(0)
        eurPrice = Split(Split(response, """eur"":")(1), "}")(0)
        GetCryptoPrice = UCase(symbol) & ": $ " & usdPrice & " | " & ChrW(8364) & " " & eurPrice
    Else
        GetCryptoPrice = "Errore nel recupero del prezzo di " & symbol
    End If
End Function

' Lista delle criptovalute da monitorare
Dim cryptoList
cryptoList = Array("bitcoin", "ethereum", "ripple", "cardano", "dogecoin")

' Output dei prezzi delle criptovalute
Dim crypto
For Each crypto In cryptoList
    Echo GetCryptoPrice(crypto)
Next

'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
