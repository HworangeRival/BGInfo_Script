Option Explicit

Function GetRandomPhilosophicalQuote()
    Dim arrQuotes
    arrQuotes = Array( _
        "Conosci te stesso. - Socrate", _
        "Penso, dunque sono. - Cartesio", _
        "L'uomo e' condannato ad essere libero. - Sartre", _
        "La vita non esaminata non vale la pena di essere vissuta. - Socrate", _
        "Diventa cio' che sei. - Nietzsche", _
        "La bellezza salvera' il mondo. - Dostoevskij", _
        "L'uomo e' misura di tutte le cose. - Protagora", _
        "Sapere e' potere. - Francesco Bacone", _
        "Il dubbio e' l'inizio della saggezza. - Aristotele", _
        "La liberta' e' l'ossigeno dell'anima. - Moshe Dayan", _
        "Tutto scorre. - Eraclito", _
        "L'arte di vivere e' più simile alla lotta che alla danza. - Marco Aurelio", _
        "Chi lotta puo' perdere, chi non lotta ha gia' perso. - Bertolt Brecht", _
        "La vita e' cio' che ti accade mentre sei occupato a fare altri piani. - John Lennon", _
        "Il segreto della felicita' e' la liberta', il segreto della liberta' e' il coraggio. - Tucidide", _
        "Esse est percipi. (Essere e' essere percepito) - Berkeley", _
        "Cogito ergo sum. (Penso quindi sono) - Cartesio", _
        "La bellezza e' negli occhi di chi guarda. - Hume", _
        "L'uomo e' un lupo per l'uomo. - Hobbes", _
        "Dio e' morto. - Nietzsche", _
        "So di non sapere. - Socrate", _
        "Il tempo e' denaro. - Benjamin Franklin", _
        "La verita' vi rendera' liberi. - Gesù di Nazareth", _
        "L'inferno sono gli altri. - Sartre", _
        "Veni, vidi, vici. (Venni, vidi, vinsi) - Giulio Cesare", _
        "Memento mori. (Ricordati che devi morire) - Antica Roma", _
        "Carpe diem. (Cogli l'attimo) - Orazio", _
        "Panta rei. (Tutto scorre) - Eraclito", _
        "Homo homini lupus. (L'uomo e' lupo per l'uomo) - Plauto", _
        "Nulla dies sine linea. (Nessun giorno senza una linea) - Plinio il Vecchio", _
        "Chi sa di non sapere, sa. - Confucio", _
        "L'essenziale e' invisibile agli occhi. - Antoine de Saint-Exupery", _
        "La felicita' non e' avere cio' che si desidera, ma desiderare cio' che si ha. - Oscar Wilde", _
        "La vita e' un mistero da vivere, non un problema da risolvere. - Gandhi", _
        "Il vero viaggio di scoperta non consiste nel cercare nuove terre, ma nell'avere nuovi occhi. - Marcel Proust", _
        "La pazienza e' amara, ma il suo frutto e' dolce. - Jean-Jacques Rousseau", _
        "Il silenzio e' l'elemento in cui si formano tutte le grandi cose. - Thomas Carlyle", _
        "La saggezza inizia nella meraviglia. - Socrate", _
        "L'unica vera saggezza e' sapere di non sapere nulla. - Socrate", _
        "La vita senza ricerca non e' degna di essere vissuta. - Socrate" _
    )
    
    Randomize
    GetRandomPhilosophicalQuote = arrQuotes(Int(Rnd * UBound(arrQuotes)))
End Function

Echo GetRandomPhilosophicalQuote()
'
'                     :::    ::: :::       :::  ::::::::  :::::::::      :::     ::::    :::  ::::::::  ::::::::::      :::::::::  ::::::::::: :::     :::     :::     :::                             
'        :+: :+:      :+:    :+: :+:       :+: :+:    :+: :+:    :+:   :+: :+:   :+:+:   :+: :+:    :+: :+:             :+:    :+:     :+:     :+:     :+:   :+: :+:   :+:             :+: :+:         
'                     +:+    +:+ +:+       +:+ +:+    +:+ +:+    +:+  +:+   +:+  :+:+:+  +:+ +:+        +:+             +:+    +:+     +:+     +:+     +:+  +:+   +:+  +:+                             
'                     +#++:++#++ +#+  +:+  +#+ +#+    +:+ +#++:++#:  +#++:++#++: +#+ +:+ +#+ :#:        +#++:++#        +#++:++#:      +#+     +#+     +:+ +#++:++#++: +#+                             
'                     +#+    +#+ +#+ +#+#+ +#+ +#+    +#+ +#+    +#+ +#+     +#+ +#+  +#+#+# +#+   +#+# +#+             +#+    +#+     +#+      +#+   +#+  +#+     +#+ +#+                             
'#+# #+# #+# #+#      #+#    #+#  #+#+# #+#+#  #+#    #+# #+#    #+# #+#     #+# #+#   #+#+# #+#    #+# #+#             #+#    #+#     #+#       #+#+#+#   #+#     #+# #+#             #+# #+# #+# #+# 
'### ###              ###    ###   ###   ###    ########  ###    ### ###     ### ###    ####  ########  ##########      ###    ### ###########     ###     ###     ### ##########              ### ### 
'
