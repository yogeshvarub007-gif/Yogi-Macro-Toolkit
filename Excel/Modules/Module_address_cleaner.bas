Attribute VB_Name = "Module_address_cleaner"
Option Explicit

Sub Karnataka_Address_Cleaner()

    Dim ws As Worksheet, auditWs As Worksheet
    Dim addrCol As String, outCol As String
    Dim addrColNum As Long, outColNum As Long
    Dim LastRow As Long, i As Long
    Dim raw As String, cleaned As String
    Dim dictBlr As Object, dictDist As Object
    Dim re As Object
    
    On Error GoTo EH
    
    Set ws = ActiveSheet
    
    addrCol = UCase(InputBox("Column with addresses (e.g. A):", "Address Column", "A"))
    If addrCol = "" Then Exit Sub
    addrColNum = ws.Range(addrCol & "1").Column
    
    outCol = UCase(InputBox("Output column for cleaned result (e.g. B):", "Output Column", "B"))
    If outCol = "" Then Exit Sub
    outColNum = ws.Range(outCol & "1").Column
    
    ' Create/Reset audit sheet
    Set auditWs = CreateAuditSheet("Karnataka_Address_Audit")
    auditWs.Range("A1:C1").Value = Array("Row", "Raw Address", "Cleaned (Locality, District)")
    
    ' Dictionaries
    Set dictBlr = CreateObject("Scripting.Dictionary")
    dictBlr.CompareMode = 1
    LoadBLRLocalities dictBlr
    
    Set dictDist = CreateObject("Scripting.Dictionary")
    dictDist.CompareMode = 1
    LoadKarnatakaDistricts dictDist
    
    ' RegExp (late bound)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    
    LastRow = ws.Cells(ws.Rows.Count, addrColNum).End(xlUp).Row
    
    For i = 2 To LastRow
        
        raw = Trim(CStr(ws.Cells(i, addrColNum).Value))
        
        If raw = "" Then
            ws.Cells(i, outColNum).Value = ""
            auditWs.Cells(i, 1).Value = i
            auditWs.Cells(i, 2).Value = ""
            auditWs.Cells(i, 3).Value = "Blank"
        Else
            cleaned = NormalizeKarnatakaAddress(raw, dictBlr, dictDist, re)
            
            ws.Cells(i, outColNum).Value = cleaned
            auditWs.Cells(i, 1).Value = i
            auditWs.Cells(i, 2).Value = raw
            auditWs.Cells(i, 3).Value = cleaned
        End If
        
    Next i
    
    auditWs.Columns("A:C").AutoFit
    MsgBox "? Karnataka Address Cleaning Completed!", vbInformation
    Exit Sub
    
EH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    
End Sub

Private Function NormalizeKarnatakaAddress( _
    ByVal raw As String, _
    ByVal dictBlr As Object, _
    ByVal dictDist As Object, _
    ByVal re As Object) As String
    
    Dim lower As String
    Dim district As String, locality As String
    Dim isBlr As Boolean
    Dim tmp As String
    
    lower = LCase(raw)
    
    ' ---------- A) Fuzzy Bangalore flag ----------
    If InStr(lower, "bang") > 0 Or _
       InStr(lower, "beng") > 0 Or _
       InStr(lower, "begal") > 0 Or _
       InStr(lower, "blr") > 0 Or _
       InStr(lower, "blore") > 0 Or _
       InStr(lower, "b lore") > 0 Then
        isBlr = True
    End If
    
    ' ---------- B) PURE BANGALORE ONLY CASE ----------
    tmp = LCase(raw)
    
    ' remove punctuation and special chars
    tmp = Replace(tmp, "-", " ")
    tmp = Replace(tmp, "_", " ")
    tmp = Replace(tmp, "/", " ")
    tmp = Replace(tmp, ",", " ")
    tmp = Replace(tmp, ".", " ")
    tmp = Replace(tmp, "#", " ")
    tmp = Replace(tmp, "&", " ")
    tmp = Replace(tmp, "(", " ")
    tmp = Replace(tmp, ")", " ")
    
    ' remove numbers (PINs, etc.)
    re.Pattern = "\b\d{1,6}\b"
    tmp = re.Replace(tmp, " ")
    
    ' remove generic words
    tmp = Replace(tmp, "karnataka", " ")
    tmp = Replace(tmp, "state", " ")
    tmp = Replace(tmp, "city", " ")
    tmp = Replace(tmp, "dist", " ")
    tmp = Replace(tmp, "district", " ")
    
    ' collapse spaces
    re.Pattern = "\s+"
    tmp = Trim(re.Replace(tmp, " "))
    
    ' If after cleaning it's ONLY a Bangalore synonym ? just return Bangalore
    If tmp = "bengaluru" Or tmp = "bangalore" Or tmp = "blr" Or _
       tmp = "blore" Or tmp = "b lore" Or tmp = "bangaluru" Or _
       tmp = "bangaore" Or tmp = "bangalur" Or tmp = "bamgalore" Then
        NormalizeKarnatakaAddress = "Bangalore"
        Exit Function
    End If
    
    ' ---------- C) District detection ----------
    district = DetectDistrict(lower, dictDist)
    If district = "Bangalore" Then isBlr = True
    
    ' ---------- D) Locality detection ----------
    ' D1: Bangalore-specific dictionary (ALWAYS CHECK FIRST - HIGHEST PRIORITY)
    locality = DetectBLRLocality(lower, dictBlr)
    
    ' D2: Fallback locality from string parts (ONLY if dictionary didn't find anything)
    If locality = "" Then
        locality = FallbackLocality(raw, district, re)
        ' Never allow Bangalore city word as "locality"
        Select Case LCase(locality)
            Case "bengaluru", "bangalore", "blr", "blore", "b lore", "bangaore", "bangalur", "bamgalore"
                locality = ""
        End Select
    End If
    
    ' ---------- E) Final decision tree ----------
    
    ' 1) Bangalore-specific formatting
    If district = "Bangalore" Or isBlr Then
        If locality <> "" Then
            NormalizeKarnatakaAddress = NormalizeAbbrev(locality) & ", Bangalore"
        Else
            NormalizeKarnatakaAddress = "Bangalore"
        End If
        Exit Function
    End If
    
    ' 2) Non-Bangalore – district + locality
    If district <> "" And locality <> "" Then
        NormalizeKarnatakaAddress = locality & ", " & district
        Exit Function
    End If
    
    ' 3) Non-Bangalore – district only
    If district <> "" Then
        NormalizeKarnatakaAddress = district
        Exit Function
    End If
    
    ' 4) Locality only ? ALWAYS Bangalore
    If locality <> "" Then
        NormalizeKarnatakaAddress = NormalizeAbbrev(locality) & ", Bangalore"
        Exit Function
    End If
    
    ' 5) Nothing solid found
    NormalizeKarnatakaAddress = "Unknown"
    
End Function

Private Function DetectDistrict(ByVal lower As String, dictDist As Object) As String
    Dim k As Variant, patterns() As String, p As Variant
    
    For Each k In dictDist.Keys
        patterns = Split(k, "|")
        For Each p In patterns
            If p <> "" Then
                If InStr(lower, p) > 0 Then
                    DetectDistrict = dictDist(k)
                    Exit Function
                End If
            End If
        Next p
    Next k
End Function

Private Sub LoadKarnatakaDistricts(dictDist As Object)

    AddDist dictDist, "bangalore|bengaluru|blr|b lore|blore|bangaore|bangalur|bamgalore", "Bangalore"

    AddDist dictDist, "mysore|mysuru", "Mysore"
    AddDist dictDist, "mangalore|mangaluru|udupi|manipal", "Mangalore"
    AddDist dictDist, "hubli|hubballi|dharwad", "Hubli-Dharwad"
    AddDist dictDist, "belagavi|belgaum", "Belagavi"
    AddDist dictDist, "ballari|bellary", "Ballari"
    AddDist dictDist, "davangere|davanagere", "Davangere"
    AddDist dictDist, "tumkur|tumakuru", "Tumkur"
    AddDist dictDist, "shimoga|shivamogga", "Shimoga"
    AddDist dictDist, "hassan", "Hassan"
    AddDist dictDist, "mandya", "Mandya"
    AddDist dictDist, "kolar", "Kolar"
    AddDist dictDist, "vijayapur|bijapur", "Vijayapur"
    AddDist dictDist, "gulbarga|kalaburagi", "Kalaburagi"
    AddDist dictDist, "raichur", "Raichur"
    AddDist dictDist, "bidar", "Bidar"
    AddDist dictDist, "karwar|uttara kannada", "Uttara Kannada"
    AddDist dictDist, "chikmagalur|chikkamagaluru", "Chikkamagaluru"
    AddDist dictDist, "chitradurga", "Chitradurga"
    AddDist dictDist, "haveri", "Haveri"
    AddDist dictDist, "gadag", "Gadag"
    AddDist dictDist, "yadgir|yadagiri", "Yadgir"
    AddDist dictDist, "kopal|koppal", "Koppal"
    AddDist dictDist, "bagalkot", "Bagalkot"
    AddDist dictDist, "ramanagara|ramanagaram", "Ramanagara"
    AddDist dictDist, "chamarajanagar|chamrajnagar", "Chamarajanagar"
    AddDist dictDist, "kodagu|coorg", "Kodagu"
    
End Sub

Private Sub AddDist(dict As Object, k As String, v As String)
    If Not dict.Exists(k) Then dict.Add k, v
End Sub

Private Sub LoadBLRLocalities(dictBlr As Object)

' Central Core
AddBLR dictBlr, "majestic|kg bus stand|kempegowda", "Majestic"
AddBLR dictBlr, "gandhinagar", "Gandhinagar"
AddBLR dictBlr, "seshadripuram|seshadri puram|sheshadripuram", "Seshadripuram"
AddBLR dictBlr, "vijayanagar|vijaynagar", "Vijayanagar"
AddBLR dictBlr, "attiguppe", "Attiguppe"
AddBLR dictBlr, "chandra layout", "Chandra Layout"
AddBLR dictBlr, "magadi road", "Magadi Road"
AddBLR dictBlr, "basaveshwar nagar|basaveshwaranagar|bvn", "Basaveshwaranagar"

' Outer West
AddBLR dictBlr, "rajajinagar|rajaji nagar", "Rajajinagar"
AddBLR dictBlr, "mahalakshmi layout", "Mahalakshmi Layout"
AddBLR dictBlr, "malleshwaram|malleswaram|malleswar", "Malleshwaram"
AddBLR dictBlr, "yeshwanthpur|yeshpur|ysrpur", "Yeshwanthpur"
AddBLR dictBlr, "peenya", "Peenya"
AddBLR dictBlr, "rt nagar|r.t nagar", "RT Nagar"

' Far West / Southwest (Kengeri - RR Nagar Belt)
AddBLR dictBlr, "kengeri|kengari", "Kengeri"
AddBLR dictBlr, "rajrajeshwari nagar|rajarajeshwarinagar|rr nagar|r r nagar|rajrajeshwari|rajarajeshwari nagar", "Rajarajeshwari Nagar"
AddBLR dictBlr, "kengeri satellite town|kengeri st", "Kengeri Satellite Town"
AddBLR dictBlr, "mysore road|mysor road", "Mysore Road"
AddBLR dictBlr, "nagarbhavi|nagarabhavi", "Nagarbhavi"
AddBLR dictBlr, "jnana bharathi|jnanabharathi|jnana bharati", "Jnana Bharathi"
AddBLR dictBlr, "ullal|ullal upanagara|ullal main road", "Ullal"
AddBLR dictBlr, "smv layout|smv|shivamoga layout", "SMV Layout"
AddBLR dictBlr, "kenchenahalli|kenchanahalli", "Kenchenahalli"
AddBLR dictBlr, "global village|gv tech park", "Global Village"
AddBLR dictBlr, "bhel layout", "BHEL Layout"
AddBLR dictBlr, "pattanagere", "Pattanagere"
AddBLR dictBlr, "dubasi palya|dubasipalya", "Dubasipalya"
AddBLR dictBlr, "kothanur dinne", "Kothanur Dinne"
AddBLR dictBlr, "mailasandra", "Mailasandra"
AddBLR dictBlr, "sunkalpalya", "Sunkalpalya"
AddBLR dictBlr, "hemmegepura|hemigepura", "Hemmigepura"
AddBLR dictBlr, "kommaghatta|komaghatta", "Kommaghatta"
AddBLR dictBlr, "kengeri upanagara", "Kengeri Upanagara"
AddBLR dictBlr, "bengaluru university|bangalore university", "Bangalore University"
AddBLR dictBlr, "channasandra|chennasandra", "Channasandra"
AddBLR dictBlr, "elachenahalli", "Elachenahalli"
AddBLR dictBlr, "soundarya layout", "Soundarya Layout"

' South Bangalore
AddBLR dictBlr, "jayanagar|jaynagar", "Jayanagar"
AddBLR dictBlr, "j p nagar|jp nagar|jpnagar", "JP Nagar"
AddBLR dictBlr, "banashankari|bsk", "Banashankari"
AddBLR dictBlr, "bsk 3rd stage|bsk iii stage|bsk third stage|banashankari 3rd stage|banashankari third stage", "BSK 3rd Stage"
AddBLR dictBlr, "bsk 2nd stage|bsk ii stage|bsk second stage|banashankari 2nd stage|banashankari second stage", "BSK 2nd Stage"
AddBLR dictBlr, "basavanagudi|basavangudi", "Basavanagudi"
AddBLR dictBlr, "padmanabhanagar", "Padmanabhanagar"
AddBLR dictBlr, "uttarahalli|uttarhalli", "Uttarahalli"
AddBLR dictBlr, "chikkalsandra", "Chikkalsandra"
AddBLR dictBlr, "kumaraswamy layout|ks layout|kumarswamy layout", "Kumarswamy Layout"
AddBLR dictBlr, "arekere", "Arekere"
AddBLR dictBlr, "hulimavu|hulimau", "Hulimavu"
AddBLR dictBlr, "puttenahalli", "Puttenahalli"
AddBLR dictBlr, "bommanahalli", "Bommanahalli"
AddBLR dictBlr, "btm|b t m", "BTM Layout"
AddBLR dictBlr, "bannerghatta road|bannerghatta", "Bannerghatta Road"
AddBLR dictBlr, "south end circle|south end|south end road", "South End Circle"
AddBLR dictBlr, "iti layout", "ITI Layout"
AddBLR dictBlr, "katriguppe|katrguppe", "Katriguppe"

' East Bangalore
AddBLR dictBlr, "indiranagar|indira nagar", "Indiranagar"
AddBLR dictBlr, "old madras road|omr", "Old Madras Road"
AddBLR dictBlr, "cv raman nagar|c v raman nagar", "CV Raman Nagar"
AddBLR dictBlr, "kaggadasapura|kaggad", "Kaggadasapura"
AddBLR dictBlr, "ramamurthy nagar|rm nagar", "Ramamurthy Nagar"
AddBLR dictBlr, "horamavu|hormavu", "Horamavu"
AddBLR dictBlr, "banaswadi|bannaswadi", "Banaswadi"
AddBLR dictBlr, "kalyan nagar|kalyannagar", "Kalyan Nagar"
AddBLR dictBlr, "hrbr layout", "HRBR Layout"
AddBLR dictBlr, "hbr layout", "HBR Layout"
AddBLR dictBlr, "hebbal", "Hebbal"
AddBLR dictBlr, "nagawara|nagavara", "Nagawara"
AddBLR dictBlr, "yelahanka|yelanka", "Yelahanka"
AddBLR dictBlr, "vidyaranyapura", "Vidyaranayapura"

' Whitefield / IT Corridors
AddBLR dictBlr, "whitefield|white field", "Whitefield"
AddBLR dictBlr, "kadugodi|kadu godhi", "Kadugodi"
AddBLR dictBlr, "hoodi", "Hoodi"
AddBLR dictBlr, "kundalahalli|kundalhalli", "Kundalahalli"
AddBLR dictBlr, "marathahalli|marathalli|marath", "Marathahalli"
AddBLR dictBlr, "brookefield", "Brookefield"
AddBLR dictBlr, "varthur", "Varthur"
AddBLR dictBlr, "gunjur", "Gunjur"
AddBLR dictBlr, "sarjapur road|sarjapur rd|sarjapur|sarjapura", "Sarjapur Road"
AddBLR dictBlr, "bellandur", "Bellandur"
AddBLR dictBlr, "haralur|harlur", "Haralur"
AddBLR dictBlr, "kaikondrahalli", "Kaikondrahalli"

' Electronic City Belt
AddBLR dictBlr, "electronic city|e city", "Electronic City"
AddBLR dictBlr, "bommasandra", "Bommasandra"
AddBLR dictBlr, "attibele|attibelle", "Attibele"
AddBLR dictBlr, "anekal", "Anekal"
AddBLR dictBlr, "huskur", "Huskur"
AddBLR dictBlr, "singasandra", "Singasandra"
AddBLR dictBlr, "kudlu|kudlu gate", "Kudlu"

' Central Market Belt
AddBLR dictBlr, "kr market|k r market|city market", "KR Market"
AddBLR dictBlr, "chickpet|chikpet", "Chickpet"
AddBLR dictBlr, "avenue road", "Avenue Road"
AddBLR dictBlr, "lalbagh|lal bagh", "Lalbagh"

' CBD / MG Road
AddBLR dictBlr, "mg road|m g road", "MG Road"
AddBLR dictBlr, "brigade road|brigade rd", "Brigade Road"
AddBLR dictBlr, "church street", "Church Street"
AddBLR dictBlr, "commercial street", "Commercial Street"
AddBLR dictBlr, "palace road", "Palace Road"
AddBLR dictBlr, "residency road", "Residency Road"
AddBLR dictBlr, "richmond town", "Richmond Town"
AddBLR dictBlr, "st marks road|st mark road", "St Marks Road"
AddBLR dictBlr, "race course road|race course", "Race Course Road"

' Additional localities
AddBLR dictBlr, "hosur road", "Hosur Road"
AddBLR dictBlr, "begur", "Begur"
AddBLR dictBlr, "kothanur", "Kothanur"
AddBLR dictBlr, "mathikere|mattikere", "Mathikere"
AddBLR dictBlr, "soladevanahalli", "Soladevanahalli"
AddBLR dictBlr, "kadubeesanahalli", "Kadubeesanahalli"
AddBLR dictBlr, "bagaluru", "Bagaluru"
AddBLR dictBlr, "koramangala", "Koramangala"
AddBLR dictBlr, "madivala|madiwala", "Madivala"
AddBLR dictBlr, "sanjaynagar", "Sanjaynagar"
AddBLR dictBlr, "itpl", "ITPL"
AddBLR dictBlr, "silk board", "Silk Board"
AddBLR dictBlr, "tavarekere|tavarakere", "Tavarekere"
AddBLR dictBlr, "chikkabanavara", "Chikkabanavara"
AddBLR dictBlr, "chikkamarali village", "Chikkamarali Village"
AddBLR dictBlr, "rajanukunte", "Rajanukunte"
AddBLR dictBlr, "hesaragatta|hesaraghatta", "Hesaragatta"
AddBLR dictBlr, "hulahalli", "Hulahalli"
AddBLR dictBlr, "gokul", "Gokul"
AddBLR dictBlr, "kempapura", "Kempapura"
AddBLR dictBlr, "kumbalagodu", "Kumbalagodu"
AddBLR dictBlr, "bidadi", "Bidadi"
AddBLR dictBlr, "vasanthapura", "Vasanthapura"
AddBLR dictBlr, "magadi town", "Magadi Town"
AddBLR dictBlr, "srinivasa nagar", "Srinivasa Nagar"

' ==============================
' ?? TOP TIER UNIVERSITIES (Mapped to Locality)
' ==============================

AddBLR dictBlr, "christ university|christ college|christu|christ hostel", "Hosur Road"
AddBLR dictBlr, "christ academy", "Begur"
AddBLR dictBlr, "reva university|reva campus|reva hostel", "Yelahanka"
AddBLR dictBlr, "rv college of engineering|rvce|rvu|rv campus|rv engg clg|r v engg clg", "Mysore Road"
AddBLR dictBlr, "jain university|jain campus|jain hostel|jain jayanagar|jain(deemed to be university)", "Jayanagar"
AddBLR dictBlr, "jain university kanakapura", "Kanakapura Road"
AddBLR dictBlr, "pes university|pes college|pesu|pes hostel|pes institute", "Banashankari"
AddBLR dictBlr, "cmr university|cmru", "Bagaluru"
AddBLR dictBlr, "cmrit|c m r institute of technology|cmrit hostel", "Brookefield"
AddBLR dictBlr, "ms ramaiah|msrit|ramaiah medical|ramaiah university", "Mathikere"
AddBLR dictBlr, "dayananda sagar|dsi|dsce|dayananda sagar college", "Kumarswamy Layout"
AddBLR dictBlr, "acharya institute|acharya college|acharya campus", "Soladevanahalli"
AddBLR dictBlr, "presidency university", "Yelahanka"
AddBLR dictBlr, "presidency college", "Hebbal"
AddBLR dictBlr, "new horizon college|nhce|new horizon", "Kadubeesanahalli"
AddBLR dictBlr, "nitte meenakshi|nitte institute", "Yelahanka"
AddBLR dictBlr, "mount carmel college|mcc", "Palace Road"
AddBLR dictBlr, "st joseph|sjcc|sju campus|st josephs college", "Residency Road"
AddBLR dictBlr, "st joseph arts|st joseph commerce", "Brigade Road"
AddBLR dictBlr, "kslu|karnataka state law university", "Hubballi (Not BLR override)"
AddBLR dictBlr, "nmkrv college", "Jayanagar"

' ==============================
' ?? ENGINEERING CLUSTERS
' ==============================

AddBLR dictBlr, "bms college of engineering|bmsce", "Basavanagudi"
AddBLR dictBlr, "bmsit", "Yelahanka"
AddBLR dictBlr, "bit bangalore institute technology", "Kengeri"
AddBLR dictBlr, "sir mvit|sir mv institute", "Yelahanka"
AddBLR dictBlr, "iit bangalore|iit research campus", "Malleshwaram"
AddBLR dictBlr, "east west engineering college", "Magadi Road"
AddBLR dictBlr, "oxford engineering college", "Bommanahalli"
AddBLR dictBlr, "impact college", "Yelahanka"
AddBLR dictBlr, "vydehi institute|vydehi medical", "Whitefield"

' ==============================
' ?? MEDICAL & PARAMEDICAL CLUSTERS
' ==============================

AddBLR dictBlr, "st johns medical college|st johns hospital", "Koramangala"
AddBLR dictBlr, "ms ramaiah medical|ramaiah hospital", "Mathikere"
AddBLR dictBlr, "vydehi medical college", "Whitefield"
AddBLR dictBlr, "nimhans|nimhans campus", "Lalbagh Road"
AddBLR dictBlr, "sanjay gandhi institute", "Jayanagar"
AddBLR dictBlr, "fortis bg road|fortis bannergatta", "Bannerghatta Road"
AddBLR dictBlr, "apollo jayanagar", "Jayanagar"
AddBLR dictBlr, "apollo bannergatta", "Bannerghatta Road"

' ==============================
' ?? TOP ARTS / COMMERCE / LAW
' ==============================

AddBLR dictBlr, "st josephs law", "Brigade Road"
AddBLR dictBlr, "kristu jayanti college|kristu jayanti university", "Kothanur"
AddBLR dictBlr, "msw college", "Seshadripuram"
AddBLR dictBlr, "cmr law", "Kalyan Nagar"
AddBLR dictBlr, "baldwin women college", "Richmond Town"
AddBLR dictBlr, "bishop cotton girls", "St Marks Road"
AddBLR dictBlr, "bishop cotton boys", "St Marks Road"
AddBLR dictBlr, "rns first grade college|rnsfgc|rns|rnsfgca", "Channasandra"
AddBLR dictBlr, "sri krishna degree college", "ITI Layout"

' ==============================
' ?? EDUCATION BELTS (Major Clusters)
' ==============================

AddBLR dictBlr, "koramangala forum|forum mall", "Koramangala"
AddBLR dictBlr, "msrit road", "Mathikere"
AddBLR dictBlr, "hebbal campus belt", "Hebbal"
AddBLR dictBlr, "itpl campus", "ITPL"
AddBLR dictBlr, "marathahalli bridge", "Marathahalli"
AddBLR dictBlr, "btm water tank", "BTM Layout"
AddBLR dictBlr, "mg road metro", "MG Road"
AddBLR dictBlr, "brindavan college", "Yelahanka"

' ==============================
' ?? UNIVERSITY HOSTELS ONLY – MAP TO LOCALITY
' ==============================

AddBLR dictBlr, "pes hostel", "Banashankari"
AddBLR dictBlr, "christ hostel", "Hosur Road"
AddBLR dictBlr, "reva hostel", "Yelahanka"
AddBLR dictBlr, "jain hostel", "Jayanagar"
AddBLR dictBlr, "cmrit hostel", "Brookefield"
AddBLR dictBlr, "ms ramaiah hostel", "Mathikere"
AddBLR dictBlr, "sjcc hostel|st joseph hostel", "Residency Road"
AddBLR dictBlr, "st johns hostel", "Koramangala"

' Additional neighborhoods
AddBLR dictBlr, "ramohalli", "Ramohalli"
AddBLR dictBlr, "gururayanapura", "Gururayanapura"
AddBLR dictBlr, "koluru village", "Koluru Village"

End Sub

Private Sub AddBLR(dict As Object, k As String, v As String)
    If Not dict.Exists(k) Then dict.Add k, v
End Sub

Private Function DetectBLRLocality(ByVal lower As String, dictBlr As Object) As String
    Dim k As Variant, patterns() As String, p As Variant
    Dim bestMatch As String, bestLength As Integer
    
    ' Find the longest matching pattern (to prioritize specific matches)
    For Each k In dictBlr.Keys
        patterns = Split(k, "|")
        For Each p In patterns
            If p <> "" Then
                If InStr(lower, p) > 0 Then
                    If Len(p) > bestLength Then
                        bestMatch = dictBlr(k)
                        bestLength = Len(p)
                    End If
                End If
            End If
        Next p
    Next k
    
    DetectBLRLocality = bestMatch
End Function

Private Function FallbackLocality(ByVal raw As String, ByVal district As String, ByVal re As Object) As String
    Dim work As String
    Dim noise, nn
    Dim parts() As String, j As Long, p As String
    Dim best As String
    Dim candidates As Collection
    Dim candidate As String
    
    Set candidates = New Collection
    
    work = LCase(raw)
    
    ' Remove special characters first
    work = Replace(work, "#", " ")
    work = Replace(work, "&", " ")
    work = Replace(work, "(", " ")
    work = Replace(work, ")", " ")
    
    ' Remove PIN codes
    re.Pattern = "\b\d{6}\b"
    work = re.Replace(work, " ")
    
    ' Remove common noise words / tokens
    noise = Array("-", "/", "flat", "no.", "no", "house", "landmark", "near", "opp", _
                  "opposite", "behind", "beside", "next to", "bus stop", "police", "post", _
                  "plot", "site", "apartment", "apt", "h.no", "street", "st", "road", "rd", _
                  "main", "stage", "phase", "block", "college", "university", "institute", _
                  "school", "academy", "engineering", "medical", "hospital", "hostel", _
                  "campus", "sciences", "arts", "commerce", "science", "degree", "pg", "ug", _
                  "deemed", "autonomous", "auto", "mous")
    For Each nn In noise
        work = Replace(work, nn, " ")
    Next nn
    
    ' Remove standalone numbers
    re.Pattern = "\b\d+\b"
    work = re.Replace(work, " ")
    
    ' Normalize commas and spaces
    re.Pattern = "[,;]+"
    work = re.Replace(work, ",")
    
    re.Pattern = "\s+"
    work = Trim(re.Replace(work, " "))
    
    If work = "" Then Exit Function
    
    ' Split by comma and spaces to find locality candidates
    parts = Split(work, ",")
    
    For j = LBound(parts) To UBound(parts)
        p = Trim(parts(j))
        If p <> "" Then
            ' Skip if it contains the district word itself
            If district <> "" Then
                If InStr(LCase(p), LCase(Split(district, " ")(0))) > 0 Then GoTo ContinueLoop
            End If
            
            ' Skip generic Karnataka references
            If InStr(p, "karnataka") > 0 Then GoTo ContinueLoop
            
            If Not IsNoiseFragment(p) Then
                ' Prefer longer, more specific strings
                If Len(p) >= 4 Then
                    On Error Resume Next
                    candidates.Add p
                    On Error GoTo 0
                End If
            End If
        End If
ContinueLoop:
    Next j
    
    ' Get the best candidate (first meaningful one, or longest)
    If candidates.Count > 0 Then
        best = candidates(1)
        For j = 2 To candidates.Count
            candidate = candidates(j)
            ' Prefer first position, but consider longer if significantly better
            If Len(candidate) > Len(best) + 3 Then
                best = candidate
            End If
        Next j
        
        FallbackLocality = Application.WorksheetFunction.Proper(best)
    End If
    
End Function

Private Function IsNoiseFragment(s As Variant) As Boolean
    Dim low As String: low = LCase(Trim(CStr(s)))
    Dim noiseWords, w
    
    If low = "" Then IsNoiseFragment = True: Exit Function
    If Len(low) < 3 Then IsNoiseFragment = True: Exit Function
    
    ' Filter out educational institution keywords
    If InStr(low, "college") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "university") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "institute") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "school") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "academy") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "engineering") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "medical") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "hospital") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "hostel") > 0 Then IsNoiseFragment = True: Exit Function
    If InStr(low, "campus") > 0 Then IsNoiseFragment = True: Exit Function
    
    noiseWords = Array("near", "opp", "main", "cross", "block", "stage", "phase", _
                       "layout", "area", "post", "village", "dist", "district", _
                       "hobli", "taluk", "pin", "code", "karnataka", "state", _
                       "industrial", "town", "suburb", "auto", "mous", "sciences", _
                       "arts", "commerce", "science", "degree", "pg", "ug")
    
    For Each w In noiseWords
        If low = w Then IsNoiseFragment = True: Exit Function
    Next w
End Function

Private Function CreateAuditSheet(SheetName As String) As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(SheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set CreateAuditSheet = Sheets.Add
    CreateAuditSheet.Name = SheetName
End Function

Private Function NormalizeAbbrev(loc As String) As String
    Dim t As String: t = UCase(Trim(loc))

    ' Remove spaces between abbreviation letters
    t = Replace(t, "R R", "RR")
    t = Replace(t, "J P", "JP")
    t = Replace(t, "H S R", "HSR")
    t = Replace(t, "H S", "HS")
    t = Replace(t, "K R", "KR")
    t = Replace(t, "M G", "MG")
    t = Replace(t, "C V", "CV")
    t = Replace(t, "H A L", "HAL")
    t = Replace(t, "B D A", "BDA")
    t = Replace(t, "A E C S", "AECS")
    t = Replace(t, "B B M P", "BBMP")
    t = Replace(t, "B S K", "BSK")
    t = Replace(t, "B T M", "BTM")
    t = Replace(t, "I T I", "ITI")
    t = Replace(t, "I T P L", "ITPL")

    ' Convert to Proper Case for output consistency
    NormalizeAbbrev = Application.WorksheetFunction.Proper(t)
End Function


