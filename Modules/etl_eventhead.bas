Attribute VB_Name = "Create Event Head1"
Function Tier_to_Store(Optional shared_input_week As String)
'Debug.Print Now

Dim objRecordset As ADODB.Recordset
Set objRecordset = New ADODB.Recordset

Dim arry As Variant
Dim EventCode As Variant
Dim Week As Variant
Dim Version As Variant
Dim FYear As Variant
Dim Tiers As Variant
Dim ExceptionList As Variant
Dim Items As Variant
Dim IncludeOrExclude As String
Dim i As Integer
Dim x As Single
Dim incl As String
Dim excl As String
Dim newoutput As String
Dim output As String
Dim strSQL As String
Dim input_week As String
Dim Regex As Object
Dim pattern As String

'Regex pattern to handle Tier info listed with () ex., U(17)
pattern = "\(.*"

Set Regex = CreateObject("vbscript.regexp")
With Regex
    .Multiline = False
    .Global = False
    .IgnoreCase = True
    .pattern = pattern
End With

input_week = shared_input_week 'shared_input_week on kickoff sub

If input_week = "" Then
    input_week = InputBox("Type Ad Weeks to import via SQL in")
Else
    'do nothing input_week entered
End If

objRecordset.ActiveConnection = CurrentProject.Connection
'objRecordset.Open ("Main")
objRecordset.Open ("Select [Week Prefix], [Week Suffix], [FY], [Tiers], Nz([Exclusions and Inclusions]) as [Exclusions and Inclusions], [DMM], [Group Number], Replace(Main.[Item Description], '""', """") as [Item Description], [Featured RIN], [Page], nz([Strike Point]) as [Strike Point] " & _
"From Main " & _
"Where [Featured RIN] is not null " & _
"And [Tiers] Is Not Null " & _
"And [FY] = 'FY2017' " & _
"And [Week Prefix] in " & input_week & " " & _
"And IsNumeric(Left([Featured RIN], 1)) ")
'"And [Featured RIN] <> 'VARIOUS' " & _
'"And [Featured RIN] <> 'NA' ")

Do While objRecordset.EOF = False
    Week = objRecordset.Fields("Week Prefix")
    Version = objRecordset.Fields("Week Suffix")
    FYear = objRecordset.Fields("FY")
    Tiers = objRecordset.Fields("Tiers")
    ExceptionList = rdotdigit(objRecordset.Fields("Exclusions and Inclusions")) 'rdotdigit is a tempory patch until cleaned in AdDB
    Items = objRecordset.Fields("Featured RIN")
    Page = objRecordset.Fields("Page")
    StrikePoint = objRecordset.Fields("Strike Point")
    DMM = objRecordset.Fields("DMM")
    GroupNum = objRecordset.Fields("Group Number")
    Item_Desc = objRecordset.Fields("Item Description")
    'Debug.Print Item_Desc
    
    'Loop trhough Tiers
    arry = Split(objRecordset.Fields("Tiers"), ",", , vbTextCompare)
        For x = LBound(arry) To UBound(arry)
            Debug.Print arry(x)
            'newoutput = TierStore(Trim(Left(arry(x), 4))) OLD WAY
            newoutput = TierStore(Trim(Regex.Replace(arry(x), "")))
            Debug.Print (newoutput)
            
            If InStr(1, output, newoutput) > 0 Then
            'do nothing - already included
            Else
                output = output & newoutput & ", "
            End If
            Debug.Print (output)
        Next x
        
    'Loop through Inclusions & Exclusions
    'If both Exclusions & Inclusions present
    If InStr(1, ExceptionList, "Excluding") > 0 And InStr(1, ExceptionList, "; Including") > 0 Then
        excl = Left(ExceptionList, InStr(ExceptionList, ";") - 1)
        arry = Split(excl, ",", , vbTextCompare)
        For x = LBound(arry) To UBound(arry)
            arry(x) = Replace(arry(x), "Excluding: ", "")
            'Debug.Print (arry(x))
            newoutput = Exceptions(Trim(Replace(arry(x), ";", ""))) 'Trim(arry(x))
                If InStr(1, output, newoutput) > 0 And Len(newoutput) > 1 Then
                    output = Replace(output, newoutput & ", ", "")
                Else
                    'do nothing - already included
                End If
        Next x
    Else 'If only Exclusions
    If InStr(1, ExceptionList, "Excluding") > 0 Then
        excl = Trim(Mid(ExceptionList, InStr(ExceptionList, "Excluding: "), Len(ExceptionList) - InStr(ExceptionList, "Excluding: ") + 1))
        arry = Split(excl, ",", , vbTextCompare)
        For x = LBound(arry) To UBound(arry)
            arry(x) = Replace(arry(x), "Excluding: ", "")
            Debug.Print (arry(x))
            newoutput = Exceptions(Trim(Replace(arry(x), ";", ""))) 'Trim(arry(x))
                If InStr(1, output, newoutput) > 0 And Len(newoutput) > 1 Then
                    output = Replace(output, newoutput & ", ", "")
                Else
                    'do nothing - already included
                End If
        Next x
    End If
    End If

    'If Inclusions Exist
    If InStr(1, ExceptionList, "Including") > 0 Then
        incl = Trim(Mid(ExceptionList, InStr(ExceptionList, "Including: "), Len(ExceptionList) - InStr(ExceptionList, "Including: ") + 1))
        arry = Split(incl, ",", , vbTextCompare)
        For x = LBound(arry) To UBound(arry)
            arry(x) = Replace(arry(x), "Including: ", "")
            'Debug.Print (arry(x))
            newoutput = Exceptions(Trim(Replace(arry(x), ";", ""))) 'Trim(arry(x))
                If InStr(1, output, newoutput) > 0 Then
                    'do nothing - already included
                Else
                    output = output & newoutput & ", "
                End If
                Debug.Print (output)
        Next x
    End If
    
    'Remove superfluous commas...
    If Left(output, 2) = ", " Then output = Mid(output, 3, 500)
    Do Until Right(output, 2) <> ", "
        output = Left(output, Len(output) - 2)
    Loop
                
    'Insert into Table
    strSQL = "Insert Into Event_Head " _
    & "(Week, Version, FYear, Tiers, ExceptionList, Page, [Strike Point], DMM, GroupNum, [Item Description], Items, Ad_Stores) Values" _
    & "(""" & Week & """,""" & Version & """,""" & FYear & """,""" & Tiers & """,""" & ExceptionList & """,""" & Page & """,""" & StrikePoint & """,""" & DMM & """,""" & GroupNum & """,""" & Item_Desc & """,""" & Items & """,""" & output & """);"
    'Debug.Print strSQL
    CurrentDb.Execute strSQL

    objRecordset.MoveNext
    Debug.Print (output)
    output = ""
Loop

objRecordset.Close
'Debug.Print Now
End Function

Function Exceptions(sReplace As String)

    Select Case sReplace
        Case "Web Store"
            Exceptions = "780"
        Case "Bethesda"
            Exceptions = "745"
        Case "Norfolk"
            Exceptions = "010"
        Case "Pearl"
            Exceptions = "437"
        Case "San Diego"
            Exceptions = "305"
        Case "Guam"
            Exceptions = "440"
        Case "Jax"
            Exceptions = "164"
        Case "Little Creek"
            Exceptions = "016"
        Case "Oceana"
            Exceptions = "034"
        Case "Pensacola"
            Exceptions = "191"
        Case "Yoko"
            Exceptions = "464"
        Case "Yokosuka"
            Exceptions = "462"
        Case "Bangor"
            Exceptions = "407"
        Case "Grt Lakes"
            Exceptions = "110"
        Case "Hueneme"
            Exceptions = "335"
        Case "Lemoore"
            Exceptions = "366"
        Case "Mayport"
            Exceptions = "169"
        Case "Memphis"
            Exceptions = "263"
        Case "N. Island"
            Exceptions = "292"
        Case "Whidbey"
            Exceptions = "393"
        Case "Orlando"
            Exceptions = "265"
        Case "Annapolis"
            Exceptions = "751"
        Case "Bremerton"
            Exceptions = "398"
        Case "Charleston"
            Exceptions = "236"
        Case "Corpus"
            Exceptions = "274"
        Case "Everett"
            Exceptions = "409"
        Case "Gulfport"
            Exceptions = "722"
        Case "N London"
            Exceptions = "067"
        Case "N Orleans"
            Exceptions = "730"
        Case "Newport"
            Exceptions = "055"
        Case "Pax River"
            Exceptions = "093"
        Case "Fallon"
            Exceptions = "384"
        Case "Key West"
            Exceptions = "227"
        Case "Kings Bay"
            Exceptions = "183"
        Case "Meridian"
            Exceptions = "273"
        Case "Mitchell Fld"
            Exceptions = "057"
        Case "Monterey"
            Exceptions = "390"
        Case "Portsmouth"
            Exceptions = "028"
        Case "Whiting Fld"
            Exceptions = "246"
        Case "Atsugi"
            Exceptions = "481"
        Case "Bahrain"
            Exceptions = "652"
        Case "Naples"
            Exceptions = "650"
        Case "Rota"
            Exceptions = "716"
        Case "Sasebo"
            Exceptions = "490"
        Case "Sigonella"
            Exceptions = "138"
        Case Else
            Exceptions = ""
    End Select

End Function


Function TierStore(sReplace As String)

    Select Case sReplace
        Case "F"
            TierStore = "010, 437, 305"
        Case "2.0"
            TierStore = "016, 034, 164, 191, 745"
        Case "2.1"
            TierStore = "440, 464"
        Case "3.0"
            TierStore = "292, 393, 407, 335, 169, 263, 110, 751, 93, 366"
        Case "3.1"
            TierStore = "650, 716, 490"
        Case "4.0"
            TierStore = "730, 236, 409, 274, 067, 055, 722"
        Case "4.1"
            TierStore = "138"
        Case "5.0"
            TierStore = "057, 227, 183, 028, 291, 243, 293, 384, 200, 350, 434, 329"
        Case "5.1"
            TierStore = "652, 715, 481, 459, 452, 133, 485"
        Case "P"
            TierStore = "437"
        Case "O"
            TierStore = "265"
        Case "G"
            TierStore = "221"
        Case "U"
            TierStore = "440"
        Case "B"
            TierStore = "745"
        Case "Y"
            TierStore = "464"
        Case "IF"
            TierStore = "412, 43, 305, 442"
        Case "IIF"
            TierStore = "726, 462, 191, 397"
        Case "IIIF"
            TierStore = "169, 650, 274, 138, 489, 395, 716, 386, 728"
        Case "I"
            TierStore = "010, 437, 305, 745"
        Case "II"
            TierStore = "016, 034, 164, 191, 440, 464"
        Case "III"
            TierStore = "407, 110, 335, 366, 169, 263, 292, 393"
        Case "IV"
            TierStore = "265, 751, 398, 236, 274, 409, 722, 067, 730, 055, 093"
        Case "V"
            TierStore = "384, 227, 183, 243, 057, 390, 028, 246"
        Case "VI"
            TierStore = "481, 652, 650, 716, 490, 138"
        Case Else
            TierStore = ""
    End Select

End Function

