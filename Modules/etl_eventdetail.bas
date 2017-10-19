Attribute VB_Name = "Create Event Detail1"
Function ItemLocList(Optional shared_input_week As String)

Dim objRecordset As ADODB.Recordset
Set objRecordset = New ADODB.Recordset
    
Dim EventCode As Variant
Dim Week As Variant
Dim Version As Variant
Dim FYear As Variant
Dim Items As Variant
Dim item As Variant
Dim loc As Variant
Dim arry As Variant
Dim ArryRINs As Variant
Dim x As Single 'store array number
Dim y As Single 'rin array number
Dim strSQL As String
Dim input_week As String
    
'input_week = InputBox("Type Ad Weeks to import via SQL in") 'old way
input_week = shared_input_week 'shared_input_week on kickoff sub

If input_week = "" Then
    input_week = InputBox("Type Ad Weeks to import via SQL in")
    Else
        'do nothing input_week entered
End If
    
objRecordset.ActiveConnection = CurrentProject.Connection
objRecordset.Open ("Select * " & _
"From Event_Head " & _
"Where [Week] in " & input_week & " ")

'Loop through all cells in AdStores column
'For Each Record In objRecordset.Fields("Ad_Stores")
Do While objRecordset.EOF = False
    'EventCode = objRecordset.Fields("EventCode")
    Head_Index = objRecordset.Fields("Index")
    Week = objRecordset.Fields("Week")
    Version = objRecordset.Fields("Version")
    FYear = objRecordset.Fields("FYear")
    Item_Desc = objRecordset.Fields("Item Description")
    arry = Split(objRecordset.Fields("Ad_Stores"), ",")
    ArryRINs = Split(objRecordset.Fields("Items"), ",", , vbTextCompare)
        For x = LBound(arry) To UBound(arry)
            For y = LBound(ArryRINs) To UBound(ArryRINs)
                'n = n + 1 'row number
                loc = Val(Trim(arry(x)))
                item = Trim(ArryRINs(y))
                
                strSQL = "Insert Into Event_Detail " _
                & "(Head_Index, Week, Version, FYear, Loc, Item, Item_Desc) Values" _
                & "(""" & Head_Index & """,""" & Week & """,""" & Version & """,""" & FYear & """,""" & loc & """,""" & item & """,""" & Item_Desc & """);"
                Debug.Print strSQL
                CurrentDb.Execute strSQL
            Next y
        Next x
objRecordset.MoveNext
Loop
objRecordset.Close

End Function

