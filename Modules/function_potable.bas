Attribute VB_Name = "MultiPO"
Function MultiPOArray()

Dim objRecordset As ADODB.Recordset
Set objRecordset = New ADODB.Recordset
    
Dim order_nbr As Variant
Dim item As Variant
Dim loc As Variant
Dim Item_Loc As String
Dim arry() As Variant
Dim upr As Variant
Dim dist As Variant
Dim output As Variant
Dim newoutput As Variant
Dim x As Single
Dim y As Single
Dim i As Integer
Dim strSQL As String
    
objRecordset.ActiveConnection = CurrentProject.Connection
objRecordset.Open ("Orders_report")
    
y = 0
    
While (Not objRecordset.EOF)
item = objRecordset.Fields("Item")
loc = objRecordset.Fields("Loc")
Item_Loc = objRecordset.Fields("Item_Loc")
output = ""
    upr = DCount("Item_Loc", "orders_report", "[Item_Loc] = '" & Item_Loc & "'") - 1
    Debug.Print (upr)
    ReDim arry(0 To upr)
    
    Do While arry(upr) = ""
        If objRecordset.Fields("Item_Loc") = Item_Loc Then
            arry(x) = objRecordset.Fields("order_nbr").Value
            newoutput = objRecordset.Fields("order_nbr").Value
        End If
        objRecordset.MoveNext
        x = x + 1
        output = output & newoutput & ", "
    Loop
    
    x = 0
    y = y + 1
   Debug.Print output
strSQL = "Insert Into PO_List " _
& "(Item, Loc, PO_List) Values" _
& "(""" & item & """, """ & loc & """, """ & output & """);"
CurrentDb.Execute strSQL
Wend
objRecordset.Close
End Function
