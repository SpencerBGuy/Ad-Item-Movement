Attribute VB_Name = "basConcatenate"
Option Compare Database

Function Concatenate(pstrSQL As String, _
        Optional pstrDelim As String = ", ") _
        As String

    Dim rs As New ADODB.Recordset
    rs.Open pstrSQL, CurrentProject.Connection
'SELECT item, loc, Concatenate("SELECT order_nbr FROM pos WHERE item =""" & [item] & """ AND order_nbr =""" & [order_nbr] & """") AS po_list
'FROM pos;
    Dim strConcat As String 'build return string
    With rs
    'Debug.Print (rs.Source)
    'Debug.Print (rs.Fields(0))
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
            Debug.Print (rs.Fields(0))
                strConcat = strConcat & _
                .Fields(0) & pstrDelim
                .MoveNext
                Debug.Print (strConcat)
            Loop
        End If
        .Close
    End With
    Set rs = Nothing

    'Debug.Print (Len(strConcat))

    If Len(strConcat) > 0 Then
        strConcat = Left(strConcat, _
        Len(strConcat) - Len(pstrDelim))
    End If
    Concatenate = strConcat
End Function

