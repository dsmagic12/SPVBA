Option Compare Database

Const dbAttachment As Long = 101
Const dbBigInt As Long = 16
Const dbBinary As Long = 9
Const dbBoolean As Long = 1
Const dbByte As Long = 2
Const dbChar As Long = 18
Const dbComplexByte As Long = 102
Const dbComplexDecimal As Long = 108
Const dbComplexDouble As Long = 106
Const dbComplexGUID As Long = 107
Const dbComplexInteger As Long = 103
Const dbComplexLong As Long = 104
Const dbComplexSingle As Long = 105
Const dbComplexText As Long = 109
Const dbCurrency As Long = 5
Const dbDate As Long = 8
Const dbDecimal As Long = 20
Const dbDouble As Long = 7
Const dbFloat As Long = 21
Const dbGUID As Long = 15
Const dbInteger As Long = 3
Const dbLong As Long = 4
Const dbLongBinary As Long = 11
Const dbMemo As Long = 12
Const dbNumeric As Long = 19
Const dbSingle As Long = 6
Const dbText As Long = 10
Const dbTime As Long = 22
Const dbTimeStamp As Long = 23
Const dbVarBinary As Long = 17


Public Sub moveData()
    Dim DestinationTableName, SourceTableName
    DestinationTableName = "[Field Types Demo]"
    SourceTableName = "[localField Types Demo]"
    Dim s As DAO.Recordset, o As DAO.Recordset, existing As DAO.Recordset
    Dim s2 As DAO.Recordset2, o2 As DAO.Recordset2
    Dim sFld As DAO.Field, oFld As DAO.Field
    Set s = CurrentDb.OpenRecordset("SELECT Title, Person, [Comments Plain], [Comments Rich], [Choice Dropdown], [Choice Radio], Number, Currency, [Date Only], [Date and Time], [Lookup Single], [Yes No]  FROM " & SourceTableName, dbOpenDynaset)
    Set o = CurrentDb.OpenRecordset("SELECT * FROM " & DestinationTableName, dbOpenDynaset)
    Dim skip
    
    Do While Not s.EOF
        Debug.Print Now() & "... working on |" & s!Title.Value & "|"
            o.AddNew
                For Each sFld In s.Fields
                    If IsNull(s.Fields(sFld.Name).Value) = False Then
                        For Each oFld In o.Fields
                            If sFld.Name = oFld.Name Then
                                Select Case o.Fields(sFld.Name).Type
                                    Case dbAttachment, dbComplexByte, dbComplexDecimal, dbComplexDouble, dbComplexGUID, dbComplexInteger, dbComplexLong, dbComplexSingle, dbComplexText
                                        Set s2 = s.Fields(sFld.Name).Value
                                        Set o2 = o.Fields(oFld.Name).Value
                                        'clear existing values from our output multivalue field
                                        Do While Not o2.EOF
                                            o2.Delete
                                            o2.MoveNext
                                        Loop
                                        ' populate each value in our source multivalue field into our output multivalue field
                                        Do While Not s2.EOF
                                            o2.AddNew
                                                o2!Value.Value = s2!Value.Value
                                            o2.Update
                                            s2.MoveNext
                                        Loop
                                        s2.Close
                                        Set s2 = Nothing
                                        o2.Close
                                        Set o2 = Nothing
                                    Case dbBigInt, dbBinary, dbBoolean, dbByte, dbChar, dbCurrency, dbDate, dbDecimal, dbDouble, dbFloat, dbGUID, dbInteger, dbLong, dbLongBinary, dbMemo, dbNumeric, dbSingle, dbText, dbTime, dbTimeStamp, dbVarBinary
                                        ' for these simple data types, we can use simple syntax to set field values
                                        o.Fields(oFld.Name).Value = s.Fields(sFld.Name).Value
                                End Select
                            End If
                        Next
                    End If
                Next
            o.Update
        s.MoveNext
    Loop
End Sub
