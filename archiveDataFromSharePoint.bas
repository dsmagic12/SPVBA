Option Compare Database
Const ATTACHMENT_PATHS_TABLE_NAME As String = "AttachmentPaths"
Const ATTACHMENTS_PATH As String = "C:\Users\USERNAME\Desktop\SP Archive\attachments"

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


Public Sub archiveDataFromSharePoint()
    Dim DestinationTableName, SourceTableName
    DestinationTableName = "[localField Types Demo]"
    SourceTableName = "[Field Types Demo]"
    Dim s As DAO.Recordset, o As DAO.Recordset, oAttachmentPaths As DAO.Recordset, existing As DAO.Recordset
    Dim s2 As DAO.Recordset2, o2 As DAO.Recordset2
    Dim sFld As DAO.field, oFld As DAO.field
    Dim sAttachments As DAO.Field2, oAttachments As DAO.Field2
    Set s = CurrentDb.OpenRecordset("SELECT Attachments, ID, Title, Person, [Comments Plain], [Comments Rich], [Choice Dropdown], [Choice Radio], Number, Currency, [Date Only], [Date and Time], [Lookup Single], [Yes No]  FROM " & SourceTableName, dbOpenDynaset)
    Set o = CurrentDb.OpenRecordset("SELECT * FROM " & DestinationTableName, dbOpenDynaset)
    Set oAttachmentPaths = CurrentDb.OpenRecordset("SELECT * FROM [" & ATTACHMENT_PATHS_TABLE_NAME & "]", dbOpenDynaset)
    
    Dim skip
    Dim iMax, iCurr
    iCurr = 1
    iMax = 10000
    
    Do While Not s.EOF
        If iCurr >= iMax Then
            Exit Do
        End If
        DoEvents
        Debug.Print Now() & "... working on |" & s!ID.Value & "|"
        ' Delete existing archived records if they exist
        CurrentDb.Execute "DELETE * FROM " & DestinationTableName & " WHERE ArchiveID = " & s!ID.Value
        
        o.AddNew
            For Each sFld In s.Fields
                If IsNull(s.Fields(sFld.Name).Value) = False Then
                    For Each oFld In o.Fields
                        If sFld.Name = "ID" Then
                            o.Fields("ArchiveID").Value = sFld.Value
                        Else
                            If sFld.Name = oFld.Name Then
                                Select Case o.Fields(sFld.Name).Type
                                    Case dbComplexByte, dbComplexDecimal, dbComplexDouble, dbComplexGUID, dbComplexInteger, dbComplexLong, dbComplexSingle, dbComplexText
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
                                    Case dbAttachment
                                        Set s2 = s.Fields(sFld.Name).Value
                                        Set o2 = o.Fields(oFld.Name).Value
                                        'clear existing values from our output multivalue field
                                        Do While Not o2.EOF
                                            o2.Delete
                                            o2.MoveNext
                                        Loop
                                        ' populate each value in our source multivalue field into our output multivalue field
                                        Do While Not s2.EOF
                                            ' Clean up the URL path information from the list item's attached file (leaving just its file name and extension)
                                            Dim sSavePath As String
                                            sSavePath = s2.Fields("FileName").Value
                                            sSavePath = Right(sSavePath, Len(sSavePath) - (InStrRev(sSavePath, "/")))
                                            sSavePath = ATTACHMENTS_PATH & "\(" & s.Fields("ID").Value & ") " & sSavePath
                                            ' Delete the file at the local path, if it exists
                                            deleteExistingFile sSavePath
                                            Debug.Print Now() & "... ready to save attachment |" & sSavePath & "|"
                                            'Save the attached file locally
                                            Set sAttachments = s2.Fields("FileData")
                                            sAttachments.SaveToFile sSavePath
                                            ' Add a record to the AttachmentPaths table containing the locally saved file path of the attached file
                                            oAttachmentPaths.AddNew
                                                oAttachmentPaths!ArchiveID = s.Fields("ID").Value
                                                oAttachmentPaths!FilePath = sSavePath
                                            oAttachmentPaths.Update
                                            ' Attach the locally saved file to the archived record in [DestinationTableName]
                                            o2.AddNew
                                                o2.Fields("FileName").Value = sSavePath
                                                Set oAttachments = o2.Fields("FileData")
                                                oAttachments.LoadFromFile sSavePath
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
                        End If
                    Next
                End If
            Next
        o.Update
        iCurr = iCurr + 1
        s.MoveNext
    Loop
End Sub

Public Sub deleteExistingFile(sFilePath As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    If fso.FileExists(sFilePath) = True Then
        fso.DeleteFile sFilePath, True
    End If
    Set fso = Nothing
End Sub
