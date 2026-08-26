Attribute VB_Name = "modCustomUtility"
Public Function FindHeaderColumn( _
        ByVal ws As Worksheet, _
        ByVal headerRow As Long, _
        ByVal headerName As String) As Long

    Dim found As Range

    Set found = ws.Rows(headerRow).Find( _
                    What:=headerName, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    MatchCase:=False)

    If found Is Nothing Then

        FindHeaderColumn = 0

    Else

        FindHeaderColumn = found.Column

    End If

End Function

Public Function FindHeaderRow( _
        ByVal ws As Worksheet, _
        ByVal headerColumn As String, _
        ByVal headerName As String) As Long

    Dim found As Range


    Set found = ws.Range(headerColumn).Find(What:=headerName, _
                                        LookIn:=xlValues, _
                                        LookAt:=xlWhole)
    
    If found Is Nothing Then

        FindHeaderRow = 0

    Else

        FindHeaderRow = found.Row

    End If

End Function

