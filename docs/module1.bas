Attribute VB_Name = "Module1"

Sub GetInsertSQL()
    Dim wsSrc As Worksheet: Set wsSrc = ActiveSheet
    Dim LastRow As Long: LastRow = wsSrc.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Dim LastCol As Long: LastCol = wsSrc.UsedRange.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Dim i As Long, j As Long
    Dim strQuery As String
    Dim strOutput As String
    Dim StrHeader As String
    Dim StrFooter As String
    
    Dim StrInsertInto As String
    
    StrHeader = "SET IDENTITY_INSERT " + wsSrc.Name + " ON"
    StrFooter = "SET IDENTITY_INSERT " + wsSrc.Name + " OFF"
    
    strQuery = ""
    For j = 1 To LastCol
        strQuery = strQuery + "[" + CStr(wsSrc.Cells(1, j)) + "], "
    Next j
   
    strQuery = Left(strQuery, Len(strQuery) - 2)
    strQuery = "INSERT INTO [" + wsSrc.Name + "] (" + strQuery + ")"
    StrInsertInto = strQuery
    strOutput = strQuery + " VALUES "
   
    
    For i = 2 To LastRow
        strQuery = ""
        
        For j = 1 To LastCol
            If IsNumeric(wsSrc.Cells(i, j)) = True Then
                If Len(wsSrc.Cells(i, j).Text) > 0 Then
                    strQuery = strQuery + Replace(CStr(wsSrc.Cells(i, j).Text), "'", "''") + ", "
                Else
                    strQuery = strQuery + "NULL" + ", "
                End If
            Else
                If Len(wsSrc.Cells(i, j).Text) > 0 Then
                    strQuery = strQuery + "N'" + Replace(CStr(wsSrc.Cells(i, j).Text), "'", "''") + "', "
                Else
                    strQuery = strQuery + "NULL" + ", "
                End If
            End If
        Next j
    
    
    If i > 2 Then
        strQuery = vbCr & StrInsertInto + " VALUES " + "(" + Left(strQuery, Len(strQuery) - 2) + ")"
    Else
        strQuery = "(" + Left(strQuery, Len(strQuery) - 2) + ")"
    End If
    strOutput = strOutput + strQuery
        
    Next i
     
    OutputForm.txtOutput.Text = StrHeader & vbCr & strOutput & vbCr & StrFooter
    OutputForm.Show vbModal
   
End Sub

