Attribute VB_Name = "Module2"
Sub UpdateSQL()
    Dim wsSrc As Worksheet: Set wsSrc = ActiveSheet
    Dim LastRow As Long: LastRow = wsSrc.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Dim LastCol As Long: LastCol = wsSrc.UsedRange.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Dim i As Long, j As Long
    Dim strQuery As String
    Dim strOutput As String
    Dim StrHeader As String
    Dim StrFooter As String
    
    Dim StrInsertInto As String
    
   
    strQuery = ""
    strQuery = "UPDATE [" + wsSrc.Name + "]"
    StrInsertInto = strQuery
    strOutput = strQuery + " SET "
   
    
    For i = 2 To LastRow
        strQuery = ""
        
        For j = 2 To LastCol
            If IsNumeric(wsSrc.Cells(i, j)) = True Then
                If Len(wsSrc.Cells(i, j).Text) > 0 Then
                    strQuery = strQuery + "[" + wsSrc.Cells(1, j).Text + "] = " + Replace(CStr(wsSrc.Cells(i, j).Text), "'", "''") + ", "
                Else
                    strQuery = strQuery + "[" + wsSrc.Cells(1, j).Text + "] = " + "NULL" + ", "
                End If
            Else
                If Len(wsSrc.Cells(i, j).Text) > 0 Then
                    strQuery = strQuery + "[" + wsSrc.Cells(1, j).Text + "] = " + "N'" + Replace(CStr(wsSrc.Cells(i, j).Text), "'", "''") + "', "
                Else
                    strQuery = strQuery + "[" + wsSrc.Cells(1, j).Text + "] = " + "NULL" + ", "
                End If
            End If
        Next j
    
    
        If i > 2 Then
            If IsNumeric(wsSrc.Cells(i, 1)) = True Then
                strQuery = vbCr & StrInsertInto + " SET " + Left(strQuery, Len(strQuery) - 2) + " WHERE " + wsSrc.Cells(1, 1).Text + " = " + CStr(wsSrc.Cells(i, 1).Text)
            Else
                strQuery = vbCr & StrInsertInto + " SET " + Left(strQuery, Len(strQuery) - 2) + " WHERE " + wsSrc.Cells(1, 1).Text + " = N'" + CStr(wsSrc.Cells(i, 1).Text) + "'"
            End If
        Else
            If IsNumeric(wsSrc.Cells(i, 1)) = True Then
                strQuery = Left(strQuery, Len(strQuery) - 2) + " WHERE " + wsSrc.Cells(1, 1).Text + " = " + CStr(wsSrc.Cells(i, 1).Text)
            Else
               strQuery = Left(strQuery, Len(strQuery) - 2) + " WHERE " + wsSrc.Cells(1, 1).Text + " = N'" + CStr(wsSrc.Cells(i, 1).Text) + "'"
            End If
            
        End If
        strOutput = strOutput + strQuery
        
    Next i
     
    OutputForm.txtOutput.Text = strOutput & vbCr
    OutputForm.Show vbModal

End Sub
