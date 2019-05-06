Option Explicit

Sub SetWorkbookList()
    Dim listFormula As String
    
    Dim book As Workbook
    For Each book In Workbooks
        If Len(listFormula) > 0 Then
            listFormula = listFormula & ","
        End If
        listFormula = listFormula & book.Name
    Next
    
    
    Dim targetCell As Range
    Set targetCell = Me.ActiveSheet.Range("B2")
    
    ' clear validation
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, Operator:=xlEqual, Formula1:=listFormula
    End With
    
End Sub

