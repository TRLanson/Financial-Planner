Attribute VB_Name = "IncomeFormModule"
Sub ShowIncomeForm()
    AddIncomeForm.Show vbModeless
End Sub

Sub ClearIncomeForm()
    Dim tbl As ListObject
    
    Set tbl = ThisWorkbook.Sheets("Income&Goals").ListObjects("income_table")
    tbl.DataBodyRange.ClearContents
End Sub
