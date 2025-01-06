Attribute VB_Name = "GoalsModule"
Sub ShowGoalsForm()
    AddGoalsForm.Show vbModeless
End Sub

Sub ClearGoalTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngToClear As Range
    
    Set ws = ThisWorkbook.Sheets("Income&Goals")


    Set tbl = ws.ListObjects("goals_table")

    ' Define the range for the first two columns of the table (excluding headers)
    Set rngToClear = tbl.ListColumns(1).DataBodyRange
    Set rngToClear = Union(rngToClear, tbl.ListColumns(2).DataBodyRange)
    rngToClear.ClearContents
End Sub

