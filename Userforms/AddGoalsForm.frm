VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddGoalsForm 
   Caption         =   "UserForm2"
   ClientHeight    =   5124
   ClientLeft      =   288
   ClientTop       =   1176
   ClientWidth     =   10764
   OleObjectBlob   =   "AddGoalsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddGoalsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseButton_Click()
    AddGoalsForm.Hide
End Sub
Private Sub goal_submit_Click()
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Income&Goals")
    
    ' start on second row (headers first row)
    intRow = 4
    
    '
    If (monetary_goal_box.Value <> "" And IsNumeric(monetary_goal_box.Value)) Then
        If (goalDaytxt.Value <> "" And goalYeartxt.Value <> "" And goalmonthtxt <> "" And IsNumeric(goalDaytxt.Value) And IsNumeric(goalYeartxt.Value) And IsNumeric(goalmonthtxt.Value)) Then
            Do While (ws.Cells(intRow, "G") <> "")
                intRow = intRow + 1
            Loop
            
            ws.Cells(intRow, "G") = monetary_goal_box.Value
            ws.Cells(intRow, "H") = goalYeartxt.Value + "-" + goalmonthtxt.Value + "-" + goalDaytxt.Value
            ws.Cells(intRow, "H").NumberFormat = "yyyy-mm-dd;@"
        Else
            MsgBox "Please enter a valid numeric day, month, and year."
        End If
    Else
        MsgBox "Please enter a valid numeric monetary goal."
    End If
        
    goalDaytxt.Value = ""
    goalmonthtxt.Value = ""
    goalYeartxt.Value = ""
    monetary_goal_box.Value = ""
    AddGoalsForm.Hide
End Sub

