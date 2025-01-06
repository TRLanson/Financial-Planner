VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddIncomeForm 
   Caption         =   "UserForm3"
   ClientHeight    =   5670
   ClientLeft      =   156
   ClientTop       =   588
   ClientWidth     =   10524
   OleObjectBlob   =   "AddIncomeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddIncomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub incomeClose_Click()
    AddIncomeForm.Hide
End Sub
Private Sub UserForm_Initialize()
    With AddIncomeForm.cboxSource
        .Clear
        .AddItem "Main Salary"
        .AddItem "Side Salary 1"
        .AddItem "Side Salary 2"
        .AddItem "Academics"
    End With
    With AddIncomeForm.cboxCategory
        .Clear
        .AddItem "Work"
        .AddItem "Scholarship"
        .AddItem "OSAP"
        .AddItem "Grant"
        .AddItem "Bursary"
    End With
End Sub

Private Sub incomeSubmit_Click()
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Income&Goals")
    
    ' start on fourth row (headers + title)
    intRow = 4
    
    If (cboxSource <> "") Then
        If (incomeDaytxt <> "" And incomeMonthtxt <> "" And incomeYeartxt <> "" And IsNumeric(incomeDaytxt.Value) And IsNumeric(incomeMonthtxt.Value) And IsNumeric(incomeYeartxt.Value)) Then
            If (cboxCategory <> "") Then
                If (descriptiontxt <> "" And IsNumeric(descriptiontxt)) Then
                    Do While (ws.Cells(intRow, "A") <> "")
                        intRow = intRow + 1
                        
                    Loop
                    
                    'Write date into cell
                    ws.Cells(intRow, "A") = incomeYeartxt.Value + "-" + incomeMonthtxt.Value + "-" + incomeDaytxt.Value
            
                    'Format cell so Excel recognizes a date
                    ws.Cells(intRow, "A").NumberFormat = "yyyy-mm-dd;@"
                    
                    'Write item into cell
                    ws.Cells(intRow, "B") = cboxSource.Value
                    
                    'Write category into cell
                    ws.Cells(intRow, "C") = cboxCategory.Value
                    
                    'Write description into cell
                    ws.Cells(intRow, "D") = descriptiontxt.Value
                
                Else
                    MsgBox "Please enter a valid numeric description of the income."
                End If
            Else
                MsgBox "Please choose an income category."
            End If
        Else
            MsgBox "Please enter a valid numeric value for the income day, month, and year."
        End If
    Else
        MsgBox "Please choose a source of the income."
    End If
    
    incomeDaytxt = ""
    incomeMonthtxt = ""
    incomeYeartxt = ""
    cboxSource = ""
    cboxCategory = ""
    descriptiontxt = ""
    AddIncomeForm.Hide
    
End Sub

