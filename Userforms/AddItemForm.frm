VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddItemForm 
   Caption         =   "Add Expense/Income"
   ClientHeight    =   5784
   ClientLeft      =   180
   ClientTop       =   708
   ClientWidth     =   11496
   OleObjectBlob   =   "AddItemForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label5_Click()

End Sub

Private Sub UserForm_Initialize()
    'Set up drop-down menu for form
    With AddItemForm.cboxCategory
        .AddItem "Shopping"
        .AddItem "Bills"
        .AddItem "Groceries"
        .AddItem "Entertainment"
        .AddItem "Tuition"
        .AddItem "Rent"
        .AddItem "Utilities"
        .AddItem "Other"
    End With
End Sub

Private Sub SubmitBtn_Click()

    'Set workbook and sheet
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Expenses")

    'start on second row (headers are first row)
    intRow = 4
    
    If (txtDescription.Value <> "") And IsNumeric(txtDescription.Value) Then
    'Test value of Item textbox
        If (txtItem.Value <> "") Then
        
            'Test value of date textboxes
            If (txtDay.Value <> "" And txtMonth.Value <> "" And txtYear.Value <> "") Then
            
                'Test value of Category combobox
                If (cboxCategory.Value <> "") Then
        
                    'Go through rows, if they contain data, increment
                    Do While (ws.Cells(intRow, "A") <> "")
                    
                        'Increment row counter
                        intRow = intRow + 1
                    
                    Loop
                                    
                    'Write date into cell
                    ws.Cells(intRow, "A") = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
            
                    'Format cell so Excel recognizes a date
                    ws.Cells(intRow, "A").NumberFormat = "yyyy-mm-dd;@"
                    
                    'Write item into cell
                    ws.Cells(intRow, "B") = txtItem.Value
                    
                    'Write category into cell
                    ws.Cells(intRow, "C") = cboxCategory.Value
                    
                    'Write description into cell
                    ws.Cells(intRow, "D") = txtDescription.Value
                
                Else
                    'Give error for no category
                    MsgBox ("Please select a category")
                End If
            
            Else
                'Give error message for no date
                MsgBox ("Please enter a valid date")
            End If
            
        Else
            'Give error message for no item
            MsgBox ("Please enter an item")
        End If
    
    Else
        MsgBox ("Please enter a valid numerical description (price)")
    End If
    
    ' Empty all boxes and close the form
    txtItem.Value = ""
    txtDay.Value = ""
    txtYear.Value = ""
    txtMonth.Value = ""
    cboxCategory.Value = ""
    txtDescription.Value = ""
    AddItemForm.Hide
    
End Sub

