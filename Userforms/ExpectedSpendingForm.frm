VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExpectedSpendingForm 
   Caption         =   "UserForm2"
   ClientHeight    =   6312
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7416
   OleObjectBlob   =   "ExpectedSpendingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExpectedSpendingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SubmitButton_Click()
    Dim ws As Worksheet
    Dim category As String
    Dim spendingAmount As Double
    Dim targetRow As Long

    Set ws = ThisWorkbook.Sheets("Expected Spending")

    ' Get the entered spending amount
    If IsNumeric(Me.exspendingtxt.Value) And Me.exspendingtxt.Value <> "" Then
        spendingAmount = CDbl(Me.exspendingtxt.Value)
    Else
        MsgBox "Please enter a valid numeric value for expected spending.", vbExclamation
        Exit Sub
    End If

    ' Determine the selected category
    If Me.OBBills.Value Then
        category = "Bills"
    ElseIf Me.OBEntertainment.Value Then
        category = "Entertainment"
    ElseIf Me.OBTuition.Value Then
        category = "Tuition"
    ElseIf Me.OBUtilities.Value Then
        category = "Utilities"
    ElseIf Me.OBRent.Value Then
        category = "Rent"
    ElseIf Me.OBGroceries.Value Then
        category = "Groceries"
    ElseIf Me.OBShopping.Value Then
        category = "Shopping"
    ElseIf Me.OBOther.Value Then
        category = "Other"
    Else
        MsgBox "Please select a category.", vbExclamation
        Exit Sub
    End If

    ' Find the row of the selected category in the table
    On Error Resume Next
    targetRow = Application.Match(category, ws.Range("A:A"), 0)
    On Error GoTo 0

    If targetRow > 0 Then
        ' Update the "Expected Spending" column in the corresponding row
        ws.Cells(targetRow, 2).Value = spendingAmount
    Else
        MsgBox "Category not found in the table.", vbCritical
    End If

    ' Clear the form for the next input
    Me.exspendingtxt.Value = ""
    Me.OBBills.Value = False
    Me.OBEntertainment.Value = False
    Me.OBTuition.Value = False
    Me.OBUtilities.Value = False
    Me.OBRent.Value = False
    Me.OBGroceries.Value = False
    Me.OBShopping.Value = False
    Me.OBOther.Value = False
End Sub

Private Sub CloseButton_Click()
    ' Close the UserForm
    Unload Me
End Sub


