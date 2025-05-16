Attribute VB_Name = "SplitTextToColumns"
' SplitTextToColumns.bas
' -----------------------------
' Module Name: SplitTextToColumns
' Description: Splits text in cells from column A into individual characters across columns
' Author: Jeffrey Bulanadi
' Created: May 11, 2025
' Last Modified: May 11, 2025
' -----------------------------
' Usage: 
'  1. Select cells in column A containing text to split
'  2. Run this macro
'  3. Text will be split into individual characters across columns
' -----------------------------

Option Explicit

Public Sub SplitTextToColumns()
    ' Declare variables
    Dim cell As Range
    Dim i As Integer
    Dim text As String
    Dim ws As Worksheet
    
    ' Set active worksheet
    Set ws = ActiveSheet
    
    ' Loop through each cell in specified range
    For Each cell In ws.Range("A1:A10") ' Adjust the range as needed
        ' Get cell value
        text = cell.Value
        
        ' Process only if cell has content
        If Len(text) > 0 Then
            ' Split text into individual characters
            For i = 1 To Len(text)
                cell.Offset(0, i - 1).Value = Mid(text, i, 1)
            Next i
        End If
    Next cell
    
    ' Notify user when complete
    MsgBox "Text has been split into columns successfully!", vbInformation, "Split Text Complete"
End Sub