Attribute VB_Name = "Module1"
'Excel macro to move bulk text from one column to a comment in another column while ignoring blank rows
'by Ryan Wissman

Sub Text_To_Comments()

Dim sText As String
Dim iStep As Long
Dim originCol As Integer
Dim destinationCol As Integer
Dim startRow As Integer

originCol = 1 'Column number for text to be copied from
destinationCol = 2 'Column number for text to be copied to
startRow = 1 'Row number to begin from, will step through rows until data ends

For iStep = startRow To ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    sText = ActiveSheet.Cells(iStep, originCol).Value
    
    'Delete any existing comment
    Cells(iStep, destinationCol).ClearComments
    
    'Check if text exists before creating comment on same line
    If Len(Trim(sText)) <> 0 Then
    
        'create comment and resize to fit all text
        Cells(iStep, destinationCol).AddComment
        Cells(iStep, destinationCol).Comment.Text Text:=sText
        Cells(iStep, destinationCol).Comment.Shape.TextFrame.AutoSize = True
        
    End If
    
Next iStep

End Sub
