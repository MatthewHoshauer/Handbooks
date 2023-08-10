'------------------------------------
'      VBA Programming Guide 
'     Author: Matthew Hoshauer
'------------------------------------

' + ---------Commentary -----------------
' Mostly used in tandem w Excel or Word
' Simple 
'-------------------------------

' + Creating a Procedure

Sub ExampleProcedure()
    range("a2").value = 2
End Sub

' + Objects, Properties, Methods ---------------
'
' Range("A1") / Sheets(1) = Objects
' Font.Size / .Value = Properties
' .Delete = Methods
'
' Note: If you don't specify sheet, will do active sheet

Sub ExampleProcedure()
    range("a2").value = "string of text" ' You can also make values strings

    Dim Row
    Row = 5
        range("a" & row).value = 1
End Sub

' + -------- Variables ----------- +
'
'  Integer - 32,768 and 32,768
'  Long    - 2,147,483,648 and 2,147,483,648
'  Double  - Decimals
'  String  - 
'  Date    - Dates must be between signs -> #31/12/1999
'  Boolean - T/F
'  Variant - ANY variable
'  Workbook  - wkbook object
'  Worksheet - wksheet object
'  Object    - Any objects

' + --------- Further Examples --------- +

Worksheets(1).Range("A1").Borders.LineStyle = xlDouble

' + ------- Branching and Looping -------- +

Sub Macro1()
    If Worksheets(1).Range("A1").Value = "Yes!" Then
        Dim i As Integer
        For i = 2 To 10
            Worksheets(1).Range("A" & i).Value = "OK! " & i
        Next i
    Else
        MsgBox "Put Yes! in cell A1"
    End If
End Sub


' + ------- Delete Empty Rows in Excel ------ +

Sub DeleteEmptyRows()
    SelectedRange = Selection.Rows.Count
    ActiveCell.Offset(0, 0).Select
    For i = 1 To SelectedRange
        If ActiveCell.Value = "" Then
            Selection.EntireRow.Delete
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next i
End Sub

' + Delete Empty Text Boxes in Powerpoint

Sub RemoveEmptyTextBoxes()
    Dim SlideObj As Slide
    Dim ShapeObj As Shape
    Dim ShapeIndex As Integer
    For Each SlideObj In ActivePresentation.Slides
        For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1
            Set ShapeObj = SlideObj.Shapes(ShapeIndex)
            If ShapeObj.Type = msoTextBox Then
                If Trim(ShapeObj.TextFrame.TextRange.Text) = "" Then
                    ShapeObj.Delete
                End If
            End If
        Next ShapeIndex
    Next SlideObj
End Sub