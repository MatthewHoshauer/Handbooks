'------------------------------------
'      VBA Programming Guide 
'     Author: Matthew Hoshauer
'------------------------------------

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