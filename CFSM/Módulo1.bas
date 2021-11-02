Attribute VB_Name = "Módulo1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Next.Select
    Range("B3:B51").Select
    Selection.Copy
    Sheets("Report").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("B3").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Columns("B:B").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
