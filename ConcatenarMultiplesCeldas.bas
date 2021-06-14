Attribute VB_Name = "M�dulo3"
Sub ConcatenarMultiplesCeldas()
Attribute ConcatenarMultiplesCeldas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Une el contenido de un rango de celdas seleccionado, borra el contenido extra y coloca todo el contenido en la primera celda selccionada
'

'definici�n de variables
    Dim rangeText As String
    Dim rangeCount, numberColumn, numberRow As Integer
    
'inicializaci�n de la variables
    rangeText = ""
    numberColumn = ActiveCell.Column
    numberRow = ActiveCell.Row
    rangeCount = numberRow
    
'ciclo que lee el contenido de las celdas y lo une
    For Each RangeCells In Selection
        rangeText = rangeText + RangeCells.Value + Chr(10)
        rangeCount = rangeCount + 1
    Next RangeCells
    
'remover el �ltimo caracter de la cadena de texto y colocar el texto
    rangeText = Left(rangeText, Len(rangeText) - 1)
    Cells(numberRow, numberColumn) = rangeText

'se borra el contenido de las celdas
    numberRow = numberRow + 1
    ActiveSheet.Range(Cells(numberRow, numberColumn), Cells(rangeCount - 1, numberColumn)).Select
    Selection.ClearContents

End Sub
