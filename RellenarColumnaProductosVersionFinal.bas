Sub rellenarProductosFinancierosVersionFinal()
'
'macro para crear m�ltiple informaci�n repetida con diferentes categor�a de una lista a una tabla
'

'activa la hoja del espacio de trabajo
Workbooks("4. MatrizComparativaCompilado_03julio2021.xlsx").Sheets("FormatoFinalMatriz").Activate

'verifica la cantidad de productos a duplicar
listadoProductos = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

'inicia el contador de los productos copiados
conLimSuperior = 2
conLimInferior = 3

For i = 219 To 282 'listadoProductos

    'activa la hoja del listado a copiar y copia el contenido necesario
    Workbooks("4. MatrizComparativaCompilado_03julio2021.xlsx").Sheets("FormatoFinalMatriz").Activate
    ActiveSheet.Range(Cells(i, 3), Cells(i, 22)).Copy
    conCantidadRegistros = Cells(i, 47)
    conLimInferior = conLimInferior + conCantidadRegistros - 2
    
    'activa la hoja donde se coloca la informaci�n y coloca la informaci�n necesaria
    Workbooks("4. MatrizComparativaCompilado_03julio2021.xlsx").Sheets("BD_FormatoFinal_03jul21").Activate
    ActiveSheet.Range("B" & conLimSuperior & ":U" & conLimInferior).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'activa la hoja del listado a copiar y buscar que eslabon y actividad tiene que copiar
    conOperacion = 1
    For j = 1 To 24
        Workbooks("4. MatrizComparativaCompilado_03julio2021.xlsx").Sheets("FormatoFinalMatriz").Activate
        If Cells(i, 22 + j) = "X" Then
            ActiveSheet.Range(Cells(1, 22 + j), Cells(2, 22 + j)).Copy
            Workbooks("4. MatrizComparativaCompilado_03julio2021.xlsx").Sheets("BD_FormatoFinal_03jul21").Activate
            ActiveSheet.Range("J" & conLimSuperior + conOperacion - 1).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=True
            conOperacion = conOperacion + 1
        End If
    Next j
          
    'aumenta el contador a la cantidad de operaciones
    conLimSuperior = conLimSuperior + conCantidadRegistros
    conLimInferior = conLimInferior + 2
    
Next i

End Sub

