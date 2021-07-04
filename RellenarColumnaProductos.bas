Attribute VB_Name = "RellenarColumnaProductos"
Sub rellenarProductosFinancieros()
'
'macro para crear múltiple información repetida con diferentes categoría de una lista a una tabla
'

'activa la hoja del espacio de trabajo
Workbooks("4.1 MatrizComparativaCompiladoNit_21junio2021.xlsx").Sheets("ENTRADA").Activate

'verifica la cantidad de productos a duplicar
listadoProductos = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

'inicia el contador de los productos copiados
conLimSuperior = 2
conLimInferior = 25

For i = 2 To listadoProductos

    'activa la hoja del listado a copiar y copia el contenido necesario
    Workbooks("4.1 MatrizComparativaCompiladoNit_21junio2021.xlsx").Sheets("ENTRADA").Activate
    ActiveSheet.Range(Cells(i, 1), Cells(i, 17)).Copy
    
    'activa la hoja donde se coloca la información y coloca la información necesaria
    Workbooks("4.1 MatrizComparativaCompiladoNit_21junio2021.xlsx").Sheets("SALIDA").Activate
    ActiveSheet.Range("A" & conLimSuperior & ":Q" & conLimInferior).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'activa la hoja del listado a copiar y copia el contenido necesario
    Workbooks("4.1 MatrizComparativaCompiladoNit_21junio2021.xlsx").Sheets("Todos").Activate
    ActiveSheet.Range(Cells(1, 1), Cells(24, 2)).Copy
    
    'activa la hoja donde se coloca la información y coloca la información necesaria
    Workbooks("4.1 MatrizComparativaCompiladoNit_21junio2021.xlsx").Sheets("SALIDA").Activate
    ActiveSheet.Range("F" & conLimSuperior).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'aumenta el contador a la cantidad de operaciones
    conLimSuperior = conLimSuperior + 24
    conLimInferior = conLimInferior + 24
    
Next i

End Sub
Sub rellenarProductosFinancierosVersionFinal()
'
'macro para crear múltiple información repetida con diferentes categoría de una lista a una tabla
'

'activa la hoja del espacio de trabajo
Workbooks("4. MatrizComparativaCompilado_03julio2021.xlsx").Sheets("FormatoFinalMatriz").Activate

'verifica la cantidad de productos a duplicar
listadoProductos = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

'inicia el contador de los productos copiados
conLimSuperior = 2
conLimInferior = 3

For i = 24 To 30 'listadoProductos

    'activa la hoja del listado a copiar y copia el contenido necesario
    Workbooks("4. MatrizComparativaCompilado_03julio2021.xlsx").Sheets("FormatoFinalMatriz").Activate
    ActiveSheet.Range(Cells(i, 3), Cells(i, 22)).Copy
    conCantidadRegistros = Cells(i, 47)
    conLimInferior = conLimInferior + conCantidadRegistros - 2
    
    'activa la hoja donde se coloca la información y coloca la información necesaria
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

