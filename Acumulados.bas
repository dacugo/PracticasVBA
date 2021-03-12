Attribute VB_Name = "Acumulados"
Sub ValoresAcumulados()

Dim departamento1, municipio1, nombreTecnico1, familia1 As String
Dim departamento2, municipio2, nombreTecnico2, familia2 As String
Dim anno, contador As Integer
Dim areaAcumulada As Double
Dim numeroRegistros As Long

numeroRegistros = Cells(Rows.Count, "A").End(xlUp).Row
i = 2
departamento1 = Cells(i, 4)
municipio1 = Cells(i, 5)
nombreTecnico1 = Cells(i, 6)
familia1 = Cells(i, 7)
anno1 = Cells(i, 9)
contador = 0

For i = 3 To 38 'numeroRegistros

    departamento2 = Cells(i, 4)
    municipio2 = Cells(i, 5)
    nombreTecnico2 = Cells(i, 6)
    familia2 = Cells(i, 7)
    anno2 = Cells(i, 9)

    If departamento1 = departamento2 And municipio1 = municipio2 And nombreTecnico1 = nombreTecnico2 Then
        If contador = 0 Then
            Cells(i, 11) = Cells(i, 10)
            areaAcumulada = Cells(i - 1, 10)
            contador = contador + 1
        Else
            areaAcumulada = areaAcumulada + Cells(i - 1, 10)
            Cells(i - 1, 11) = areaAcumulada
        End If
    Else
        If contador <> 0 Then
            areaAcumulada = areaAcumulada + Cells(i - 1, 10)
            Cells(i - 1, 11) = areaAcumulada
            Cells(i - 1, 11).Interior.Color = 5296274
            Cells(i, 11) = Cells(i, 10)
            departamento1 = Cells(i, 4)
            municipio1 = Cells(i, 5)
            nombreTecnico1 = Cells(i, 6)
            familia1 = Cells(i, 7)
            anno1 = Cells(i, 9)
            contador = 0
            areaAcumulada = 0
        Else
            Cells(i, 11) = Cells(i, 10)
            Cells(i - 1, 11).Interior.Color = 5296274
            departamento1 = Cells(i, 4)
            municipio1 = Cells(i, 5)
            nombreTecnico1 = Cells(i, 6)
            familia1 = Cells(i, 7)
            anno1 = Cells(i, 9)
            contador = 0
            areaAcumulada = 0
        End If
    End If
   
Next i

End Sub
