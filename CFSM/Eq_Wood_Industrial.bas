Attribute VB_Name = "Eq_Wood_Industrial"
Sub SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f, g As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
        
    hojUsu_Summary.Cells(k, 56).Formula = "=(((Wood_Industrial!K" & i & "*Wood_Industrial!L" & i & _
    "*(1-(Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!M" & i & "*Wood_Industrial!N" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!O" & i & "*Wood_Industrial!P" & i & ")-((Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & _
        ")*(Wood_Industrial!O" & j & "*Wood_Industrial!P" & j & "))))" & Chr(10) & _
    "+((Wood_Industrial!Q" & i & "*Wood_Industrial!R" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!S" & i & "*Wood_Industrial!T" & i & ")-((Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & _
        ")*(Wood_Industrial!S" & j & "*Wood_Industrial!T" & j & ")))))" & Chr(10) & _
    "*Wood_Industrial!B" & i & _
    ")+(Summary!BD" & l & "*(Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & ")*(Wood_Industrial!W" & j & "*Wood_Industrial!X" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 56).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 56) = hojUsu_Forecast.Cells(c_ini, 87)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 56).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 56) = 0
        
        End If
        
    Case 3
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 88) = hojUsu_Summary.Cells(k, 56).Value
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 56).Formula = "=(((Wood_Industrial!K" & i & "*Wood_Industrial!L" & i & _
    "*(1-(Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!M" & i & "*Wood_Industrial!N" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!O" & i & "*Wood_Industrial!P" & i & ")-((Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & _
        ")*(Wood_Industrial!O" & j & "*Wood_Industrial!P" & j & "))))" & Chr(10) & _
    "+((Wood_Industrial!Q" & i & "*Wood_Industrial!R" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!S" & i & "*Summary!BT" & k & ")-((Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & _
        ")*(Wood_Industrial!S" & j & "*Summary!BT" & l & ")))))" & Chr(10) & _
    "*Wood_Industrial!B" & i & _
    ")+(Summary!BD" & l & "*(Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!U" & i & "*Wood_Industrial!V" & i & ")*(Wood_Industrial!W" & j & "*Wood_Industrial!X" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 56).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 56) = hojUsu_Forecast.Cells(c_ini, 87)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 56).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 56) = 0
        
        End If
        
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 89) = hojUsu_Summary.Cells(k, 56).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(6, 21).Value = "=Summary!BD" & k - 1 & "+Summary!BF" & k - 1

End Sub
Sub SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f, g As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
        
    hojUsu_Summary.Cells(k, 58).Formula = "=(((Wood_Industrial!Z" & i & "*Wood_Industrial!AA" & i & _
    "*(1-(Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!AB" & i & "*Wood_Industrial!AC" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!AD" & i & "*Wood_Industrial!AE" & i & ")-((Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & _
        ")*(Wood_Industrial!AD" & j & "*Wood_Industrial!AE" & j & ")))))" & Chr(10) & _
    "*Wood_Industrial!C" & i & _
    ")+(Summary!BF" & l & "*(Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & ")*(Wood_Industrial!AH" & j & "*Wood_Industrial!AI" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 58).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 58) = hojUsu_Forecast.Cells(c_ini, 91)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 58).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 58) = 0
        
        End If
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 92) = hojUsu_Summary.Cells(k, 58).Value
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 58).Formula = "=(((Wood_Industrial!Z" & i & "*Wood_Industrial!AA" & i & _
    "*(1-(Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!AB" & i & "*Wood_Industrial!AC" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!AD" & i & "*Summary!BV" & k & ")-((Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & _
        ")*(Wood_Industrial!AD" & j & "*Summary!BV" & l & ")))))" & Chr(10) & _
    "*Wood_Industrial!C" & i & _
    ")+(Summary!BF" & l & "*(Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!AF" & i & "*Wood_Industrial!AG" & i & ")*(Wood_Industrial!AH" & j & "*Wood_Industrial!AI" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 58).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 58) = hojUsu_Forecast.Cells(c_ini, 91)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 58).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 58) = 0
        
        End If
        
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 93) = hojUsu_Summary.Cells(k, 58).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(6, 21).Value = "=Summary!BD" & k - 1 & "+Summary!BF" & k - 1

End Sub
Sub CONSUMPTION_WOOD_INDUSTRIAL()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
       
    hojUsu_Summary.Cells(k, 60).Formula = "=(((Wood_Industrial!AK" & i & "*Wood_Industrial!AL" & i & _
    "*(1-(Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!AM" & i & "*Wood_Industrial!AN" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!AO" & i & "*Wood_Industrial!AP" & i & ")-((Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & _
        ")*(Wood_Industrial!AO" & j & "*Wood_Industrial!AP" & j & "))))" & Chr(10) & _
    "+((Wood_Industrial!AQ" & i & "*Wood_Industrial!AR" & i & ")" & Chr(10) & _
    "*(((Wood_Industrial!AS" & i & "* Wood_Industrial!AT" & i & ")/(Wood_Industrial!AU" & i & "* Wood_Industrial!AV" & i & _
        "))-(( Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & _
        ")*( (Wood_Industrial!AS" & j & "* Wood_Industrial!AT" & j & ")/(Wood_Industrial!AU" & j & "* Wood_Industrial!AV" & j & "))))))" & Chr(10) & _
    "*Wood_Industrial!C" & i & _
    ")+(Summary!BH" & l & "*(Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & ")*(Wood_Industrial!AY" & j & "*Wood_Industrial!AZ" & j & "))"
            
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 60).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 60) = hojUsu_Forecast.Cells(c_ini, 95)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 60).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 60) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 96) = hojUsu_Summary.Cells(k, 60).Value
    c_ini = c_ini + 1
        
Next k
        
Case 2, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 60).Formula = "=(((Wood_Industrial!AK" & i & "*Wood_Industrial!AL" & i & _
    "*(1-(Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!AM" & i & "*Wood_Industrial!AN" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!AO" & i & "*Wood_Industrial!AP" & i & ")-((Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & _
        ")*(Wood_Industrial!AO" & j & "*Wood_Industrial!AP" & j & "))))" & Chr(10) & _
    "+((Wood_Industrial!AQ" & i & "*Wood_Industrial!AR" & i & ")" & Chr(10) & _
    "*(((Wood_Industrial!AS" & i & "*Summary!BN" & k & ")/(Wood_Industrial!AU" & i & "* Wood_Industrial!AV" & i & _
        "))-(( Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & _
        ")*( (Wood_Industrial!AS" & j & "*Summary!BN" & l & ")/(Wood_Industrial!AU" & j & "* Wood_Industrial!AV" & j & "))))))" & Chr(10) & _
    "*Wood_Industrial!C" & i & _
    ")+(Summary!BH" & l & "*(Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & ")*(Wood_Industrial!AY" & j & "*Wood_Industrial!AZ" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 60).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 60) = hojUsu_Forecast.Cells(c_ini, 95)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 60).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 60) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 97) = hojUsu_Summary.Cells(k, 60).Value
    c_ini = c_ini + 1
    
Next k

'en caso de usar "Module (Industrys - NPW)"
Case 4

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 60).Formula = "=(((Wood_Industrial!AK" & i & "*Wood_Industrial!AL" & i & _
    "*(1-(Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!AM" & i & "*Wood_Industrial!AN" & i & ")" & Chr(10) & _
        "*((Wood_Industrial!AO" & i & "*MWMconnectedUWM!F" & i & ")-((Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & _
        ")*(Wood_Industrial!AO" & j & "*MWMconnectedUWM!F" & j & "))))" & Chr(10) & _
    "+((Wood_Industrial!AQ" & i & "*Wood_Industrial!AR" & i & ")" & Chr(10) & _
    "*(((Wood_Industrial!AS" & i & "*Summary!BN" & k & ")/(Wood_Industrial!AU" & i & "* Wood_Industrial!AV" & i & _
        "))-(( Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & _
        ")*( (Wood_Industrial!AS" & j & "*Summary!BN" & l & ")/(Wood_Industrial!AU" & j & "* Wood_Industrial!AV" & j & "))))))" & Chr(10) & _
    "*Wood_Industrial!C" & i & _
    ")+(Summary!BH" & l & "*(Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!AW" & i & "*Wood_Industrial!AX" & i & ")*(Wood_Industrial!AY" & j & "*Wood_Industrial!AZ" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 60).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 60) = hojUsu_Forecast.Cells(c_ini, 95)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 60).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 60) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 97) = hojUsu_Summary.Cells(k, 60).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(7, 21).Value = "=Summary!BH" & k - 1

End Sub
Sub EXPORTS_WOOD_INDUSTRIAL()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
        
    hojUsu_Summary.Cells(k, 62).Formula = _
        "=(((Wood_Industrial!BE" & i & "*Wood_Industrial!BF" & i & _
            "*(1-(Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!BG" & i & "*Wood_Industrial!BH" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!BI" & i & "* Wood_Industrial!BJ" & i & _
            ")-(( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & _
            ")*(Wood_Industrial!BI" & j & "* Wood_Industrial!BJ" & j & "))))" & Chr(10) & _
        "+((Wood_Industrial!BK" & i & "*Wood_Industrial!BL" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!BM" & i & "* Wood_Industrial!BN" & i & _
            ")/(Wood_Industrial!BO" & i & "* Wood_Industrial!BP" & i & _
            "))-(( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & _
            ")*( (Wood_Industrial!BM" & j & "* Wood_Industrial!BN" & j & _
            ")/(Wood_Industrial!BO" & j & "* Wood_Industrial!BP" & j & ")))))" & Chr(10) & _
        "*Wood_Industrial!E" & i & ")" & Chr(10) & _
        "+(Summary!BJ" & l & Chr(10) & _
        "*( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!BS" & j & "*Wood_Industrial!BT" & j & ")*( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 62).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 62) = hojUsu_Forecast.Cells(c_ini, 99)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 62).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 62) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 100) = hojUsu_Summary.Cells(k, 62).Value
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 62).Formula = _
        "=(((Wood_Industrial!BE" & i & "*Wood_Industrial!BF" & i & _
            "*(1-(Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!BG" & i & "*Wood_Industrial!BH" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!BI" & i & "* Wood_Industrial!BJ" & i & _
            ")-(( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & _
            ")*(Wood_Industrial!BI" & j & "* Wood_Industrial!BJ" & j & "))))" & Chr(10) & _
        "+((Wood_Industrial!BK" & i & "*Wood_Industrial!BL" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!BM" & i & "*Summary!BP" & k & _
            ")/(Wood_Industrial!BO" & i & "* Wood_Industrial!BP" & i & _
            "))-(( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & _
            ")*( (Wood_Industrial!BM" & j & "*Summary!BP" & l & _
            ")/(Wood_Industrial!BO" & j & "* Wood_Industrial!BP" & j & ")))))" & Chr(10) & _
        "*Wood_Industrial!E" & i & ")" & Chr(10) & _
        "+(Summary!BJ" & l & Chr(10) & _
        "*( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!BS" & j & "*Wood_Industrial!BT" & j & ")*( Wood_Industrial!BQ" & i & "*Wood_Industrial!BR" & i & ")))"
               
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 62).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 62) = hojUsu_Forecast.Cells(c_ini, 99)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 62).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 62) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 101) = hojUsu_Summary.Cells(k, 62).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(8, 21).Value = "=Summary!BJ" & k - 1

End Sub
Sub IMPORTS_WOOD_INDUSTRIAL()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
        
    hojUsu_Summary.Cells(k, 64).Formula = _
        "=(((Wood_Industrial!BV" & i & "*Wood_Industrial!BW" & i & _
            "*(1-(Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!BX" & i & "*Wood_Industrial!BY" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!BZ" & i & "* Wood_Industrial!CA" & i & _
            ")-((Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & _
            ")*(Wood_Industrial!BZ" & j & "* Wood_Industrial!CA" & j & "))))" & Chr(10) & _
        "+((Wood_Industrial!CB" & i & "*Wood_Industrial!CC" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!CD" & i & "* Wood_Industrial!CE" & i & _
            ")/(Wood_Industrial!CF" & i & "* Wood_Industrial!CG" & i & _
            "))-((Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & _
            ")*( (Wood_Industrial!CD" & j & "* Wood_Industrial!CE" & j & _
            ")/(Wood_Industrial!CF" & j & "* Wood_Industrial!CG" & j & ")))))" & Chr(10) & _
        "+((Wood_Industrial!CH" & i & "*Wood_Industrial!CI" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!CJ" & i & "* Wood_Industrial!CK" & i & _
            ")-((Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & _
            ")*(Wood_Industrial!CJ" & j & "* Wood_Industrial!CK" & j & "))))" & Chr(10) & _
        "*Wood_Industrial!F" & i & ")" & Chr(10) & _
        "+(Summary!BL" & l & "*(Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!CN" & j & "*Wood_Industrial!CO" & j & ")*( Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 64).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 64) = hojUsu_Forecast.Cells(c_ini, 103)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 64).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 64) = 0
        
        End If
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 104) = hojUsu_Summary.Cells(k, 64).Value
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 64).Formula = _
        "=(((Wood_Industrial!BV" & i & "*Wood_Industrial!BW" & i & _
            "*(1-(Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!BX" & i & "*Wood_Industrial!BY" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!BZ" & i & "* Summary!BH" & k & _
            ")-((Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & _
            ")*(Wood_Industrial!BZ" & j & "* Summary!BH" & l & "))))" & Chr(10) & _
        "+((Wood_Industrial!CB" & i & "*Wood_Industrial!CC" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!CD" & i & "*Summary!BN" & k & _
            ")/(Wood_Industrial!CF" & i & "*Summary!BR" & k & _
            "))-((Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & _
            ")*( (Wood_Industrial!CD" & j & "*Summary!BN" & l & _
            ")/(Wood_Industrial!CF" & j & "*Summary!BR" & l & ")))))" & Chr(10) & _
        "+((Wood_Industrial!CH" & i & "*Wood_Industrial!CI" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!CJ" & i & "* Wood_Industrial!CK" & i & _
            ")-((Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & _
            ")*(Wood_Industrial!CJ" & j & "* Wood_Industrial!CK" & j & "))))" & Chr(10) & _
        "*Wood_Industrial!F" & i & ")" & Chr(10) & _
        "+(Summary!BL" & l & "*(Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!CN" & j & "*Wood_Industrial!CO" & j & ")*( Wood_Industrial!CL" & i & "*Wood_Industrial!CM" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 64).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 64) = hojUsu_Forecast.Cells(c_ini, 103)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 64).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 64) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 105) = hojUsu_Summary.Cells(k, 64).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(9, 21).Value = "=Summary!BL" & k - 1

End Sub
Sub PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
        
    hojUsu_WoodIndustrial.Cells(i, 118).Formula = _
        "=(((Wood_Industrial!CQ" & i & "*Wood_Industrial!CR" & i & _
            "*(1-(Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!CS" & i & "*Wood_Industrial!CT" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!CU" & i & "*Wood_Industrial!CV" & i & _
            ")/(Wood_Industrial!CW" & i & "*Wood_Industrial!CX" & i & _
            "))-((Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & _
            ")*((Wood_Industrial!CU" & j & "*Wood_Industrial!CV" & j & _
            ")/(Wood_Industrial!CW" & j & "*Wood_Industrial!CX" & j & ")))))" & Chr(10) & _
        "+(Wood_Industrial!CY" & i & "*Wood_Industrial!CZ" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!DA" & i & "*Wood_Industrial!DB" & i & _
            ")/((Wood_Industrial!DC" & i & "*Wood_Industrial!DD" & i & _
            ")*(Wood_Industrial!DE" & i & "*Wood_Industrial!DF" & i & _
            ")))-((Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & _
            ")*((Wood_Industrial!DA" & j & "*Wood_Industrial!DB" & j & _
            ")/((Wood_Industrial!DC" & j & "*Wood_Industrial!DD" & j & _
            ")*(Wood_Industrial!DE" & j & "*Wood_Industrial!DF" & j & "))))))" & Chr(10) & _
        "*Wood_Industrial!G" & i & ")" & Chr(10) & _
        "+(Wood_Industrial!DN" & j & "*(Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!DI" & j & "*Wood_Industrial!DJ" & j & ")*(Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & "))"


    hojUsu_Summary.Cells(k, 66).Formula = _
    "=(Wood_Industrial!DK" & i & "*Wood_Industrial!DL" & i & ")*(Wood_Industrial!DM" & i & "*Wood_Industrial!DN" & i & ")"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 66).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 66) = hojUsu_Forecast.Cells(c_ini, 107)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 66).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 66) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 108) = hojUsu_Summary.Cells(k, 66).Value
    hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 66)
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_WoodIndustrial.Cells(i, 118).Formula = _
        "=(((Wood_Industrial!CQ" & i & "*Wood_Industrial!CR" & i & _
        "*(1-(Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & ")))" & Chr(10) & _
    "+((Wood_Industrial!CS" & i & "*Wood_Industrial!CT" & i & ")" & Chr(10) & _
        "*(((Wood_Industrial!CU" & i & "*(Summary!BD" & k & "+Summary!BF" & k & _
        "))/(Wood_Industrial!CW" & i & "*Wood_Industrial!CX" & i & _
        "))-((Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & _
        ")*((Wood_Industrial!CU" & j & "*(Summary!BD" & l & "+Summary!BF" & l & _
        "))/(Wood_Industrial!CW" & j & "*Wood_Industrial!CX" & j & ")))))" & Chr(10) & _
    "+(Wood_Industrial!CY" & i & "*Wood_Industrial!CZ" & i & ")" & Chr(10) & _
        "*(((Wood_Industrial!DA" & i & "*Summary!BN" & l & _
        ")/((Wood_Industrial!DC" & i & "*Wood_Industrial!DD" & i & _
        ")*(Wood_Industrial!DE" & i & "*Wood_Industrial!DF" & i & _
        ")))-((Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & _
        ")*((Wood_Industrial!DA" & j & "*Summary!BN" & m & _
        ")/((Wood_Industrial!DC" & j & "*Wood_Industrial!DD" & j & _
        ")*(Wood_Industrial!DE" & j & "*Wood_Industrial!DF" & j & "))))))" & Chr(10) & _
    "*Wood_Industrial!G" & i & ")" & Chr(10) & _
    "+(Wood_Industrial!DN" & j & "*(Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & "))" & Chr(10) & _
    "+((Wood_Industrial!DI" & j & "*Wood_Industrial!DJ" & j & ")*(Wood_Industrial!DG" & i & "*Wood_Industrial!DH" & i & "))"

    hojUsu_Summary.Cells(k, 66).Formula = _
    "=(Wood_Industrial!DK" & i & "*Wood_Industrial!DL" & i & ")*(Wood_Industrial!DM" & i & "*Wood_Industrial!DN" & i & ")"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 66).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 66) = hojUsu_Forecast.Cells(c_ini, 107)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 66).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 66) = 0
        
        End If
        
    Case 3
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 109) = hojUsu_Summary.Cells(k, 66).Value
    hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 66)
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(10, 21).Value = "=Summary!BN" & k - 1

End Sub
Sub PRICE_OF_EXPORTS_WOOD_INDUSTRIAL()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
        
    hojUsu_Summary.Cells(k, 68).Formula = _
        "=(((Wood_Industrial!DP" & i & "*Wood_Industrial!DQ" & i & _
            "*(1-(Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!DR" & i & "*Wood_Industrial!DS" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!DT" & i & "* Wood_Industrial!DU" & i & _
            ")-(( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & _
            ")*(Wood_Industrial!DT" & j & "* Wood_Industrial!DU" & j & "))))" & Chr(10) & _
        "+((Wood_Industrial!DV" & i & "*Wood_Industrial!DW" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!DX" & i & "* Wood_Industrial!DY" & i & _
            ")-(( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & _
            ")*(Wood_Industrial!DX" & j & "* Wood_Industrial!DY" & j & "))))" & Chr(10) & _
        "+((Wood_Industrial!DZ" & i & "*Wood_Industrial!EA" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!EB" & i & "* Wood_Industrial!EC" & i & _
            ")-(( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & _
            ")*(Wood_Industrial!EB" & j & "* Wood_Industrial!EC" & j & "))))" & Chr(10) & _
        "*Wood_Industrial!H" & i & ")" & Chr(10) & _
        "+(Summary!BP" & l & "*( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!EF" & j & "*Wood_Industrial!EG" & j & ")*( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 68).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 68) = hojUsu_Forecast.Cells(c_ini, 111)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 68).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 68) = 0
        
        End If
    
    End Select

    hojUsu_Forecast.Cells(c_ini, 112) = hojUsu_Summary.Cells(k, 68).Value
    hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 68)
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 68).Formula = _
        "=(((Wood_Industrial!DP" & i & "*Wood_Industrial!DQ" & i & _
            "*(1-(Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!DR" & i & "*Wood_Industrial!DS" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!DT" & i & "* Summary!BP" & l & _
            ")-(( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & _
            ")*(Wood_Industrial!DT" & j & "* Summary!BP" & m & "))))" & Chr(10) & _
        "+((Wood_Industrial!DV" & i & "*Wood_Industrial!DW" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!DX" & i & "* Wood_Industrial!DY" & i & _
            ")-(( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & _
            ")*(Wood_Industrial!DX" & j & "* Wood_Industrial!DY" & j & "))))" & Chr(10) & _
        "+((Wood_Industrial!DZ" & i & "*Wood_Industrial!EA" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!EB" & i & "* Wood_Industrial!EC" & i & _
            ")-(( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & _
            ")*(Wood_Industrial!EB" & j & "* Wood_Industrial!EC" & j & "))))" & Chr(10) & _
        "*Wood_Industrial!H" & i & ")" & Chr(10) & _
        "+(Summary!BP" & l & "*( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!EF" & j & "*Wood_Industrial!EG" & j & ")*( Wood_Industrial!ED" & i & "*Wood_Industrial!EE" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 68).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 68) = hojUsu_Forecast.Cells(c_ini, 111)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 68).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 68) = 0
        
        End If
    
    Case 3
    
    End Select
                       
    hojUsu_Forecast.Cells(c_ini, 113) = hojUsu_Summary.Cells(k, 68).Value
    hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 68)
    c_ini = c_ini + 1

   
Next k

End Select

hojUsu_Summary.Cells(11, 21).Value = "=Summary!BP" & k - 1

End Sub
Sub PRICE_OF_IMPORT_WOOD_INDUSTRIAL()

'Cambios según cantidad de años

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de proceso

'f permite seleccionar si se desea ver la validación del sistema o el MCC
f = hojUsu_SystemOptions.Range("SelectProcess").Value
'g permite escoger el uso de los datos que resultan negativos
g = hojUsu_SystemOptions.Range("NegativeData").Value

Select Case f

Case 1

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual
        
    hojUsu_Summary.Cells(k, 70).Formula = _
        "=(((Wood_Industrial!EI" & i & "*Wood_Industrial!EJ" & i & _
            "*(1-(Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!EK" & i & "*Wood_Industrial!EL" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!EM" & i & "* Wood_Industrial!EN" & i & _
            ")-(( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & _
            ")*(Wood_Industrial!EM" & j & "* Wood_Industrial!EN" & j & "))))" & Chr(10) & _
        "+((Wood_Industrial!EO" & i & "*Wood_Industrial!EP" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!EQ" & i & "*Wood_Industrial!ER" & i & _
            ")/(Wood_Industrial!ES" & i & "*Wood_Industrial!ET" & i & _
            "))-((Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & _
            ")*((Wood_Industrial!EQ" & j & "*Wood_Industrial!ER" & j & _
            ")/(Wood_Industrial!ES" & j & "*Wood_Industrial!ET" & j & ")))))" & Chr(10) & _
        "+((Wood_Industrial!EU" & i & "*Wood_Industrial!EV" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!EW" & i & "* Wood_Industrial!EX" & i & _
            ")-(( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & _
            ")*(Wood_Industrial!EW" & j & "* Wood_Industrial!EX" & j & "))))" & Chr(10) & _
        "*Wood_Industrial!I" & i & ")" & Chr(10) & _
        "+(Summary!BR" & l & "*( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!FA" & j & "*Wood_Industrial!FB" & j & ")*( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 70).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 70) = hojUsu_Forecast.Cells(c_ini, 115)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 70).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 70) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 116) = hojUsu_Summary.Cells(k, 70).Value
    hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 70)
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4, 5

For k = d_ini To d_fin
        
'i es el número 8 de la hoja del mercado
i = k - 31
'j es el número 7 de la hoja del mercado
j = i - 1
'l es el n anterior de la hoja en summary
l = k - 1
'k es el año actual de la hoja summary
m = l - 1
'm es el n dos años anteriores al año actual

    hojUsu_Summary.Cells(k, 70).Formula = _
        "=(((Wood_Industrial!EI" & i & "*Wood_Industrial!EJ" & i & _
            "*(1-(Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & ")))" & Chr(10) & _
        "+((Wood_Industrial!EK" & i & "*Wood_Industrial!EL" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!EM" & i & "* Summary!BR" & l & _
            ")-(( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & _
            ")*(Wood_Industrial!EM" & j & "* Summary!BR" & m & "))))" & Chr(10) & _
        "+((Wood_Industrial!EO" & i & "*Wood_Industrial!EP" & i & ")" & Chr(10) & _
            "*(((Wood_Industrial!EQ" & i & "*Summary!BN" & k & _
            ")/(Wood_Industrial!ES" & i & "*Summary!BN" & l & _
            "))-((Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & _
            ")*((Wood_Industrial!EQ" & j & "*Summary!BN" & l & _
            ")/(Wood_Industrial!ES" & j & "*Summary!BN" & m & ")))))" & Chr(10) & _
        "+((Wood_Industrial!EU" & i & "*Wood_Industrial!EV" & i & ")" & Chr(10) & _
            "*((Wood_Industrial!EW" & i & "* Wood_Industrial!EX" & i & _
            ")-(( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & _
            ")*(Wood_Industrial!EW" & j & "* Wood_Industrial!EX" & j & "))))" & Chr(10) & _
        "*Wood_Industrial!I" & i & ")" & Chr(10) & _
        "+(Summary!BR" & l & "*( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & "))" & Chr(10) & _
        "+((Wood_Industrial!FA" & j & "*Wood_Industrial!FB" & j & ")*( Wood_Industrial!EY" & i & "*Wood_Industrial!EZ" & i & ")))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 70).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 70) = hojUsu_Forecast.Cells(c_ini, 115)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 70).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 70) = 0
        
        End If
        
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 117) = hojUsu_Summary.Cells(k, 70).Value
    hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 70)
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(12, 21).Value = "=Summary!BR" & k - 1

End Sub
