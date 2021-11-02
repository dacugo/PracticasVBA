Attribute VB_Name = "Eq_Pulp_Paper_Industry"
Sub SUPPLY_PULP_PAPER_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 38).Formula = "=(((Pulp_Paper_Industry!J" & i & "*Pulp_Paper_Industry!K" & i & _
    "*(1-(Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & ")))" & Chr(10) & _
    "+((Pulp_Paper_Industry!L" & i & "*Pulp_Paper_Industry!M" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!N" & i & "*Pulp_Paper_Industry!O" & i & ")-((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & _
        ")*(Pulp_Paper_Industry!N" & j & "*Pulp_Paper_Industry!O" & j & "))))" & Chr(10) & _
    "+((Pulp_Paper_Industry!P" & i & "*Pulp_Paper_Industry!Q" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!R" & i & "*Pulp_Paper_Industry!S" & i & ")-((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & _
        ")*(Pulp_Paper_Industry!R" & j & "*Pulp_Paper_Industry!S" & j & "))))" & Chr(10) & _
    "+((Pulp_Paper_Industry!T" & i & "*Pulp_Paper_Industry!U" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!V" & i & "*Pulp_Paper_Industry!W" & i & ")-((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & _
        ")*(Pulp_Paper_Industry!V" & j & "*Pulp_Paper_Industry!W" & j & ")))))" & Chr(10) & _
    "*Pulp_Paper_Industry!B" & i & ")+(Summary!AL" & l & "*(Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & "))" & Chr(10) & _
    "+((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & ")*(Pulp_Paper_Industry!Z" & j & "*Pulp_Paper_Industry!AA" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 38).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 38) = hojUsu_Forecast.Cells(c_ini, 59)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 38).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 38) = 0
        
        End If
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 60) = hojUsu_Summary.Cells(k, 38).Value
    c_ini = c_ini + 1
        
Next k
        
Case 2, 4

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

    hojUsu_Summary.Cells(k, 38).Formula = "=(((Pulp_Paper_Industry!J" & i & "*Pulp_Paper_Industry!K" & i & _
    "*(1-(Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & ")))" & Chr(10) & _
    "+((Pulp_Paper_Industry!L" & i & "*Pulp_Paper_Industry!M" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!N" & i & "*Summary!AZ" & k & ")-((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & _
        ")*(Pulp_Paper_Industry!N" & j & "*Summary!AZ" & l & "))))" & Chr(10) & _
    "+((Pulp_Paper_Industry!P" & i & "*Pulp_Paper_Industry!Q" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!R" & i & "*Pulp_Paper_Industry!S" & i & ")-((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & _
        ")*(Pulp_Paper_Industry!R" & j & "*Pulp_Paper_Industry!S" & j & "))))" & Chr(10) & _
    "+((Pulp_Paper_Industry!T" & i & "*Pulp_Paper_Industry!U" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!V" & i & "*Pulp_Paper_Industry!W" & i & ")-((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & _
        ")*(Pulp_Paper_Industry!V" & j & "*Pulp_Paper_Industry!W" & j & ")))))" & Chr(10) & _
    "*Pulp_Paper_Industry!B" & i & ")+(Summary!AL" & l & "*(Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & "))" & Chr(10) & _
    "+((Pulp_Paper_Industry!X" & i & "*Pulp_Paper_Industry!Y" & i & ")*(Pulp_Paper_Industry!Z" & j & "*Pulp_Paper_Industry!AA" & j & "))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 38).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 38) = hojUsu_Forecast.Cells(c_ini, 59)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 38).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 38) = 0
        
        End If
        
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 61) = hojUsu_Summary.Cells(k, 38).Value
    c_ini = c_ini + 1
    
Next k

'Case 5
'
'For k = d_ini To d_fin
'
''i es el número 8 de la hoja del mercado
'i = k - 31
''j es el número 7 de la hoja del mercado
'j = i - 1
''l es el n anterior de la hoja en summary
'l = k - 1
''k es el año actual de la hoja summary
'm = l - 1
''m es el n dos años anteriores al año actual
'
'    hojUsu_Summary.Cells(k, 38).Formula = "=Module!AJ" & i & "*Pulp_Paper_Industry!B" & i
'
'    Select Case g
'
'    Case 1
'
'        If hojUsu_Summary.Cells(k, 38).Value < 0 Then
'
'        hojUsu_Summary.Cells(k, 38) = hojUsu_Forecast.Cells(c_ini, 59)
'
'        End If
'
'    Case 2
'
'        If hojUsu_Summary.Cells(k, 38).Value < 0 Then
'
'        hojUsu_Summary.Cells(k, 38) = 0
'
'        End If
'
'    Case 3
'
'    End Select
'
'    hojUsu_Forecast.Cells(c_ini, 61) = hojUsu_Summary.Cells(k, 38).Value
'    c_ini = c_ini + 1
'
'Next k

End Select

hojUsu_Summary.Cells(6, 7).Value = "=Summary!AL" & k - 1

End Sub
Sub CONSUMPTION_PULP_PAPER_INDUSTRY()

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
       
    hojUsu_PulpPaperIndustry.Cells(i, 54).Formula = _
        "=(((Pulp_Paper_Industry!AC" & i & "*Pulp_Paper_Industry!AD" & i & _
            "*(1-(Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AE" & i & "*Pulp_Paper_Industry!AF" & i & ")" & Chr(10) & _
        "*(((Pulp_Paper_Industry!AG" & i & "* Pulp_Paper_Industry!AH" & i & _
            ")/(Pulp_Paper_Industry!AI" & i & "* Pulp_Paper_Industry!AJ" & i & _
            "))-(( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & _
            ")*( (Pulp_Paper_Industry!AG" & j & "* Pulp_Paper_Industry!AH" & j & _
            ")/(Pulp_Paper_Industry!AI" & j & "* Pulp_Paper_Industry!AJ" & j & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AK" & i & "*Pulp_Paper_Industry!AL" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!AM" & i & "* Pulp_Paper_Industry!AN" & i & _
            ")-(( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & _
            ")*(Pulp_Paper_Industry!AM" & j & "* Pulp_Paper_Industry!AN" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AO" & i & "*Pulp_Paper_Industry!AP" & i & ")" & Chr(10) & _
        "*(((Pulp_Paper_Industry!AQ" & i & "* Pulp_Paper_Industry!AR" & i & _
            ")/(Pulp_Paper_Industry!AS" & i & "* Pulp_Paper_Industry!AT" & i & _
            "))-(( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & _
            ")*( (Pulp_Paper_Industry!AQ" & j & "* Pulp_Paper_Industry!AR" & j & _
            ")/(Pulp_Paper_Industry!AS" & j & "* Pulp_Paper_Industry!AT" & j & ")))))" & Chr(10) & _
        "*Pulp_Paper_Industry!C" & i & ")" & Chr(10) & _
        "+(Pulp_Paper_Industry!BB" & j & "*( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AW" & j & "*Pulp_Paper_Industry!AX" & j & _
        ")*( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & ")))"
            
    hojUsu_Summary.Cells(k, 40).Formula = _
    "=(Pulp_Paper_Industry!AY" & i & "*Pulp_Paper_Industry!AZ" & i & ")*(Pulp_Paper_Industry!BA" & i & "*Pulp_Paper_Industry!BB" & i & ")"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 40).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 40) = hojUsu_Forecast.Cells(c_ini, 63)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 40).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 40) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 64) = hojUsu_Summary.Cells(k, 40).Value
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

    hojUsu_PulpPaperIndustry.Cells(i, 54).Formula = _
        "=(((Pulp_Paper_Industry!AC" & i & "*Pulp_Paper_Industry!AD" & i & _
            "*(1-(Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AE" & i & "*Pulp_Paper_Industry!AF" & i & ")" & Chr(10) & _
        "*(((Pulp_Paper_Industry!AG" & i & "* Summary!AT" & k & _
            ")/(Pulp_Paper_Industry!AI" & i & "* Pulp_Paper_Industry!AJ" & i & _
            "))-(( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & _
            ")*( (Pulp_Paper_Industry!AG" & j & "* Summary!AT" & l & _
            ")/(Pulp_Paper_Industry!AI" & j & "* Pulp_Paper_Industry!AJ" & j & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AK" & i & "*Pulp_Paper_Industry!AL" & i & ")" & Chr(10) & _
        "*((Pulp_Paper_Industry!AM" & i & "* Pulp_Paper_Industry!AN" & i & _
            ")-(( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & _
            ")*(Pulp_Paper_Industry!AM" & j & "* Pulp_Paper_Industry!AN" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AO" & i & "*Pulp_Paper_Industry!AP" & i & ")" & Chr(10) & _
        "*(((Pulp_Paper_Industry!AQ" & i & "* Pulp_Paper_Industry!AR" & i & _
            ")/(Pulp_Paper_Industry!AS" & i & "* Pulp_Paper_Industry!AT" & i & _
            "))-(( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & _
            ")*( (Pulp_Paper_Industry!AQ" & j & "* Pulp_Paper_Industry!AR" & j & _
            ")/(Pulp_Paper_Industry!AS" & j & "* Pulp_Paper_Industry!AT" & j & ")))))" & Chr(10) & _
        "*Pulp_Paper_Industry!C" & i & ")" & Chr(10) & _
        "+(Pulp_Paper_Industry!BB" & j & "*( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!AW" & j & "*Pulp_Paper_Industry!AX" & j & _
        ")*( Pulp_Paper_Industry!AU" & i & "*Pulp_Paper_Industry!AV" & i & ")))"
            
    hojUsu_Summary.Cells(k, 40).Formula = _
    "=(Pulp_Paper_Industry!AY" & i & "*Pulp_Paper_Industry!AZ" & i & ")*(Pulp_Paper_Industry!BA" & i & "*Pulp_Paper_Industry!BB" & i & ")"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 40).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 40) = hojUsu_Forecast.Cells(c_ini, 63)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 40).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 40) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 65) = hojUsu_Summary.Cells(k, 40).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(7, 7).Value = "=Summary!AN" & k - 1

End Sub
Sub EXPORTS_PULP_PAPER_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 42).Formula = _
        "=(((Pulp_Paper_Industry!BD" & i & "*Pulp_Paper_Industry!BE" & i & _
            "*(1-(Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BF" & i & "*Pulp_Paper_Industry!BG" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!BH" & i & "* Pulp_Paper_Industry!BI" & i & _
            ")-(( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & _
            ")*(Pulp_Paper_Industry!BH" & j & "* Pulp_Paper_Industry!BI" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BJ" & i & "*Pulp_Paper_Industry!BK" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!BL" & i & "* Pulp_Paper_Industry!BM" & i & _
            ")/(Pulp_Paper_Industry!BN" & i & "* Pulp_Paper_Industry!BO" & i & _
            "))-(( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & _
            ")*( (Pulp_Paper_Industry!BL" & j & "* Pulp_Paper_Industry!BM" & j & _
            ")/(Pulp_Paper_Industry!BN" & j & "* Pulp_Paper_Industry!BO" & j & ")))))" & Chr(10) & _
        "*Pulp_Paper_Industry!D" & i & ")" & Chr(10) & _
        "+(Summary!AP" & l & Chr(10) & _
        "*( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BR" & j & "*Pulp_Paper_Industry!BS" & j & ")*( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 42).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 42) = hojUsu_Forecast.Cells(c_ini, 67)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 42).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 42) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 68) = hojUsu_Summary.Cells(k, 42).Value
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

    hojUsu_Summary.Cells(k, 42).Formula = _
        "=(((Pulp_Paper_Industry!BD" & i & "*Pulp_Paper_Industry!BE" & i & _
            "*(1-(Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BF" & i & "*Pulp_Paper_Industry!BG" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!BH" & i & "* Pulp_Paper_Industry!BI" & i & _
            ")-(( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & _
            ")*(Pulp_Paper_Industry!BH" & j & "* Pulp_Paper_Industry!BI" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BJ" & i & "*Pulp_Paper_Industry!BK" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!BL" & i & "* Summary!AV" & k & _
            ")/(Pulp_Paper_Industry!BN" & i & "* Pulp_Paper_Industry!BO" & i & _
            "))-(( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & _
            ")*( (Pulp_Paper_Industry!BL" & j & "* Summary!AV" & l & _
            ")/(Pulp_Paper_Industry!BN" & j & "* Pulp_Paper_Industry!BO" & j & ")))))" & Chr(10) & _
        "*Pulp_Paper_Industry!D" & i & ")" & Chr(10) & _
        "+(Summary!AP" & l & Chr(10) & _
        "*( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BR" & j & "*Pulp_Paper_Industry!BS" & j & ")*( Pulp_Paper_Industry!BP" & i & "*Pulp_Paper_Industry!BQ" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 42).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 42) = hojUsu_Forecast.Cells(c_ini, 67)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 42).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 42) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 69) = hojUsu_Summary.Cells(k, 42).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(8, 7).Value = "=Summary!AP" & k - 1

End Sub
Sub IMPORTS_PULP_PAPER_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 44).Formula = _
        "=(((Pulp_Paper_Industry!BU" & i & "*Pulp_Paper_Industry!BV" & i & _
            "*(1-(Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BW" & i & "*Pulp_Paper_Industry!BX" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!BY" & i & "* Pulp_Paper_Industry!BZ" & i & _
            ")-((Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & _
            ")*(Pulp_Paper_Industry!BY" & j & "* Pulp_Paper_Industry!BZ" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CA" & i & "*Pulp_Paper_Industry!CB" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!CC" & i & "* Pulp_Paper_Industry!CD" & i & _
            ")/(Pulp_Paper_Industry!CE" & i & "* Pulp_Paper_Industry!CF" & i & _
            "))-((Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & _
            ")*( (Pulp_Paper_Industry!CC" & j & "* Pulp_Paper_Industry!CD" & j & _
            ")/(Pulp_Paper_Industry!CE" & j & "* Pulp_Paper_Industry!CF" & j & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CG" & i & "*Pulp_Paper_Industry!CH" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!CI" & i & "* Pulp_Paper_Industry!CJ" & i & _
            ")-((Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & _
            ")*(Pulp_Paper_Industry!CI" & j & "* Pulp_Paper_Industry!CJ" & j & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!E" & i & ")" & Chr(10) & _
        "+(Summary!AR" & l & "*(Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CM" & j & "*Pulp_Paper_Industry!CN" & j & ")*( Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & ")))"
                    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 44).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 44) = hojUsu_Forecast.Cells(c_ini, 71)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 44).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 44) = 0
        
        End If
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 72) = hojUsu_Summary.Cells(k, 44).Value
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

    hojUsu_Summary.Cells(k, 44).Formula = _
        "=(((Pulp_Paper_Industry!BU" & i & "*Pulp_Paper_Industry!BV" & i & _
            "*(1-(Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!BW" & i & "*Pulp_Paper_Industry!BX" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!BY" & i & "* Summary!AN" & k & _
            ")-((Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & _
            ")*(Pulp_Paper_Industry!BY" & j & "* Summary!AN" & l & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CA" & i & "*Pulp_Paper_Industry!CB" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!CC" & i & "* Summary!AT" & k & _
            ")/(Pulp_Paper_Industry!CE" & i & "* Summary!AX" & k & _
            "))-((Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & _
            ")*( (Pulp_Paper_Industry!CC" & j & "* Summary!AT" & l & _
            ")/(Pulp_Paper_Industry!CE" & j & "* Summary!AX" & l & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CG" & i & "*Pulp_Paper_Industry!CH" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!CI" & i & "* Pulp_Paper_Industry!CJ" & i & _
            ")-((Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & _
            ")*(Pulp_Paper_Industry!CI" & j & "* Pulp_Paper_Industry!CJ" & j & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!E" & i & ")" & Chr(10) & _
        "+(Summary!AR" & l & "*(Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CM" & j & "*Pulp_Paper_Industry!CN" & j & ")*( Pulp_Paper_Industry!CK" & i & "*Pulp_Paper_Industry!CL" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 44).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 44) = hojUsu_Forecast.Cells(c_ini, 71)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 44).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 44) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 73) = hojUsu_Summary.Cells(k, 44).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(9, 7).Value = "=Summary!AR" & k - 1

End Sub
Sub PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 46).Formula = _
        "=(((Pulp_Paper_Industry!CP" & i & "*Pulp_Paper_Industry!CQ" & i & _
            "*(1-(Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CR" & i & "*Pulp_Paper_Industry!CS" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!CT" & i & "* Pulp_Paper_Industry!CU" & i & _
            ")/(Pulp_Paper_Industry!CV" & i & "* Pulp_Paper_Industry!CW" & i & _
            "))-(( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & _
            ")*( (Pulp_Paper_Industry!CT" & j & "* Pulp_Paper_Industry!CU" & j & _
            ")/(Pulp_Paper_Industry!CV" & j & "* Pulp_Paper_Industry!CW" & j & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CX" & i & "*Pulp_Paper_Industry!CY" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!CZ" & i & "* Pulp_Paper_Industry!DA" & i & _
            ")-(( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & _
            ")*(Pulp_Paper_Industry!CZ" & j & "* Pulp_Paper_Industry!DA" & j & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!F" & i & ")" & Chr(10) & _
        "+(Summary!AT" & l & "*( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DD" & j & "*Pulp_Paper_Industry!DE" & j & ")*( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 46).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 46) = hojUsu_Forecast.Cells(c_ini, 75)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 46).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 46) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 76) = hojUsu_Summary.Cells(k, 46).Value
    hojUsu_SetPricesPulpPaper.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 46)
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

    hojUsu_Summary.Cells(k, 46).Formula = _
        "=(((Pulp_Paper_Industry!CP" & i & "*Pulp_Paper_Industry!CQ" & i & _
            "*(1-(Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CR" & i & "*Pulp_Paper_Industry!CS" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!CT" & i & "* Summary!AL" & k & _
            ")/(Pulp_Paper_Industry!CV" & i & "* Pulp_Paper_Industry!CW" & i & _
            "))-(( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & _
            ")*( (Pulp_Paper_Industry!CT" & j & "* Summary!AL" & l & _
            ")/(Pulp_Paper_Industry!CV" & j & "* Pulp_Paper_Industry!CW" & j & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!CX" & i & "*Pulp_Paper_Industry!CY" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!CZ" & i & "* Summary!AT" & l & _
            ")-(( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & _
            ")*(Pulp_Paper_Industry!CZ" & j & "* Summary!AT" & m & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!F" & i & ")" & Chr(10) & _
        "+(Summary!AT" & l & "*( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DD" & j & "*Pulp_Paper_Industry!DE" & j & ")*( Pulp_Paper_Industry!DB" & i & "*Pulp_Paper_Industry!DC" & i & ")))"

            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 46).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 46) = hojUsu_Forecast.Cells(c_ini, 75)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 46).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 46) = 0
        
        End If
        
    Case 3
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 77) = hojUsu_Summary.Cells(k, 46).Value
    hojUsu_SetPricesPulpPaper.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 46)
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(10, 7).Value = "=Summary!AT" & k - 1

End Sub
Sub PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 48).Formula = _
        "=(((Pulp_Paper_Industry!DG" & i & "*Pulp_Paper_Industry!DH" & i & _
            "*(1-(Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DI" & i & "*Pulp_Paper_Industry!DJ" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!DK" & i & "* Pulp_Paper_Industry!DL" & i & _
            ")-(( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & _
            ")*(Pulp_Paper_Industry!DK" & j & "* Pulp_Paper_Industry!DL" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DM" & i & "*Pulp_Paper_Industry!DN" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!DO" & i & "* Pulp_Paper_Industry!DP" & i & _
            ")-(( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & _
            ")*(Pulp_Paper_Industry!DO" & j & "* Pulp_Paper_Industry!DP" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DQ" & i & "*Pulp_Paper_Industry!DR" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!DS" & i & "* Pulp_Paper_Industry!DT" & i & _
            ")-(( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & _
            ")*(Pulp_Paper_Industry!DS" & j & "* Pulp_Paper_Industry!DT" & j & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!G" & i & ")" & Chr(10) & _
        "+(Summary!AV" & l & "*( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DW" & j & "*Pulp_Paper_Industry!DX" & j & ")*( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 48).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 48) = hojUsu_Forecast.Cells(c_ini, 79)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 48).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 48) = 0
        
        End If
    
    End Select

    hojUsu_Forecast.Cells(c_ini, 80) = hojUsu_Summary.Cells(k, 48).Value
    hojUsu_SetPricesPulpPaper.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 48)
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

    hojUsu_Summary.Cells(k, 48).Formula = _
        "=(((Pulp_Paper_Industry!DG" & i & "*Pulp_Paper_Industry!DH" & i & _
            "*(1-(Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DI" & i & "*Pulp_Paper_Industry!DJ" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!DK" & i & "* Summary!AV" & l & _
            ")-(( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & _
            ")*(Pulp_Paper_Industry!DK" & j & "* Summary!AV" & m & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DM" & i & "*Pulp_Paper_Industry!DN" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!DO" & i & "* Pulp_Paper_Industry!DP" & i & _
            ")-(( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & _
            ")*(Pulp_Paper_Industry!DO" & j & "* Pulp_Paper_Industry!DP" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DQ" & i & "*Pulp_Paper_Industry!DR" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!DS" & i & "* Pulp_Paper_Industry!DT" & i & _
            ")-(( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & _
            ")*(Pulp_Paper_Industry!DS" & j & "* Pulp_Paper_Industry!DT" & j & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!G" & i & ")" & Chr(10) & _
        "+(Summary!AV" & l & "*( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!DW" & j & "*Pulp_Paper_Industry!DX" & j & ")*( Pulp_Paper_Industry!DU" & i & "*Pulp_Paper_Industry!DV" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 48).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 48) = hojUsu_Forecast.Cells(c_ini, 79)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 48).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 48) = 0
        
        End If
    
    Case 3
    
    End Select
                       
    hojUsu_Forecast.Cells(c_ini, 81) = hojUsu_Summary.Cells(k, 48).Value
    hojUsu_SetPricesPulpPaper.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 48)
    c_ini = c_ini + 1

   
Next k

End Select

hojUsu_Summary.Cells(11, 7).Value = "=Summary!AV" & k - 1

End Sub
Sub PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 50).Formula = _
        "=(((Pulp_Paper_Industry!DZ" & i & "*Pulp_Paper_Industry!EA" & i & _
            "*(1-(Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!EB" & i & "*Pulp_Paper_Industry!EC" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!ED" & i & "* Pulp_Paper_Industry!EE" & i & _
            ")-(( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & _
            ")*(Pulp_Paper_Industry!ED" & j & "* Pulp_Paper_Industry!EE" & j & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!EF" & i & "*Pulp_Paper_Industry!EG" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!EH" & i & "* Pulp_Paper_Industry!EI" & i & _
            ")/(Pulp_Paper_Industry!EJ" & i & "* Pulp_Paper_Industry!EK" & i & _
            "))-((Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & _
            ")*((Pulp_Paper_Industry!EH" & j & "* Pulp_Paper_Industry!EI" & j & _
            ")/(Pulp_Paper_Industry!EJ" & j & "* Pulp_Paper_Industry!EK" & j & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!EL" & i & "*Pulp_Paper_Industry!EM" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!EN" & i & "* Pulp_Paper_Industry!EO" & i & _
            ")-(( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & _
            ")*(Pulp_Paper_Industry!EN" & j & "* Pulp_Paper_Industry!EO" & j & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!H" & i & ")" & Chr(10) & _
        "+(Summary!AX" & l & "*( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!ER" & j & "*Pulp_Paper_Industry!ES" & j & ")*( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 50).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 50) = hojUsu_Forecast.Cells(c_ini, 83)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 50).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 50) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 84) = hojUsu_Summary.Cells(k, 50).Value
    hojUsu_SetPricesPulpPaper.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 50)
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

    hojUsu_Summary.Cells(k, 50).Formula = _
        "=(((Pulp_Paper_Industry!DZ" & i & "*Pulp_Paper_Industry!EA" & i & _
            "*(1-(Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & ")))" & Chr(10) & _
        "+((Pulp_Paper_Industry!EB" & i & "*Pulp_Paper_Industry!EC" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!ED" & i & "* Summary!AX" & l & _
            ")-((Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & _
            ")*(Pulp_Paper_Industry!ED" & j & "* Summary!AX" & m & "))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!EF" & i & "*Pulp_Paper_Industry!EG" & i & ")" & Chr(10) & _
            "*(((Pulp_Paper_Industry!EH" & i & "* Summary!AT" & k & _
            ")/(Pulp_Paper_Industry!EJ" & i & "* Summary!AT" & l & _
            "))-(( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & _
            ")*( (Pulp_Paper_Industry!EH" & j & "* Summary!AT" & l & _
            ")/(Pulp_Paper_Industry!EJ" & j & "* Summary!AT" & m & ")))))" & Chr(10) & _
        "+((Pulp_Paper_Industry!EL" & i & "*Pulp_Paper_Industry!EM" & i & ")" & Chr(10) & _
            "*((Pulp_Paper_Industry!EN" & i & "* Pulp_Paper_Industry!EO" & i & _
            ")-(( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & _
            ")*(Pulp_Paper_Industry!EN" & j & "* Pulp_Paper_Industry!EO" & j & "))))" & Chr(10) & _
        "*Pulp_Paper_Industry!H" & i & ")" & Chr(10) & _
        "+(Summary!AX" & l & "*( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & "))" & Chr(10) & _
        "+((Pulp_Paper_Industry!ER" & j & "*Pulp_Paper_Industry!ES" & j & ")*( Pulp_Paper_Industry!EP" & i & "*Pulp_Paper_Industry!EQ" & i & ")))"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 50).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 50) = hojUsu_Forecast.Cells(c_ini, 83)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 50).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 50) = 0
        
        End If
        
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 85) = hojUsu_Summary.Cells(k, 50).Value
    hojUsu_SetPricesPulpPaper.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 50)
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(12, 7).Value = "=Summary!N" & k - 1

End Sub
