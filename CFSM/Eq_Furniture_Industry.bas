Attribute VB_Name = "Eq_Furniture_Industry"
Sub SUPPLY_FURNITURE_INDUSTRY()

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
                    
            hojUsu_Summary.Cells(k, 20).Formula = "=(((Furniture_Industry!J" & i & "*Furniture_Industry!K" & i & _
            "*(1-(Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & ")))" & Chr(10) & _
            "+((Furniture_Industry!L" & i & "*Furniture_Industry!M" & i & ")" & Chr(10) & _
                "*((Furniture_Industry!N" & i & "*Furniture_Industry!O" & i & ")-((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & _
                ")*(Furniture_Industry!N" & j & "*Furniture_Industry!O" & j & "))))" & Chr(10) & _
            "+((Furniture_Industry!P" & i & "*Furniture_Industry!Q" & i & ")" & Chr(10) & _
                "*((Furniture_Industry!R" & i & "*Furniture_Industry!S" & i & ")-((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & _
                ")*(Furniture_Industry!R" & j & "*Furniture_Industry!S" & j & "))))" & Chr(10) & _
            "+((Furniture_Industry!T" & i & "*Furniture_Industry!U" & i & ")" & Chr(10) & _
                "*((Furniture_Industry!V" & i & "*Furniture_Industry!W" & i & ")-((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & _
                ")*(Furniture_Industry!V" & j & "*Furniture_Industry!W" & j & ")))))" & Chr(10) & _
            "*Furniture_Industry!B" & i & _
            ")+(Summary!T" & l & "*(Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & "))" & Chr(10) & _
            "+((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & ")*(Furniture_Industry!Z" & j & "*Furniture_Industry!AA" & j & "))"
            
            Select Case g
            
                Case 1
                
                    If hojUsu_Summary.Cells(k, 20).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 20) = hojUsu_Forecast.Cells(c_ini, 31)
                    
                    End If
                
                Case 2
                    
                    If hojUsu_Summary.Cells(k, 20).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 20) = 0
                    
                    End If
    
            End Select
            
            hojUsu_Forecast.Cells(c_ini, 32) = hojUsu_Summary.Cells(k, 20).Value
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
        
            hojUsu_Summary.Cells(k, 20).Formula = "=(((Furniture_Industry!J" & i & "*Furniture_Industry!K" & i & _
            "*(1-(Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & ")))" & Chr(10) & _
            "+((Furniture_Industry!L" & i & "*Furniture_Industry!M" & i & ")" & Chr(10) & _
                "*((Furniture_Industry!N" & i & "*Summary!AH" & k & ")-((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & _
                ")*(Furniture_Industry!N" & j & "*Summary!AH" & l & "))))" & Chr(10) & _
            "+((Furniture_Industry!P" & i & "*Furniture_Industry!Q" & i & ")" & Chr(10) & _
                "*((Furniture_Industry!R" & i & "*Furniture_Industry!S" & i & ")-((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & _
                ")*(Furniture_Industry!R" & j & "*Furniture_Industry!S" & j & "))))" & Chr(10) & _
            "+((Furniture_Industry!T" & i & "*Furniture_Industry!U" & i & ")" & Chr(10) & _
                "*((Furniture_Industry!V" & i & "*Furniture_Industry!W" & i & ")-((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & _
                ")*(Furniture_Industry!V" & j & "*Furniture_Industry!W" & j & ")))))" & Chr(10) & _
            "*Furniture_Industry!B" & i & _
            ")+(Summary!T" & l & "*(Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & "))" & Chr(10) & _
            "+((Furniture_Industry!X" & i & "*Furniture_Industry!Y" & i & ")*(Furniture_Industry!Z" & j & "*Furniture_Industry!AA" & j & "))"
        
            Select Case g
            
                Case 1
                
                    If hojUsu_Summary.Cells(k, 20).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 20) = hojUsu_Forecast.Cells(c_ini, 31)
                    
                    End If
                
                Case 2
                    
                    If hojUsu_Summary.Cells(k, 20).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 20) = 0
                    
                    End If
                    
            End Select
        
            hojUsu_Forecast.Cells(c_ini, 33) = hojUsu_Summary.Cells(k, 20).Value
            c_ini = c_ini + 1
            
        Next k
    
'        Case 5
'
'        For k = d_ini To d_fin
'
'            'i es el número 8 de la hoja del mercado
'            i = k - 31
'            'j es el número 7 de la hoja del mercado
'            j = i - 1
'            'l es el n anterior de la hoja en summary
'            l = k - 1
'            'k es el año actual de la hoja summary
'            m = l - 1
'            'm es el n dos años anteriores al año actual
'
'            hojUsu_Summary.Cells(k, 20).Formula = "=Module!AH" & i & "*Furniture_Industry!B" & i
'
'            Select Case g
'
'                Case 1
'
'                    If hojUsu_Summary.Cells(k, 20).Value < 0 Then
'
'                    hojUsu_Summary.Cells(k, 20) = hojUsu_Forecast.Cells(c_ini, 31)
'
'                    End If
'
'                Case 2
'
'                    If hojUsu_Summary.Cells(k, 20).Value < 0 Then
'
'                    hojUsu_Summary.Cells(k, 20) = 0
'
'                    End If
'
'            End Select
'
'            hojUsu_Forecast.Cells(c_ini, 33) = hojUsu_Summary.Cells(k, 20).Value
'            c_ini = c_ini + 1
'
'        Next k
    
    End Select
    
    hojUsu_Summary.Cells(6, 5).Value = "=Summary!T" & k - 1

End Sub
Sub CONSUMPTION_FURNITURE_INDUSTRY()

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
       
    hojUsu_FurnitureIndustry.Cells(i, 54).Formula = _
        "=(((Furniture_Industry!AC" & i & "*Furniture_Industry!AD" & i & _
            "*(1-(Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!AE" & i & "*Furniture_Industry!AF" & i & ")" & Chr(10) & _
        "*(((Furniture_Industry!AG" & i & "* Furniture_Industry!AH" & i & _
            ")/(Furniture_Industry!AI" & i & "* Furniture_Industry!AJ" & i & _
            "))-(( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & _
            ")*( (Furniture_Industry!AG" & j & "* Furniture_Industry!AH" & j & _
            ")/(Furniture_Industry!AI" & j & "* Furniture_Industry!AJ" & j & ")))))" & Chr(10) & _
        "+((Furniture_Industry!AK" & i & "*Furniture_Industry!AL" & i & ")" & Chr(10) & _
        "*((Furniture_Industry!AM" & i & "* Furniture_Industry!AN" & i & _
            ")-(( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & _
            ")*(Furniture_Industry!AM" & j & "* Furniture_Industry!AN" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!AO" & i & "*Furniture_Industry!AP" & i & ")" & Chr(10) & _
        "*(((Furniture_Industry!AQ" & i & "* Furniture_Industry!AR" & i & _
            ")/(Furniture_Industry!AS" & i & "* Furniture_Industry!AT" & i & _
            "))-(( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & _
            ")*( (Furniture_Industry!AQ" & j & "* Furniture_Industry!AR" & j & _
            ")/(Furniture_Industry!AS" & j & "* Furniture_Industry!AT" & j & ")))))" & Chr(10) & _
        "*Furniture_Industry!C" & i & ")" & Chr(10) & _
        "+(Furniture_Industry!BB" & j & "*( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!AW" & j & "*Furniture_Industry!AX" & j & _
        ")*( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & ")))"
            
    hojUsu_Summary.Cells(k, 22).Formula = _
    "=(Furniture_Industry!AY" & i & "*Furniture_Industry!AZ" & i & ")*(Furniture_Industry!BA" & i & "*Furniture_Industry!BB" & i & ")"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 22).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 22) = hojUsu_Forecast.Cells(c_ini, 35)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 22).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 22) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 36) = hojUsu_Summary.Cells(k, 22).Value
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

    hojUsu_FurnitureIndustry.Cells(i, 54).Formula = _
        "=(((Furniture_Industry!AC" & i & "*Furniture_Industry!AD" & i & _
            "*(1-(Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!AE" & i & "*Furniture_Industry!AF" & i & ")" & Chr(10) & _
        "*(((Furniture_Industry!AG" & i & "* Summary!AB" & k & _
            ")/(Furniture_Industry!AI" & i & "* Furniture_Industry!AJ" & i & _
            "))-(( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & _
            ")*( (Furniture_Industry!AG" & j & "* Summary!AB" & l & _
            ")/(Furniture_Industry!AI" & j & "* Furniture_Industry!AJ" & j & ")))))" & Chr(10) & _
        "+((Furniture_Industry!AK" & i & "*Furniture_Industry!AL" & i & ")" & Chr(10) & _
        "*((Furniture_Industry!AM" & i & "* Furniture_Industry!AN" & i & _
            ")-(( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & _
            ")*(Furniture_Industry!AM" & j & "* Furniture_Industry!AN" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!AO" & i & "*Furniture_Industry!AP" & i & ")" & Chr(10) & _
        "*(((Furniture_Industry!AQ" & i & "* Furniture_Industry!AR" & i & _
            ")/(Furniture_Industry!AS" & i & "* Furniture_Industry!AT" & i & _
            "))-(( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & _
            ")*( (Furniture_Industry!AQ" & j & "* Furniture_Industry!AR" & j & _
            ")/(Furniture_Industry!AS" & j & "* Furniture_Industry!AT" & j & ")))))" & Chr(10) & _
        "*Furniture_Industry!C" & i & ")" & Chr(10) & _
        "+(Furniture_Industry!BB" & j & "*( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!AW" & j & "*Furniture_Industry!AX" & j & _
        ")*( Furniture_Industry!AU" & i & "*Furniture_Industry!AV" & i & ")))"
            
    hojUsu_Summary.Cells(k, 22).Formula = _
    "=(Furniture_Industry!AY" & i & "*Furniture_Industry!AZ" & i & ")*(Furniture_Industry!BA" & i & "*Furniture_Industry!BB" & i & ")"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 22).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 22) = hojUsu_Forecast.Cells(c_ini, 35)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 22).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 22) = 0
        
        End If

    End Select
    
    hojUsu_Forecast.Cells(c_ini, 37) = hojUsu_Summary.Cells(k, 22).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(7, 5).Value = "=Summary!V" & k - 1

End Sub
Sub EXPORTS_FURNITURE_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 24).Formula = _
        "=(((Furniture_Industry!BD" & i & "*Furniture_Industry!BE" & i & _
            "*(1-(Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!BF" & i & "*Furniture_Industry!BG" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!BH" & i & "* Furniture_Industry!BI" & i & _
            ")-(( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & _
            ")*(Furniture_Industry!BH" & j & "* Furniture_Industry!BI" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!BJ" & i & "*Furniture_Industry!BK" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!BL" & i & "* Furniture_Industry!BM" & i & _
            ")/(Furniture_Industry!BN" & i & "* Furniture_Industry!BO" & i & _
            "))-(( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & _
            ")*( (Furniture_Industry!BL" & j & "* Furniture_Industry!BM" & j & _
            ")/(Furniture_Industry!BN" & j & "* Furniture_Industry!BO" & j & ")))))" & Chr(10) & _
        "*Furniture_Industry!D" & i & ")" & Chr(10) & _
        "+(Summary!X" & l & Chr(10) & _
        "*( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!BR" & j & "*Furniture_Industry!BS" & j & ")*( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 24).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 24) = hojUsu_Forecast.Cells(c_ini, 39)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 24).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 24) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 40) = hojUsu_Summary.Cells(k, 24).Value
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

    hojUsu_Summary.Cells(k, 24).Formula = _
        "=(((Furniture_Industry!BD" & i & "*Furniture_Industry!BE" & i & _
            "*(1-(Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!BF" & i & "*Furniture_Industry!BG" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!BH" & i & "* Furniture_Industry!BI" & i & _
            ")-(( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & _
            ")*(Furniture_Industry!BH" & j & "* Furniture_Industry!BI" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!BJ" & i & "*Furniture_Industry!BK" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!BL" & i & "* Summary!AD" & k & _
            ")/(Furniture_Industry!BN" & i & "* Furniture_Industry!BO" & i & _
            "))-(( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & _
            ")*( (Furniture_Industry!BL" & j & "* Summary!AD" & l & _
            ")/(Furniture_Industry!BN" & j & "* Furniture_Industry!BO" & j & ")))))" & Chr(10) & _
        "*Furniture_Industry!D" & i & ")" & Chr(10) & _
        "+(Summary!X" & l & Chr(10) & _
        "*( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!BR" & j & "*Furniture_Industry!BS" & j & ")*( Furniture_Industry!BP" & i & "*Furniture_Industry!BQ" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 24).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 24) = hojUsu_Forecast.Cells(c_ini, 39)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 24).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 24) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 41) = hojUsu_Summary.Cells(k, 24).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(8, 5).Value = "=Summary!X" & k - 1

End Sub
Sub IMPORTS_FURNITURE_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 26).Formula = _
        "=(((Furniture_Industry!BU" & i & "*Furniture_Industry!BV" & i & _
            "*(1-(Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!BW" & i & "*Furniture_Industry!BX" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!BY" & i & "* Furniture_Industry!BZ" & i & _
            ")-((Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & _
            ")*(Furniture_Industry!BY" & j & "* Furniture_Industry!BZ" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!CA" & i & "*Furniture_Industry!CB" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!CC" & i & "* Furniture_Industry!CD" & i & _
            ")/(Furniture_Industry!CE" & i & "* Furniture_Industry!CF" & i & _
            "))-((Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & _
            ")*( (Furniture_Industry!CC" & j & "* Furniture_Industry!CD" & j & _
            ")/(Furniture_Industry!CE" & j & "* Furniture_Industry!CF" & j & ")))))" & Chr(10) & _
        "+((Furniture_Industry!CG" & i & "*Furniture_Industry!CH" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!CI" & i & "* Furniture_Industry!CJ" & i & _
            ")-((Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & _
            ")*(Furniture_Industry!CI" & j & "* Furniture_Industry!CJ" & j & "))))" & Chr(10) & _
        "*Furniture_Industry!E" & i & ")" & Chr(10) & _
        "+(Summary!Z" & l & "*(Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!CM" & j & "*Furniture_Industry!CN" & j & ")*( Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & ")))"
                    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 26).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 26) = hojUsu_Forecast.Cells(c_ini, 43)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 26).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 26) = 0
        
        End If
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 44) = hojUsu_Summary.Cells(k, 26).Value
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

    hojUsu_Summary.Cells(k, 26).Formula = _
        "=(((Furniture_Industry!BU" & i & "*Furniture_Industry!BV" & i & _
            "*(1-(Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!BW" & i & "*Furniture_Industry!BX" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!BY" & i & "* Summary!V" & k & _
            ")-((Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & _
            ")*(Furniture_Industry!BY" & j & "* Summary!V" & l & "))))" & Chr(10) & _
        "+((Furniture_Industry!CA" & i & "*Furniture_Industry!CB" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!CC" & i & "* Summary!AB" & k & _
            ")/(Furniture_Industry!CE" & i & "* Summary!AF" & k & _
            "))-((Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & _
            ")*( (Furniture_Industry!CC" & j & "* Summary!AB" & l & _
            ")/(Furniture_Industry!CE" & j & "* Summary!AF" & l & ")))))" & Chr(10) & _
        "+((Furniture_Industry!CG" & i & "*Furniture_Industry!CH" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!CI" & i & "* Furniture_Industry!CJ" & i & _
            ")-((Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & _
            ")*(Furniture_Industry!CI" & j & "* Furniture_Industry!CJ" & j & "))))" & Chr(10) & _
        "*Furniture_Industry!E" & i & ")" & Chr(10) & _
        "+(Summary!Z" & l & "*(Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!CM" & j & "*Furniture_Industry!CN" & j & ")*( Furniture_Industry!CK" & i & "*Furniture_Industry!CL" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 26).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 26) = hojUsu_Forecast.Cells(c_ini, 43)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 26).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 26) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 45) = hojUsu_Summary.Cells(k, 26).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(9, 5).Value = "=Summary!Z" & k - 1

End Sub
Sub PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 28).Formula = _
        "=(((Furniture_Industry!CP" & i & "*Furniture_Industry!CQ" & i & _
            "*(1-(Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!CR" & i & "*Furniture_Industry!CS" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!CT" & i & "* Furniture_Industry!CU" & i & _
            ")/(Furniture_Industry!CV" & i & "* Furniture_Industry!CW" & i & _
            "))-(( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & _
            ")*( (Furniture_Industry!CT" & j & "* Furniture_Industry!CU" & j & _
            ")/(Furniture_Industry!CV" & j & "* Furniture_Industry!CW" & j & ")))))" & Chr(10) & _
        "+((Furniture_Industry!CX" & i & "*Furniture_Industry!CY" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!CZ" & i & "* Furniture_Industry!DA" & i & _
            ")-(( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & _
            ")*(Furniture_Industry!CZ" & j & "* Furniture_Industry!DA" & j & "))))" & Chr(10) & _
        "*Furniture_Industry!F" & i & ")" & Chr(10) & _
        "+(Summary!AB" & l & "*( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!DD" & j & "*Furniture_Industry!DE" & j & ")*( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 28).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 28) = hojUsu_Forecast.Cells(c_ini, 47)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 28).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 28) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 48) = hojUsu_Summary.Cells(k, 28).Value
    hojUsu_SetPricesFurniture.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 28)

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

    hojUsu_Summary.Cells(k, 28).Formula = _
        "=(((Furniture_Industry!CP" & i & "*Furniture_Industry!CQ" & i & _
            "*(1-(Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!CR" & i & "*Furniture_Industry!CS" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!CT" & i & "* Summary!T" & k & _
            ")/(Furniture_Industry!CV" & i & "* Furniture_Industry!CW" & i & _
            "))-(( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & _
            ")*( (Furniture_Industry!CT" & j & "* Summary!T" & l & _
            ")/(Furniture_Industry!CV" & j & "* Furniture_Industry!CW" & j & ")))))" & Chr(10) & _
        "+((Furniture_Industry!CX" & i & "*Furniture_Industry!CY" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!CZ" & i & "* Summary!AB" & l & _
            ")-(( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & _
            ")*(Furniture_Industry!CZ" & j & "* Summary!AB" & m & "))))" & Chr(10) & _
        "*Furniture_Industry!F" & i & ")" & Chr(10) & _
        "+(Summary!AB" & l & "*( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!DD" & j & "*Furniture_Industry!DE" & j & ")*( Furniture_Industry!DB" & i & "*Furniture_Industry!DC" & i & ")))"

            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 28).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 28) = hojUsu_Forecast.Cells(c_ini, 47)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 28).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 28) = 0
        
        End If
        
    Case 3
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 49) = hojUsu_Summary.Cells(k, 28).Value
    hojUsu_SetPricesFurniture.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 28)

    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(10, 5).Value = "=Summary!AB" & k - 1

End Sub
Sub PRICE_OF_EXPORTS_FURNITURE_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 30).Formula = _
        "=(((Furniture_Industry!DG" & i & "*Furniture_Industry!DH" & i & _
            "*(1-(Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!DI" & i & "*Furniture_Industry!DJ" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!DK" & i & "* Furniture_Industry!DL" & i & _
            ")-(( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & _
            ")*(Furniture_Industry!DK" & j & "* Furniture_Industry!DL" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!DM" & i & "*Furniture_Industry!DN" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!DO" & i & "* Furniture_Industry!DP" & i & _
            ")-(( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & _
            ")*(Furniture_Industry!DO" & j & "* Furniture_Industry!DP" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!DQ" & i & "*Furniture_Industry!DR" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!DS" & i & "* Furniture_Industry!DT" & i & _
            ")-(( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & _
            ")*(Furniture_Industry!DS" & j & "* Furniture_Industry!DT" & j & "))))" & Chr(10) & _
        "*Furniture_Industry!G" & i & ")" & Chr(10) & _
        "+(Summary!AD" & l & "*( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!DW" & j & "*Furniture_Industry!DX" & j & ")*( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 30).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 30) = hojUsu_Forecast.Cells(c_ini, 51)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 30).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 30) = 0
        
        End If
    
    End Select

    hojUsu_Forecast.Cells(c_ini, 52) = hojUsu_Summary.Cells(k, 30).Value
    hojUsu_SetPricesFurniture.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 30)

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

    hojUsu_Summary.Cells(k, 30).Formula = _
        "=(((Furniture_Industry!DG" & i & "*Furniture_Industry!DH" & i & _
            "*(1-(Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!DI" & i & "*Furniture_Industry!DJ" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!DK" & i & "* Summary!AD" & l & _
            ")-(( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & _
            ")*(Furniture_Industry!DK" & j & "* Summary!AD" & m & "))))" & Chr(10) & _
        "+((Furniture_Industry!DM" & i & "*Furniture_Industry!DN" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!DO" & i & "* Furniture_Industry!DP" & i & _
            ")-(( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & _
            ")*(Furniture_Industry!DO" & j & "* Furniture_Industry!DP" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!DQ" & i & "*Furniture_Industry!DR" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!DS" & i & "* Furniture_Industry!DT" & i & _
            ")-(( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & _
            ")*(Furniture_Industry!DS" & j & "* Furniture_Industry!DT" & j & "))))" & Chr(10) & _
        "*Furniture_Industry!G" & i & ")" & Chr(10) & _
        "+(Summary!AD" & l & "*( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!DW" & j & "*Furniture_Industry!DX" & j & ")*( Furniture_Industry!DU" & i & "*Furniture_Industry!DV" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 30).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 30) = hojUsu_Forecast.Cells(c_ini, 51)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 30).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 30) = 0
        
        End If
    
    End Select
                       
    hojUsu_Forecast.Cells(c_ini, 53) = hojUsu_Summary.Cells(k, 30).Value
    hojUsu_SetPricesFurniture.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 30)

    c_ini = c_ini + 1

   
Next k

End Select

hojUsu_Summary.Cells(11, 5).Value = "=Summary!AD" & k - 1

End Sub
Sub PRICE_OF_IMPORT_FURNITURE_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 32).Formula = _
        "=(((Furniture_Industry!DZ" & i & "*Furniture_Industry!EA" & i & _
            "*(1-(Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!EB" & i & "*Furniture_Industry!EC" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!ED" & i & "* Furniture_Industry!EE" & i & _
            ")-(( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & _
            ")*(Furniture_Industry!ED" & j & "* Furniture_Industry!EE" & j & "))))" & Chr(10) & _
        "+((Furniture_Industry!EF" & i & "*Furniture_Industry!EG" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!EH" & i & "* Furniture_Industry!EI" & i & _
            ")/(Furniture_Industry!EJ" & i & "* Furniture_Industry!EK" & i & _
            "))-((Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & _
            ")*((Furniture_Industry!EH" & j & "* Furniture_Industry!EI" & j & _
            ")/(Furniture_Industry!EJ" & j & "* Furniture_Industry!EK" & j & ")))))" & Chr(10) & _
        "+((Furniture_Industry!EL" & i & "*Furniture_Industry!EM" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!EN" & i & "* Furniture_Industry!EO" & i & _
            ")-(( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & _
            ")*(Furniture_Industry!EN" & j & "* Furniture_Industry!EO" & j & "))))" & Chr(10) & _
        "*Furniture_Industry!H" & i & ")" & Chr(10) & _
        "+(Summary!AF" & l & "*( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!ER" & j & "*Furniture_Industry!ES" & j & ")*( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 32).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 32) = hojUsu_Forecast.Cells(c_ini, 55)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 32).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 32) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 56) = hojUsu_Summary.Cells(k, 32).Value
    hojUsu_SetPricesFurniture.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 32)
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

    hojUsu_Summary.Cells(k, 32).Formula = _
        "=(((Furniture_Industry!DZ" & i & "*Furniture_Industry!EA" & i & _
            "*(1-(Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & ")))" & Chr(10) & _
        "+((Furniture_Industry!EB" & i & "*Furniture_Industry!EC" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!ED" & i & "* Summary!AF" & l & _
            ")-((Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & _
            ")*(Furniture_Industry!ED" & j & "* Summary!AF" & m & "))))" & Chr(10) & _
        "+((Furniture_Industry!EF" & i & "*Furniture_Industry!EG" & i & ")" & Chr(10) & _
            "*(((Furniture_Industry!EH" & i & "* Summary!AB" & k & _
            ")/(Furniture_Industry!EJ" & i & "* Summary!AB" & l & _
            "))-(( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & _
            ")*( (Furniture_Industry!EH" & j & "* Summary!AB" & l & _
            ")/(Furniture_Industry!EJ" & j & "* Summary!AB" & m & ")))))" & Chr(10) & _
        "+((Furniture_Industry!EL" & i & "*Furniture_Industry!EM" & i & ")" & Chr(10) & _
            "*((Furniture_Industry!EN" & i & "* Furniture_Industry!EO" & i & _
            ")-(( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & _
            ")*(Furniture_Industry!EN" & j & "* Furniture_Industry!EO" & j & "))))" & Chr(10) & _
        "*Furniture_Industry!H" & i & ")" & Chr(10) & _
        "+(Summary!AF" & l & "*( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & "))" & Chr(10) & _
        "+((Furniture_Industry!ER" & j & "*Furniture_Industry!ES" & j & ")*( Furniture_Industry!EP" & i & "*Furniture_Industry!EQ" & i & ")))"

    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 32).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 32) = hojUsu_Forecast.Cells(c_ini, 55)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 32).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 32) = 0
        
        End If
        
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 57) = hojUsu_Summary.Cells(k, 32).Value
    hojUsu_SetPricesFurniture.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 32)
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(12, 5).Value = "=Summary!AF" & k - 1

End Sub
