Attribute VB_Name = "Eq_Wood_Industry"
Sub SUPPLY_WOOD_INDUSTRY()

'Asignacion de años seleccionados y opciones escogidas por el usuario
Call AsignacionVariablesOpcionesUsuario

Select Case selectProcess

    'Validation
    Case 2

    For k = d_ini To d_fin
            
        Call AsignacionVariablesProcesos
            
        hojUsu_Summary.Cells(k, 2).Formula = "=(((Wood_Industry!J" & i & "*Wood_Industry!K" & i & _
        "*(1-(Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!L" & i & "*Wood_Industry!M" & i & ")" & Chr(10) & _
            "*((Wood_Industry!N" & i & "*Wood_Industry!O" & i & ")-((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*(Wood_Industry!N" & j & "*Wood_Industry!O" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!P" & i & "*Wood_Industry!Q" & i & ")" & Chr(10) & _
            "*((Wood_Industry!R" & i & "*Wood_Industry!S" & i & ")-((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*(Wood_Industry!R" & j & "*Wood_Industry!S" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!T" & i & "*Wood_Industry!U" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!V" & i & "* Wood_Industry!W" & i & ")/(Wood_Industry!X" & i & "* Wood_Industry!Y" & i & _
            "))-((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*((Wood_Industry!V" & j & "* Wood_Industry!W" & j & ")/(Wood_Industry!X" & j & "* Wood_Industry!Y" & j & "))))))" & Chr(10) & _
        "*Wood_Industry!B" & i & ")+(Summary!B" & l & "*(Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & "))" & Chr(10) & _
        "+((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & ")*(Wood_Industry!AB" & j & "*Wood_Industry!AC" & j & "))"
        
        If negativeData = 1 Then
            If hojUsu_Summary.Cells(k, 2).Value < 0 Then
                hojUsu_Summary.Cells(k, 2) = hojUsu_Forecast.Cells(c_ini, 3)
            End If
        End If
                
        hojUsu_Forecast.Cells(c_ini, 4) = hojUsu_Summary.Cells(k, 2).Value
        c_ini = c_ini + 1
            
    Next k
            
    'Opciones de Market Clearing Condition (3-Isolated)
    Case 3
    
    For k = d_ini To d_fin
            
        Call AsignacionVariablesProcesos
    
        hojUsu_Summary.Cells(k, 2).Formula = "=(((Wood_Industry!J" & i & "*Wood_Industry!K" & i & _
        "*(1-(Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!L" & i & "*Wood_Industry!M" & i & ")" & Chr(10) & _
            "*((Wood_Industry!N" & i & "*Summary!P" & k & ")-((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*(Wood_Industry!N" & j & "*Summary!P" & l & "))))" & Chr(10) & _
        "+((Wood_Industry!P" & i & "*Wood_Industry!Q" & i & ")" & Chr(10) & _
            "*((Wood_Industry!R" & i & "*Wood_Industry!S" & i & ")-((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*(Wood_Industry!R" & j & "*Wood_Industry!S" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!T" & i & "*Wood_Industry!U" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!V" & i & "* Wood_Industry!W" & i & ")/(Wood_Industry!X" & i & "* Wood_Industry!Y" & i & _
            "))-(( Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*( (Wood_Industry!V" & j & "* Wood_Industry!W" & j & ")/(Wood_Industry!X" & j & "* Wood_Industry!Y" & j & "))))))" & Chr(10) & _
        "*Wood_Industry!B" & i & ")+(Summary!B" & l & "*(Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & "))" & Chr(10) & _
        "+((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & ")*(Wood_Industry!AB" & j & "*Wood_Industry!AC" & j & "))"
        
        If negativeData = 1 Then
            If hojUsu_Summary.Cells(k, 2).Value < 0 Then
                hojUsu_Summary.Cells(k, 2) = hojUsu_Forecast.Cells(c_ini, 3)
            End If
        End If
        
        hojUsu_Forecast.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 2).Value
        c_ini = c_ini + 1
        
    Next k

    'Opciones de Market Clearing Condition (4-Conected)
    Case 4
    
    For k = d_ini To d_fin
            
        Call AsignacionVariablesProcesos
    
        hojUsu_Summary.Cells(k, 2).Formula = "=(((Wood_Industry!J" & i & "*Wood_Industry!K" & i & _
        "*(1-(Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!L" & i & "*Wood_Industry!M" & i & ")" & Chr(10) & _
            "*((Wood_Industry!N" & i & "*Summary!P" & k & ")-((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*(Wood_Industry!N" & j & "*Summary!P" & l & "))))" & Chr(10) & _
        "+((Wood_Industry!P" & i & "*Wood_Industry!Q" & i & ")" & Chr(10) & _
            "*((Wood_Industry!R" & i & "*Wood_Industry!S" & i & ")-((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*(Wood_Industry!R" & j & "*Wood_Industry!S" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!T" & i & "*Wood_Industry!U" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!V" & i & "* Summary!BN" & k & ")/(Wood_Industry!X" & i & "* Wood_Industry!Y" & i & _
            "))-(( Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & _
            ")*( (Wood_Industry!V" & j & "* Summary!BN" & l & ")/(Wood_Industry!X" & j & "* Wood_Industry!Y" & j & "))))))" & Chr(10) & _
        "*Wood_Industry!B" & i & ")+(Summary!B" & l & "*(Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & "))" & Chr(10) & _
        "+((Wood_Industry!Z" & i & "*Wood_Industry!AA" & i & ")*(Wood_Industry!AB" & j & "*Wood_Industry!AC" & j & "))"
        
        If negativeData = 1 Then
            If hojUsu_Summary.Cells(k, 2).Value < 0 Then
                hojUsu_Summary.Cells(k, 2) = hojUsu_Forecast.Cells(c_ini, 3)
            End If
        End If
        
        hojUsu_Forecast.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 2).Value
        c_ini = c_ini + 1
        
    Next k

End Select

'coloca el valor del último año evaluado en los campos de la hoja Summary
hojUsu_Summary.Cells(6, 3).Value = "=Summary!B" & k - 1

End Sub
Sub CONSUMPTION_WOOD_INDUSTRY()

'Asignacion de años seleccionados y opciones escogidas por el usuario
Call AsignacionVariablesOpcionesUsuario

Select Case selectProcess

    'Validation
    Case 2
    
        For k = d_ini To d_fin
                
            Call AsignacionVariablesProcesos
               
            hojUsu_WoodIndustry.Cells(i, 56).Formula = _
                "=(((Wood_Industry!AE" & i & "*Wood_Industry!AF" & i & _
                    "*(1-(Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & ")))" & Chr(10) & _
                "+((Wood_Industry!AG" & i & "*Wood_Industry!AH" & i & ")" & Chr(10) & _
                    "*(((Wood_Industry!AI" & i & "* Wood_Industry!AJ" & i & ")/(Wood_Industry!AK" & i & "* Wood_Industry!AL" & i & _
                    "))-(( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & _
                    ")*( (Wood_Industry!AI" & j & "* Wood_Industry!AJ" & j & ")/(Wood_Industry!AK" & j & "* Wood_Industry!AL" & j & ")))))" & Chr(10) & _
                "+((Wood_Industry!AM" & i & "*Wood_Industry!AN" & i & ")" & Chr(10) & _
                    "*((Wood_Industry!AO" & i & "*Wood_Industry!AP" & i & ")-((Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & _
                    ")*(Wood_Industry!AO" & j & "* Wood_Industry!AP" & j & "))))" & Chr(10) & _
                "+((Wood_Industry!AQ" & i & "*Wood_Industry!AR" & i & ")" & Chr(10) & _
                    "*(((Wood_Industry!AS" & i & "* Wood_Industry!AT" & i & ")/(Wood_Industry!AU" & i & "* Wood_Industry!AV" & i & _
                    "))-(( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & _
                    ")*( (Wood_Industry!AS" & j & "* Wood_Industry!AT" & j & ")/(Wood_Industry!AU" & j & "* Wood_Industry!AV" & j & ")))))" & Chr(10) & _
                "*Wood_Industry!C" & i & ")" & Chr(10) & _
                "+(Wood_Industry!BD" & j & "*( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & "))" & Chr(10) & _
                "+((Wood_Industry!AY" & j & "*Wood_Industry!AZ" & j & ")*( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & ")))"
                    
            hojUsu_Summary.Cells(k, 4).Formula = _
            "=(Wood_Industry!BA" & i & "*Wood_Industry!BB" & i & ")*(Wood_Industry!BC" & i & "*Wood_Industry!BD" & i & ")"
            
            If negativeData = 1 Then
                If hojUsu_Summary.Cells(k, 4).Value < 0 Then
                    hojUsu_Summary.Cells(k, 4) = hojUsu_Forecast.Cells(c_ini, 7)
                End If
            End If
          
            hojUsu_Forecast.Cells(c_ini, 8) = hojUsu_Summary.Cells(k, 4).Value
            c_ini = c_ini + 1
                
        Next k
            
    Case 3, 4
    
        For k = d_ini To d_fin
                
            Call AsignacionVariablesProcesos
        
            hojUsu_WoodIndustry.Cells(i, 56).Formula = _
                "=(((Wood_Industry!AE" & i & "*Wood_Industry!AF" & i & _
                    "*(1-(Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & ")))" & Chr(10) & _
                "+((Wood_Industry!AG" & i & "*Wood_Industry!AH" & i & ")" & Chr(10) & _
                    "*(((Wood_Industry!AI" & i & "* Summary!J" & k & ")/(Wood_Industry!AK" & i & "* Wood_Industry!AL" & i & _
                    "))-(( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & _
                    ")*( (Wood_Industry!AI" & j & "* Summary!J" & l & ")/(Wood_Industry!AK" & j & "* Wood_Industry!AL" & j & ")))))" & Chr(10) & _
                "+((Wood_Industry!AM" & i & "*Wood_Industry!AN" & i & ")" & Chr(10) & _
                    "*((Wood_Industry!AO" & i & "*Wood_Industry!AP" & i & ")-((Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & _
                    ")*(Wood_Industry!AO" & j & "* Wood_Industry!AP" & j & "))))" & Chr(10) & _
                "+((Wood_Industry!AQ" & i & "*Wood_Industry!AR" & i & ")" & Chr(10) & _
                    "*(((Wood_Industry!AS" & i & "* Wood_Industry!AT" & i & ")/(Wood_Industry!AU" & i & "* Wood_Industry!AV" & i & _
                    "))-(( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & _
                    ")*( (Wood_Industry!AS" & j & "* Wood_Industry!AT" & j & ")/(Wood_Industry!AU" & j & "* Wood_Industry!AV" & j & ")))))" & Chr(10) & _
                "*Wood_Industry!C" & i & ")" & Chr(10) & _
                "+(Wood_Industry!BD" & j & "*( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & "))" & Chr(10) & _
                "+((Wood_Industry!AY" & j & "*Wood_Industry!AZ" & j & ")*( Wood_Industry!AW" & i & "*Wood_Industry!AX" & i & ")))"
                    
            hojUsu_Summary.Cells(k, 4).Formula = _
            "=(Wood_Industry!AY" & i & "*Wood_Industry!AZ" & i & ")*(Wood_Industry!BA" & i & "*Wood_Industry!BB" & i & ")"
            
                If negativeData = 1 Then
                    If hojUsu_Summary.Cells(k, 4).Value < 0 Then
                        hojUsu_Summary.Cells(k, 4) = hojUsu_Forecast.Cells(c_ini, 7)
                    End If
                End If
            
            hojUsu_Forecast.Cells(c_ini, 9) = hojUsu_Summary.Cells(k, 4).Value
            c_ini = c_ini + 1
            
        Next k

End Select

hojUsu_Summary.Cells(7, 3).Value = "=Summary!D" & k - 1

End Sub
Sub EXPORTS_WOOD_INDUSTRY()

'Asignacion de años seleccionados y opciones escogidas por el usuario
Call AsignacionVariablesOpcionesUsuario

Select Case selectProcess

    'Validation
    Case 2
    
        For k = d_ini To d_fin
                
            Call AsignacionVariablesProcesos
                
            hojUsu_Summary.Cells(k, 6).Formula = _
                "=(((Wood_Industry!BF" & i & "*Wood_Industry!BG" & i & _
                    "*(1-(Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & ")))" & Chr(10) & _
                "+((Wood_Industry!BH" & i & "*Wood_Industry!BI" & i & ")" & Chr(10) & _
                    "*((Wood_Industry!BJ" & i & "*Wood_Industry!BK" & i & ")-((Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & _
                    ")*(Wood_Industry!BJ" & j & "* Wood_Industry!BK" & j & "))))" & Chr(10) & _
                "+((Wood_Industry!BL" & i & "*Wood_Industry!BM" & i & ")" & Chr(10) & _
                    "*(((Wood_Industry!BN" & i & "* Wood_Industry!BO" & i & ")/(Wood_Industry!BP" & i & "* Wood_Industry!BQ" & i & _
                    "))-(( Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & _
                    ")*( (Wood_Industry!BN" & j & "* Wood_Industry!BO" & j & ")/(Wood_Industry!BP" & j & "* Wood_Industry!BQ" & j & ")))))" & Chr(10) & _
                "*Wood_Industry!D" & i & ")" & Chr(10) & _
                "+(Summary!F" & l & Chr(10) & _
                "*( Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & "))" & Chr(10) & _
                "+((Wood_Industry!BT" & j & "*Wood_Industry!BU" & j & ")*( Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & ")))"
                    
                    
            If negativeData = 1 Then
                If hojUsu_Summary.Cells(k, 6).Value < 0 Then
                    hojUsu_Summary.Cells(k, 6) = hojUsu_Forecast.Cells(c_ini, 11)
                End If
            End If
        
            hojUsu_Forecast.Cells(c_ini, 12) = hojUsu_Summary.Cells(k, 6).Value
            c_ini = c_ini + 1
                
        Next k
            
    Case 3, 4
    
    For k = d_ini To d_fin
            
        Call AsignacionVariablesProcesos
    
        hojUsu_Summary.Cells(k, 6).Formula = _
            "=(((Wood_Industry!BF" & i & "*Wood_Industry!BG" & i & _
                "*(1-(Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & ")))" & Chr(10) & _
            "+((Wood_Industry!BH" & i & "*Wood_Industry!BI" & i & ")" & Chr(10) & _
                "*((Wood_Industry!BJ" & i & "*Wood_Industry!BK" & i & ")-((Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & _
                ")*(Wood_Industry!BJ" & j & "* Wood_Industry!BK" & j & "))))" & Chr(10) & _
            "+((Wood_Industry!BL" & i & "*Wood_Industry!BM" & i & ")" & Chr(10) & _
                "*(((Wood_Industry!BN" & i & "* Summary!L" & k & ")/(Wood_Industry!BP" & i & "* Wood_Industry!BQ" & i & _
                "))-(( Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & _
                ")*( (Wood_Industry!BN" & j & "* Summary!L" & l & ")/(Wood_Industry!BP" & j & "* Wood_Industry!BQ" & j & ")))))" & Chr(10) & _
            "*Wood_Industry!D" & i & ")" & Chr(10) & _
            "+(Summary!F" & l & Chr(10) & _
            "*( Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & "))" & Chr(10) & _
            "+((Wood_Industry!BT" & j & "*Wood_Industry!BU" & j & ")*( Wood_Industry!BR" & i & "*Wood_Industry!BS" & i & ")))"
                
        If negativeData = 1 Then
            If hojUsu_Summary.Cells(k, 6).Value < 0 Then
                hojUsu_Summary.Cells(k, 6) = hojUsu_Forecast.Cells(c_ini, 11)
            End If
        End If
    
        hojUsu_Forecast.Cells(c_ini, 13) = hojUsu_Summary.Cells(k, 6).Value
        c_ini = c_ini + 1
        
    Next k

End Select

hojUsu_Summary.Cells(8, 3).Value = "=Summary!F" & k - 1

End Sub
Sub IMPORTS_WOOD_INDUSTRY()

'Asignacion de años seleccionados y opciones escogidas por el usuario
Call AsignacionVariablesOpcionesUsuario

Select Case selectProcess

    'Validation
    Case 2

        For k = d_ini To d_fin
        
            Call AsignacionVariablesProcesos
        
            hojUsu_Summary.Cells(k, 8).Formula = _
                "=(((Wood_Industry!BU" & i & "*Wood_Industry!BV" & i & _
                    "*(1-(Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & ")))" & Chr(10) & _
                "+((Wood_Industry!BW" & i & "*Wood_Industry!BX" & i & ")" & Chr(10) & _
                    "*((Wood_Industry!BY" & i & "* Wood_Industry!BZ" & i & ")-((Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & _
                    ")*(Wood_Industry!BY" & j & "* Wood_Industry!BZ" & j & "))))" & Chr(10) & _
                "+((Wood_Industry!CA" & i & "*Wood_Industry!CB" & i & ")" & Chr(10) & _
                    "*(((Wood_Industry!CC" & i & "* Wood_Industry!CD" & i & ")/(Wood_Industry!CE" & i & "* Wood_Industry!CF" & i & _
                    "))-((Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & _
                    ")*( (Wood_Industry!CC" & j & "* Wood_Industry!CD" & j & ")/(Wood_Industry!CE" & j & "* Wood_Industry!CF" & j & ")))))" & Chr(10) & _
                "+((Wood_Industry!CG" & i & "*Wood_Industry!CH" & i & ")" & Chr(10) & _
                    "*((Wood_Industry!CI" & i & "* Wood_Industry!CJ" & i & _
                    ")-((Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & _
                    ")*(Wood_Industry!CI" & j & "* Wood_Industry!CJ" & j & "))))" & Chr(10) & _
                "*Wood_Industry!E" & i & ")" & Chr(10) & _
                "+(Summary!H" & l & "*(Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & "))" & Chr(10) & _
                "+((Wood_Industry!CM" & j & "*Wood_Industry!CN" & j & ")*( Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & ")))"

            If negativeData = 1 Then
                If hojUsu_Summary.Cells(k, 8).Value < 0 Then
                    hojUsu_Summary.Cells(k, 8) = hojUsu_Forecast.Cells(c_ini, 15)
                End If
            End If
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 8).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 8) = hojUsu_Forecast.Cells(c_ini, 15)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 8).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 8) = 0
        
        End If
    
    End Select
            
    hojUsu_Forecast.Cells(c_ini, 16) = hojUsu_Summary.Cells(k, 8).Value
    c_ini = c_ini + 1
        
Next k
        
Case 3, 4

For k = d_ini To d_fin
        
Call AsignacionVariablesProcesos

    hojUsu_Summary.Cells(k, 8).Formula = _
        "=(((Wood_Industry!BU" & i & "*Wood_Industry!BV" & i & _
            "*(1-(Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!BW" & i & "*Wood_Industry!BX" & i & ")" & Chr(10) & _
            "*((Wood_Industry!BY" & i & "* Summary!D" & k & _
            ")-((Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & _
            ")*(Wood_Industry!BY" & j & "* Summary!D" & l & "))))" & Chr(10) & _
        "+((Wood_Industry!CA" & i & "*Wood_Industry!CB" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!CC" & i & "* Summary!J" & k & _
            ")/(Wood_Industry!CE" & i & "* Summary!N" & k & _
            "))-((Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & _
            ")*( (Wood_Industry!CC" & j & "* Summary!J" & l & _
            ")/(Wood_Industry!CE" & j & "* Summary!N" & l & ")))))" & Chr(10) & _
        "+((Wood_Industry!CG" & i & "*Wood_Industry!CH" & i & ")" & Chr(10) & _
            "*((Wood_Industry!CI" & i & "* Wood_Industry!CJ" & i & _
            ")-((Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & _
            ")*(Wood_Industry!CI" & j & "* Wood_Industry!CJ" & j & "))))" & Chr(10) & _
        "*Wood_Industry!E" & i & ")" & Chr(10) & _
        "+(Summary!H" & l & "*(Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & "))" & Chr(10) & _
        "+((Wood_Industry!CM" & j & "*Wood_Industry!CN" & j & ")*( Wood_Industry!CK" & i & "*Wood_Industry!CL" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 8).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 8) = hojUsu_Forecast.Cells(c_ini, 15)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 8).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 8) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 17) = hojUsu_Summary.Cells(k, 8).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(9, 3).Value = "=Summary!H" & k - 1

End Sub
Sub PRICE_OF_CONSUMPTION_WOOD_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 10).Formula = _
        "=(((Wood_Industry!CP" & i & "*Wood_Industry!CQ" & i & _
            "*(1-(Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!CR" & i & "*Wood_Industry!CS" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!CT" & i & "* Wood_Industry!CU" & i & _
            ")/(Wood_Industry!CV" & i & "* Wood_Industry!CW" & i & _
            "))-(( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & _
            ")*( (Wood_Industry!CT" & j & "* Wood_Industry!CU" & j & _
            ")/(Wood_Industry!CV" & j & "* Wood_Industry!CW" & j & ")))))" & Chr(10) & _
        "+((Wood_Industry!CX" & i & "*Wood_Industry!CY" & i & ")" & Chr(10) & _
            "*((Wood_Industry!CZ" & i & "* Wood_Industry!DA" & i & _
            ")-(( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & _
            ")*(Wood_Industry!CZ" & j & "* Wood_Industry!DA" & j & "))))" & Chr(10) & _
        "*Wood_Industry!F" & i & ")" & Chr(10) & _
        "+(Summary!J" & l & "*( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & "))" & Chr(10) & _
        "+((Wood_Industry!DD" & j & "*Wood_Industry!DE" & j & ")*( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 10).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 10) = hojUsu_Forecast.Cells(c_ini, 19)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 10).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 10) = 0
        
        End If
    
    End Select
    
'resultados en el forecast
    hojUsu_Forecast.Cells(c_ini, 20) = hojUsu_Summary.Cells(k, 10).Value
'resultados en la hoja de set de precios
    hojUsu_SetPricesWoodIndustry.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 10)
    
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

    hojUsu_Summary.Cells(k, 10).Formula = _
        "=(((Wood_Industry!CP" & i & "*Wood_Industry!CQ" & i & _
            "*(1-(Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!CR" & i & "*Wood_Industry!CS" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!CT" & i & "* Summary!B" & k & _
            ")/(Wood_Industry!CV" & i & "* Wood_Industry!CW" & i & _
            "))-(( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & _
            ")*( (Wood_Industry!CT" & j & "* Summary!B" & l & _
            ")/(Wood_Industry!CV" & j & "* Wood_Industry!CW" & j & ")))))" & Chr(10) & _
        "+((Wood_Industry!CX" & i & "*Wood_Industry!CY" & i & ")" & Chr(10) & _
            "*((Wood_Industry!CZ" & i & "* Summary!J" & l & _
            ")-(( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & _
            ")*(Wood_Industry!CZ" & j & "* Summary!J" & m & "))))" & Chr(10) & _
        "*Wood_Industry!F" & i & ")" & Chr(10) & _
        "+(Summary!J" & l & "*( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & "))" & Chr(10) & _
        "+((Wood_Industry!DD" & j & "*Wood_Industry!DE" & j & ")*( Wood_Industry!DB" & i & "*Wood_Industry!DC" & i & ")))"

            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 10).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 10) = hojUsu_Forecast.Cells(c_ini, 19)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 10).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 10) = 0
        
        End If

    End Select
            
'resultados en el forecast
    hojUsu_Forecast.Cells(c_ini, 21) = hojUsu_Summary.Cells(k, 10).Value
'resultados en la hoja de set de precios
    hojUsu_SetPricesWoodIndustry.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 10)
    
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(10, 3).Value = "=Summary!J" & k - 1

End Sub
Sub PRICE_OF_EXPORTS_WOOD_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 12).Formula = _
        "=(((Wood_Industry!DG" & i & "*Wood_Industry!DH" & i & _
            "*(1-(Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!DI" & i & "*Wood_Industry!DJ" & i & ")" & Chr(10) & _
            "*((Wood_Industry!DK" & i & "* Wood_Industry!DL" & i & _
            ")-(( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & _
            ")*(Wood_Industry!DK" & j & "* Wood_Industry!DL" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!DM" & i & "*Wood_Industry!DN" & i & ")" & Chr(10) & _
            "*((Wood_Industry!DO" & i & "* Wood_Industry!DP" & i & _
            ")-(( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & _
            ")*(Wood_Industry!DO" & j & "* Wood_Industry!DP" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!DQ" & i & "*Wood_Industry!DR" & i & ")" & Chr(10) & _
            "*((Wood_Industry!DS" & i & "* Wood_Industry!DT" & i & _
            ")-(( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & _
            ")*(Wood_Industry!DS" & j & "* Wood_Industry!DT" & j & "))))" & Chr(10) & _
        "*Wood_Industry!G" & i & ")" & Chr(10) & _
        "+(Summary!L" & l & "*( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & "))" & Chr(10) & _
        "+((Wood_Industry!DW" & j & "*Wood_Industry!DX" & j & ")*( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 12).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 12) = hojUsu_Forecast.Cells(c_ini, 23)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 12).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 12) = 0
        
        End If
    
    End Select

'resultados en el forecast
    hojUsu_Forecast.Cells(c_ini, 24) = hojUsu_Summary.Cells(k, 12).Value
'resultados en la hoja de set de precios
    hojUsu_SetPricesWoodIndustry.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 12)
    
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

    hojUsu_Summary.Cells(k, 12).Formula = _
        "=(((Wood_Industry!DG" & i & "*Wood_Industry!DH" & i & _
            "*(1-(Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!DI" & i & "*Wood_Industry!DJ" & i & ")" & Chr(10) & _
            "*((Wood_Industry!DK" & i & "* Summary!L" & l & _
            ")-(( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & _
            ")*(Wood_Industry!DK" & j & "* Summary!L" & m & "))))" & Chr(10) & _
        "+((Wood_Industry!DM" & i & "*Wood_Industry!DN" & i & ")" & Chr(10) & _
            "*((Wood_Industry!DO" & i & "* Wood_Industry!DP" & i & _
            ")-(( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & _
            ")*(Wood_Industry!DO" & j & "* Wood_Industry!DP" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!DQ" & i & "*Wood_Industry!DR" & i & ")" & Chr(10) & _
            "*((Wood_Industry!DS" & i & "* Wood_Industry!DT" & i & _
            ")-(( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & _
            ")*(Wood_Industry!DS" & j & "* Wood_Industry!DT" & j & "))))" & Chr(10) & _
        "*Wood_Industry!G" & i & ")" & Chr(10) & _
        "+(Summary!L" & l & "*( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & "))" & Chr(10) & _
        "+((Wood_Industry!DW" & j & "*Wood_Industry!DX" & j & ")*( Wood_Industry!DU" & i & "*Wood_Industry!DV" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 12).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 12) = hojUsu_Forecast.Cells(c_ini, 23)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 12).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 12) = 0
        
        End If
    
    End Select

'resultados en el forecast
    hojUsu_Forecast.Cells(c_ini, 25) = hojUsu_Summary.Cells(k, 12).Value
'resultados en la hoja de set de precios
    hojUsu_SetPricesWoodIndustry.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 12)
    
    c_ini = c_ini + 1

   
Next k

End Select

hojUsu_Summary.Cells(11, 3).Value = "=Summary!L" & k - 1

End Sub
Sub PRICE_OF_IMPORT_WOOD_INDUSTRY()

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
        
    hojUsu_Summary.Cells(k, 14).Formula = _
        "=(((Wood_Industry!DZ" & i & "*Wood_Industry!EA" & i & _
            "*(1-(Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!EB" & i & "*Wood_Industry!EC" & i & ")" & Chr(10) & _
            "*((Wood_Industry!ED" & i & "* Wood_Industry!EE" & i & _
            ")-(( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & _
            ")*(Wood_Industry!ED" & j & "* Wood_Industry!EE" & j & "))))" & Chr(10) & _
        "+((Wood_Industry!EF" & i & "*Wood_Industry!EG" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!EH" & i & "* Wood_Industry!EI" & i & _
            ")/(Wood_Industry!EJ" & i & "* Wood_Industry!EK" & i & _
            "))-(( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & _
            ")*( (Wood_Industry!EH" & j & "* Wood_Industry!EI" & j & _
            ")/(Wood_Industry!EJ" & j & "* Wood_Industry!EK" & j & ")))))" & Chr(10) & _
        "+((Wood_Industry!EL" & i & "*Wood_Industry!EM" & i & ")" & Chr(10) & _
            "*((Wood_Industry!EN" & i & "* Wood_Industry!EO" & i & _
            ")-(( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & _
            ")*(Wood_Industry!EN" & j & "* Wood_Industry!EO" & j & "))))" & Chr(10) & _
        "*Wood_Industry!H" & i & ")" & Chr(10) & _
        "+(Summary!N" & l & "*( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & "))" & Chr(10) & _
        "+((Wood_Industry!ER" & j & "*Wood_Industry!ES" & j & ")*( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & ")))"

    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 14).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 14) = hojUsu_Forecast.Cells(c_ini, 27)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 14).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 14) = 0
        
        End If
    
    End Select
    
'resultados en el forecast
    hojUsu_Forecast.Cells(c_ini, 28) = hojUsu_Summary.Cells(k, 14).Value
'resultados en la hoja de set de precios
    hojUsu_SetPricesWoodIndustry.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 14)
    
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

    hojUsu_Summary.Cells(k, 14).Formula = _
        "=(((Wood_Industry!DZ" & i & "*Wood_Industry!EA" & i & _
            "*(1-(Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & ")))" & Chr(10) & _
        "+((Wood_Industry!EB" & i & "*Wood_Industry!EC" & i & ")" & Chr(10) & _
            "*((Wood_Industry!ED" & i & "* Summary!N" & l & _
            ")-((Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & _
            ")*(Wood_Industry!ED" & j & "* Summary!N" & m & "))))" & Chr(10) & _
        "+((Wood_Industry!EF" & i & "*Wood_Industry!EG" & i & ")" & Chr(10) & _
            "*(((Wood_Industry!EH" & i & "* Summary!J" & k & _
            ")/(Wood_Industry!EJ" & i & "* Summary!J" & l & _
            "))-(( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & _
            ")*( (Wood_Industry!EH" & j & "* Summary!J" & l & _
            ")/(Wood_Industry!EJ" & j & "* Summary!J" & m & ")))))" & Chr(10) & _
        "+((Wood_Industry!EL" & i & "*Wood_Industry!EM" & i & ")" & Chr(10) & _
            "*((Wood_Industry!EN" & i & "* Wood_Industry!EO" & i & _
            ")-(( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & _
            ")*(Wood_Industry!EN" & j & "* Wood_Industry!EO" & j & "))))" & Chr(10) & _
        "*Wood_Industry!H" & i & ")" & Chr(10) & _
        "+(Summary!N" & l & "*( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & "))" & Chr(10) & _
        "+((Wood_Industry!ER" & j & "*Wood_Industry!ES" & j & ")*( Wood_Industry!EP" & i & "*Wood_Industry!EQ" & i & ")))"
   
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 14).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 14) = hojUsu_Forecast.Cells(c_ini, 27)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 14).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 14) = 0
        
        End If
        
    End Select
    
'resultados en el forecast
    hojUsu_Forecast.Cells(c_ini, 29) = hojUsu_Summary.Cells(k, 14).Value
'resultados en la hoja de set de precios
    hojUsu_SetPricesWoodIndustry.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 14)
    
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(11, 3).Value = "=Summary!N" & k - 1

End Sub
