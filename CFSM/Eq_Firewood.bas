Attribute VB_Name = "Eq_Firewood"
Sub SUPPLY_FIREWOOD()

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
                    
            hojUsu_Summary.Cells(k, 76).Formula = "=(((Firewood!J" & i & "*Firewood!K" & i & _
            "*(1-(Firewood!R" & i & "*Firewood!S" & i & ")))" & Chr(10) & _
            "+((Firewood!L" & i & "*Firewood!M" & i & ")" & Chr(10) & _
                "*(((Firewood!N" & i & "* Firewood!O" & i & ")/(Firewood!P" & i & "* Firewood!Q" & i & _
                "))-(( Firewood!R" & i & "*Firewood!S" & i & _
                ")*( (Firewood!N" & j & "* Firewood!O" & j & ")/(Firewood!P" & j & "* Firewood!Q" & j & "))))))" & Chr(10) & _
            "*Firewood!B" & i & _
            ")+(Summary!BX" & l & "*(Firewood!R" & i & "*Firewood!S" & i & "))" & Chr(10) & _
            "+((Firewood!R" & i & "*Firewood!S" & i & ")*(Firewood!T" & j & "*Firewood!U" & j & "))"
            
            Select Case g
            
                Case 1
                
                    If hojUsu_Summary.Cells(k, 76).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 76) = hojUsu_Forecast.Cells(c_ini, 119)
                    
                    End If
                
                Case 2
                    
                    If hojUsu_Summary.Cells(k, 76).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 76) = 0
                    
                    End If
    
            End Select
            
            hojUsu_Forecast.Cells(c_ini, 120) = hojUsu_Summary.Cells(k, 76).Value
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
        
            hojUsu_Summary.Cells(k, 76).Formula = "=(((Firewood!J" & i & "*Firewood!K" & i & _
            "*(1-(Firewood!R" & i & "*Firewood!S" & i & ")))" & Chr(10) & _
            "+((Firewood!L" & i & "*Firewood!M" & i & ")" & Chr(10) & _
                "*(((Firewood!N" & i & "* Summary!CL" & k & ")/(Firewood!P" & i & "* Firewood!Q" & i & _
                "))-(( Firewood!R" & i & "*Firewood!S" & i & _
                ")*( (Firewood!N" & j & "* Summary!CL" & l & ")/(Firewood!P" & j & "* Firewood!Q" & j & "))))))" & Chr(10) & _
            "*Firewood!B" & i & _
            ")+(Summary!BX" & l & "*(Firewood!R" & i & "*Firewood!S" & i & "))" & Chr(10) & _
            "+((Firewood!R" & i & "*Firewood!S" & i & ")*(Firewood!T" & j & "*Firewood!U" & j & "))"
        
            Select Case g
            
                Case 1
                
                    If hojUsu_Summary.Cells(k, 76).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 76) = hojUsu_Forecast.Cells(c_ini, 119)
                    
                    End If
                
                Case 2
                    
                    If hojUsu_Summary.Cells(k, 76).Value < 0 Then
                    
                    hojUsu_Summary.Cells(k, 76) = 0
                    
                    End If
                    
                Case 3
            
            End Select
        
            hojUsu_Forecast.Cells(c_ini, 121) = hojUsu_Summary.Cells(k, 76).Value
            c_ini = c_ini + 1
            
        Next k
    
    End Select
    
    hojUsu_Summary.Cells(6, 23).Value = "=Summary!BX" & k - 1

End Sub
Sub CONSUMPTION_FIREWOOD()

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
       
    hojUsu_Firewood.Cells(i, 46).Formula = _
        "=(((Firewood!W" & i & "*Firewood!X" & i & "*(1-(Firewood!AM" & i & "*Firewood!AN" & i & ")))" & Chr(10) & _
        "+((Firewood!Y" & i & "*Firewood!Z" & i & ")" & Chr(10) & _
            "*((Firewood!AA" & i & "* Firewood!AB" & i & ")-(( Firewood!AM" & i & "*Firewood!AN" & i & _
            ")*(Firewood!AA" & j & "* Firewood!AB" & j & "))))" & Chr(10) & _
        "+((Firewood!AC" & i & "*Firewood!AD" & i & ")" & Chr(10) & _
            "*((Firewood!AE" & i & "* Firewood!AF" & i & ")-(( Firewood!AM" & i & "*Firewood!AN" & i & _
            ")*(Firewood!AE" & j & "* Firewood!AF" & j & "))))" & Chr(10) & _
        "+((Firewood!AG" & i & "*Firewood!AH" & i & ")" & Chr(10) & _
            "*(((Firewood!AI" & i & "* Firewood!AJ" & i & ")/(Firewood!AK" & i & "* Firewood!AL" & i & _
            "))-(( Firewood!AM" & i & "*Firewood!AN" & i & _
            ")*( (Firewood!AI" & j & "* Firewood!AJ" & j & ")/(Firewood!AK" & j & "* Firewood!AL" & j & ")))))" & Chr(10) & _
        "*Firewood!C" & i & ")" & Chr(10) & _
        "+(Firewood!AT" & j & "*( Firewood!AM" & i & "*Firewood!AN" & i & "))" & Chr(10) & _
        "+((Firewood!AO" & j & "*Firewood!AP" & j & ")*( Firewood!AM" & i & "*Firewood!AN" & i & ")))"
            
    hojUsu_Summary.Cells(k, 78).Formula = _
    "=(Firewood!AQ" & i & "*Firewood!AR" & i & ")*(Firewood!AS" & i & "*Firewood!AT" & i & ")"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 78).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 78) = hojUsu_Forecast.Cells(c_ini, 123)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 78).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 78) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 124) = hojUsu_Summary.Cells(k, 78).Value
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

'    hojUsu_Firewood.Cells(i, 46).Formula = _
'        "=(((Firewood!W" & i & "*Firewood!X" & i & "*(1-(Firewood!AM" & i & "*Firewood!AN" & i & ")))" & Chr(10) & _
'        "+((Firewood!Y" & i & "*Firewood!Z" & i & ")" & Chr(10) & _
'            "*((Firewood!AA" & i & "* Firewood!AB" & i & ")-(( Firewood!AM" & i & "*Firewood!AN" & i & _
'            ")*(Firewood!AA" & j & "* Firewood!AB" & j & "))))" & Chr(10) & _
'        "+((Firewood!AC" & i & "*Firewood!AD" & i & ")" & Chr(10) & _
'            "*((Firewood!AE" & i & "* Summary!CF" & k & ")-(( Firewood!AM" & i & "*Firewood!AN" & i & _
'            ")*(Firewood!AE" & j & "* Summary!CF" & l & "))))" & Chr(10) & _
'        "+((Firewood!AG" & i & "*Firewood!AH" & i & ")" & Chr(10) & _
'            "*(((Firewood!AI" & i & "* Firewood!AJ" & i & ")/(Firewood!AK" & i & "* Firewood!AL" & i & _
'            "))-(( Firewood!AM" & i & "*Firewood!AN" & i & _
'            ")*( (Firewood!AI" & j & "* Firewood!AJ" & j & ")/(Firewood!AK" & j & "* Firewood!AL" & j & ")))))" & Chr(10) & _
'        "*Firewood!C" & i & ")" & Chr(10) & _
'        "+(Firewood!AT" & j & "*( Firewood!AM" & i & "*Firewood!AN" & i & "))" & Chr(10) & _
'        "+((Firewood!AO" & j & "*Firewood!AP" & j & ")*( Firewood!AM" & i & "*Firewood!AN" & i & ")))"
'
'    hojUsu_Summary.Cells(k, 78).Formula = _
'    "=(Firewood!AQ" & i & "*Firewood!AR" & i & ")*(Firewood!AS" & i & "*Firewood!AT" & i & ")"
    
    hojUsu_Firewood.Cells(i, 46).Formula = _
        "=(((Firewood!W" & i & "*Firewood!X" & i & "*(1-(Firewood!AM" & i & "*Firewood!AN" & i & ")))" & Chr(10) & _
        "+((Firewood!Y" & i & "*Firewood!Z" & i & ")" & Chr(10) & _
            "*((Firewood!AA" & i & "* Firewood!AB" & i & ")-(( Firewood!AM" & i & "*Firewood!AN" & i & _
            ")*(Firewood!AA" & j & "* Firewood!AB" & j & "))))" & Chr(10) & _
        "+((Firewood!AC" & i & "*Firewood!AD" & i & ")" & Chr(10) & _
            "*((Firewood!AE" & i & "* Summary!CF" & k & ")-(( Firewood!AM" & i & "*Firewood!AN" & i & _
            ")*(Firewood!AE" & j & "* Summary!CF" & l & "))))" & Chr(10) & _
        "+((Firewood!AG" & i & "*Firewood!AH" & i & ")" & Chr(10) & _
            "*(((Firewood!AI" & i & "* Firewood!AJ" & i & ")/(Firewood!AK" & i & "* Firewood!AL" & i & _
            "))-(( Firewood!AM" & i & "*Firewood!AN" & i & _
            ")*( (Firewood!AI" & j & "* Firewood!AJ" & j & ")/(Firewood!AK" & j & "* Firewood!AL" & j & ")))))" & Chr(10) & _
        "*Firewood!C" & i & ")" & Chr(10) & _
        "+(Firewood!AT" & j & "*( Firewood!AM" & i & "*Firewood!AN" & i & "))" & Chr(10) & _
        "+((Firewood!AO" & j & "*Firewood!AP" & j & ")*( Firewood!AM" & i & "*Firewood!AN" & i & ")))"
            
    hojUsu_Summary.Cells(k, 78).Formula = _
    "=(Firewood!AQ" & i & "*Firewood!AR" & i & ")*(Firewood!AS" & i & "*Firewood!AT" & i & ")"
    
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 78).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 78) = hojUsu_Forecast.Cells(c_ini, 123)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 78).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 78) = 0
        
        End If
    
    Case 3
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 125) = hojUsu_Summary.Cells(k, 78).Value
    c_ini = c_ini + 1
    
Next k

End Select

hojUsu_Summary.Cells(7, 23).Value = "=Summary!BZ" & k - 1

End Sub
Sub EXPORTS_FIREWOOD()



End Sub
Sub IMPORTS_FIREWOOD()



End Sub
Sub PRICE_OF_CONSUMPTION_FIREWOOD()

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
        
    hojUsu_Summary.Cells(k, 84).Formula = _
        "=(((Firewood!AV" & i & "*Firewood!AW" & i & "*(1-(Firewood!BH" & i & "*Firewood!BI" & i & ")))" & Chr(10) & _
        "+((Firewood!AX" & i & "*Firewood!AY" & i & ")" & Chr(10) & _
            "*((Firewood!AZ" & i & "* Firewood!BA" & i & ")-(( Firewood!BH" & i & "*Firewood!BI" & i & _
            ")*(Firewood!AZ" & j & "* Firewood!BA" & j & ")))" & Chr(10) & _
        "+((Firewood!BB" & i & "*Firewood!BC" & i & ")" & Chr(10) & _
            "*(((Firewood!BD" & i & "* Firewood!BE" & i & ")/(Firewood!BF" & i & "* Firewood!BG" & i & _
            "))-(( Firewood!BH" & i & "*Firewood!BI" & i & _
            ")*( (Firewood!BD" & j & "* Firewood!BE" & j & ")/(Firewood!BF" & j & "* Firewood!BG" & j & "))))))" & Chr(10) & _
        "*Firewood!D" & i & ")" & Chr(10) & _
        "+(Summary!CF" & l & "*( Firewood!BH" & i & "*Firewood!BI" & i & "))" & Chr(10) & _
        "+((Firewood!BJ" & j & "*Firewood!BK" & j & ")*( Firewood!BH" & i & "*Firewood!BI" & i & ")))"
            
    Select Case g
    
    Case 1
    
        If hojUsu_Summary.Cells(k, 84).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 84) = hojUsu_Forecast.Cells(c_ini, 127)
        
        End If
    
    Case 2
        
        If hojUsu_Summary.Cells(k, 84).Value < 0 Then
        
        hojUsu_Summary.Cells(k, 84) = 0
        
        End If
    
    End Select
    
    hojUsu_Forecast.Cells(c_ini, 128) = hojUsu_Summary.Cells(k, 84).Value
    hojUsu_SetPricesFirewood.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 84)
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
    
        hojUsu_Summary.Cells(k, 84).Formula = _
            "=(((Firewood!AV" & i & "*Firewood!AW" & i & "*(1-(Firewood!BH" & i & "*Firewood!BI" & i & ")))" & Chr(10) & _
            "+((Firewood!AX" & i & "*Firewood!AY" & i & ")" & Chr(10) & _
                "*((Firewood!AZ" & i & "* Summary!BX" & k & ")-(( Firewood!BH" & i & "*Firewood!BI" & i & _
                ")*(Firewood!AZ" & j & "* Summary!BX" & l & ")))" & Chr(10) & _
            "+((Firewood!BB" & i & "*Firewood!BC" & i & ")" & Chr(10) & _
                "*(((Firewood!BD" & i & "* Summary!CF" & l & ")/(Firewood!BF" & i & "* Firewood!BG" & i & _
                "))-(( Firewood!BH" & i & "*Firewood!BI" & i & _
                ")*( (Firewood!BD" & j & "* Summary!CF" & m & ")/(Firewood!BF" & j & "* Firewood!BG" & j & "))))))" & Chr(10) & _
            "*Firewood!D" & i & ")" & Chr(10) & _
            "+(Summary!CF" & l & "*( Firewood!BH" & i & "*Firewood!BI" & i & "))" & Chr(10) & _
            "+((Firewood!BJ" & j & "*Firewood!BK" & j & ")*( Firewood!BH" & i & "*Firewood!BI" & i & ")))"
    
        Select Case g
        
        Case 1
        
            If hojUsu_Summary.Cells(k, 84).Value < 0 Then
            
            hojUsu_Summary.Cells(k, 84) = hojUsu_Forecast.Cells(c_ini, 127)
            
            End If
        
        Case 2
            
            If hojUsu_Summary.Cells(k, 84).Value < 0 Then
            
            hojUsu_Summary.Cells(k, 84) = 0
            
            End If
            
        Case 3
        
        End Select
                
        hojUsu_Forecast.Cells(c_ini, 129) = hojUsu_Summary.Cells(k, 84).Value
        hojUsu_SetPricesFirewood.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 84)
        c_ini = c_ini + 1
        
    Next k
    
'    hojUsu_Firewood.Activate
'    Range("BM9:BM21").Copy
'    hojUsu_Summary.Activate
'    Range("CF40").Activate
'    ActiveSheet.Paste

End Select

hojUsu_Summary.Cells(10, 23).Value = "=Summary!CF" & k - 1

End Sub
Sub PRICE_OF_EXPORTS_FIREWOOD()



End Sub
Sub PRICE_OF_IMPORT_FIREWOOD()



End Sub
Sub SUPPLY_FINAL_RURAL_CONSUMPTION()



End Sub
Sub CONSUMPTION_FINAL_RURAL_CONSUMPTION()



End Sub
Sub EXPORTS_FINAL_RURAL_CONSUMPTION()



End Sub
Sub IMPORTS_FINAL_RURAL_CONSUMPTION()



End Sub
Sub PRICE_OF_CONSUMPTION_FINAL_RURAL_CONSUMPTION()



End Sub
Sub PRICE_OF_EXPORTS_FINAL_RURAL_CONSUMPTION()


End Sub
Sub PRICE_OF_IMPORT_FINAL_RURAL_CONSUMPTION()



End Sub
