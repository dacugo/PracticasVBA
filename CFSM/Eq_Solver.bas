Attribute VB_Name = "Eq_Solver"
Sub RUN_SOLVER()

'Call RESET

'hojUsu_SystemOptions.Range("SelectProcess") = 1

Call CALL_EQUATIONS

hojUsu_Summary.Activate

Dim market As String

market = hojUsu_SystemOptions.Range("MarketsInputs").Value

Select Case market

    Case "Wood_Industry"
            
        Call SOLVER_WOOD_INDUSTRY
            
    Case "Furniture_Industry"
    
        Call SOLVER_FURNITURE_INDUSTRY
            
    Case "Pulp_Paper_Industry"
    
        Call SOLVER_PULP_PAPER_INDUSTRY
            
    Case "Wood_Industrial"
    
        Call SOLVER_WOOD_INDUSTRIAL
            
    Case "Firewood"
    
        Call SOLVER_FIREWOOD
            
    Case "All"
    
        Call SOLVER_WOOD_INDUSTRY
        Call SOLVER_FURNITURE_INDUSTRY
        Call SOLVER_PULP_PAPER_INDUSTRY
        Call SOLVER_WOOD_INDUSTRIAL
        Call SOLVER_FIREWOOD
                   
End Select

hojUsu_SystemOptions.Activate

End Sub
Sub SOLVER_WOOD_INDUSTRY()

Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRangeSolver")
b = hojUsu_SystemOptions.Range("FinalYearRangeSolver")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de engine y verificación de interacción de las Eq

metodoSolver = hojUsu_SystemOptions.Range("IterationMethod")
'hojUsu_SystemOptions.Range("SelectProcess") = 2
tipoSolver = hojUsu_SystemOptions.Range("VariablesSolver")
setPricesValue = hojUsu_SystemOptions.Range("OriginForVariablesTwo")

Call SUPPLY_WOOD_INDUSTRY
Call CONSUMPTION_WOOD_INDUSTRY
Call EXPORTS_WOOD_INDUSTRY
Call IMPORTS_WOOD_INDUSTRY
Call PRICE_OF_CONSUMPTION_WOOD_INDUSTRY
Call PRICE_OF_EXPORTS_WOOD_INDUSTRY
Call PRICE_OF_IMPORT_WOOD_INDUSTRY

For k = d_ini To d_fin

    'i es el número 8 de la hoja del mercado
    i = k - 31
    'j es el número 7 de la hoja del mercado
    j = i - 1
    'l es el n anterior de la hoja en summary
    l = k - 1
    'k es el año actual de la hoja summary
    contadorSetPrice = k - 37
    
    hojUsu_Summary.Cells(6, 3).Value = "=Summary!B" & k
    hojUsu_Summary.Cells(7, 3).Value = "=Summary!D" & k
    hojUsu_Summary.Cells(8, 3).Value = "=Summary!F" & k
    hojUsu_Summary.Cells(9, 3).Value = "=Summary!H" & k
    hojUsu_Summary.Cells(10, 3).Value = "=Summary!J" & k
    hojUsu_Summary.Cells(11, 3).Value = "=Summary!L" & k
    hojUsu_Summary.Cells(12, 3).Value = "=Summary!N" & k
    hojUsu_Summary.Cells(13, 3).Value = "=Summary!P" & k
    
    If tipoSolver = 1 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(k, 16) = hojUsu_WoodIndustry.Cells(i, 15).Value
            
            Case 2
            
                hojUsu_Summary.Cells(k, 16) = hojUsu_WoodIndustry.Cells(i, 15).Value
                
            Case 3
            
                hojUsu_Summary.Cells(k, 16) = hojUsu_WoodIndustry.Cells(i, 15).Value
            
        End Select
    
    ElseIf tipoSolver = 2 Then
    
        Select Case setPricesValue
        
            Case 1
            
                hojUsu_Summary.Cells(l, 10) = hojUsu_Forecast.Cells(j, 21).Value
                hojUsu_Summary.Cells(l, 12) = hojUsu_Forecast.Cells(j, 25).Value
                hojUsu_Summary.Cells(l, 14) = hojUsu_Forecast.Cells(j, 29).Value
                hojUsu_Summary.Cells(l, 16) = hojUsu_WoodIndustry.Cells(j, 15).Value
        
                hojUsu_Summary.Cells(k, 10) = hojUsu_Forecast.Cells(i, 21).Value
                hojUsu_Summary.Cells(k, 12) = hojUsu_Forecast.Cells(i, 25).Value
                hojUsu_Summary.Cells(k, 14) = hojUsu_Forecast.Cells(i, 29).Value
                hojUsu_Summary.Cells(k, 16) = hojUsu_WoodIndustry.Cells(i, 15).Value
                
            Case 2
            
                hojUsu_Summary.Cells(l, 10) = hojUsu_Forecast.Cells(j, 19).Value
                hojUsu_Summary.Cells(l, 12) = hojUsu_Forecast.Cells(j, 23).Value
                hojUsu_Summary.Cells(l, 14) = hojUsu_Forecast.Cells(j, 27).Value
                hojUsu_Summary.Cells(l, 16) = hojUsu_WoodIndustry.Cells(j, 15).Value
        
                hojUsu_Summary.Cells(k, 10) = hojUsu_Forecast.Cells(i, 19).Value
                hojUsu_Summary.Cells(k, 12) = hojUsu_Forecast.Cells(i, 23).Value
                hojUsu_Summary.Cells(k, 14) = hojUsu_Forecast.Cells(i, 27).Value
                hojUsu_Summary.Cells(k, 16) = hojUsu_WoodIndustry.Cells(i, 15).Value
            
            Case 3
            
                hojUsu_Summary.Cells(l, 10) = hojUsu_Forecast.Cells(j, 20).Value
                hojUsu_Summary.Cells(l, 12) = hojUsu_Forecast.Cells(j, 24).Value
                hojUsu_Summary.Cells(l, 14) = hojUsu_Forecast.Cells(j, 28).Value
                hojUsu_Summary.Cells(l, 16) = hojUsu_WoodIndustry.Cells(j, 15).Value
        
                hojUsu_Summary.Cells(k, 10) = hojUsu_Forecast.Cells(i, 20).Value
                hojUsu_Summary.Cells(k, 12) = hojUsu_Forecast.Cells(i, 24).Value
                hojUsu_Summary.Cells(k, 14) = hojUsu_Forecast.Cells(i, 28).Value
                hojUsu_Summary.Cells(k, 16) = hojUsu_WoodIndustry.Cells(i, 15).Value
            
        End Select
    
    End If
    
    Select Case tipoSolver
    
        Case 1
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$C$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$P$" & k, Engine:=1
                SolverAdd CellRef:="$B$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$D$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$F$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$H$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$J$" & k, Relation:=3, FormulaText:="$BE$3"
                SolverAdd CellRef:="$L$" & k, Relation:=3, FormulaText:="$BF$3"
                SolverAdd CellRef:="$N$" & k, Relation:=3, FormulaText:="$BG$3"
                SolverAdd CellRef:="$P$" & k, Relation:=3, FormulaText:="$BH$3"
                SolverAdd CellRef:="$J$" & k, Relation:=1, FormulaText:="$BI$3"
                SolverAdd CellRef:="$L$" & k, Relation:=1, FormulaText:="$BJ$3"
                SolverAdd CellRef:="$N$" & k, Relation:=1, FormulaText:="$BK$3"
                SolverAdd CellRef:="$P$" & k, Relation:=1, FormulaText:="$BL$3"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$C$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$P$" & k, Engine:=2
                SolverAdd CellRef:="$B$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$D$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$F$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$H$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$J$" & k, Relation:=3, FormulaText:="$BE$3"
                SolverAdd CellRef:="$L$" & k, Relation:=3, FormulaText:="$BF$3"
                SolverAdd CellRef:="$N$" & k, Relation:=3, FormulaText:="$BG$3"
                SolverAdd CellRef:="$P$" & k, Relation:=3, FormulaText:="$BH$3"
                SolverAdd CellRef:="$J$" & k, Relation:=1, FormulaText:="$BI$3"
                SolverAdd CellRef:="$L$" & k, Relation:=1, FormulaText:="$BJ$3"
                SolverAdd CellRef:="$N$" & k, Relation:=1, FormulaText:="$BK$3"
                SolverAdd CellRef:="$P$" & k, Relation:=1, FormulaText:="$BL$3"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$C$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$P$" & k, Engine:=3
                SolverAdd CellRef:="$B$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$D$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$F$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$H$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$J$" & k, Relation:=3, FormulaText:="$BE$3"
                SolverAdd CellRef:="$L$" & k, Relation:=3, FormulaText:="$BF$3"
                SolverAdd CellRef:="$N$" & k, Relation:=3, FormulaText:="$BG$3"
                SolverAdd CellRef:="$P$" & k, Relation:=3, FormulaText:="$BH$3"
                SolverAdd CellRef:="$J$" & k, Relation:=1, FormulaText:="$BI$3"
                SolverAdd CellRef:="$L$" & k, Relation:=1, FormulaText:="$BJ$3"
                SolverAdd CellRef:="$N$" & k, Relation:=1, FormulaText:="$BK$3"
                SolverAdd CellRef:="$P$" & k, Relation:=1, FormulaText:="$BL$3"
            
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
        Case 2
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$C$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$P$" & k & ",$J$" & k & ",$L$" & k & ",$N$" & k, Engine:=1
                SolverAdd CellRef:="$B$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$D$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$F$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$H$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$J$" & k, Relation:=3, FormulaText:="$BE$3"
                SolverAdd CellRef:="$L$" & k, Relation:=3, FormulaText:="$BF$3"
                SolverAdd CellRef:="$N$" & k, Relation:=3, FormulaText:="$BG$3"
                SolverAdd CellRef:="$P$" & k, Relation:=3, FormulaText:="$BH$3"
                SolverAdd CellRef:="$J$" & k, Relation:=1, FormulaText:="$BI$3"
                SolverAdd CellRef:="$L$" & k, Relation:=1, FormulaText:="$BJ$3"
                SolverAdd CellRef:="$N$" & k, Relation:=1, FormulaText:="$BK$3"
                SolverAdd CellRef:="$P$" & k, Relation:=1, FormulaText:="$BL$3"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$C$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$P$" & k & ",$J$" & k & ",$L$" & k & ",$N$" & k, Engine:=2
                SolverAdd CellRef:="$B$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$D$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$F$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$H$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$J$" & k, Relation:=3, FormulaText:="$BE$3"
                SolverAdd CellRef:="$L$" & k, Relation:=3, FormulaText:="$BF$3"
                SolverAdd CellRef:="$N$" & k, Relation:=3, FormulaText:="$BG$3"
                SolverAdd CellRef:="$P$" & k, Relation:=3, FormulaText:="$BH$3"
                SolverAdd CellRef:="$J$" & k, Relation:=1, FormulaText:="$BI$3"
                SolverAdd CellRef:="$L$" & k, Relation:=1, FormulaText:="$BJ$3"
                SolverAdd CellRef:="$N$" & k, Relation:=1, FormulaText:="$BK$3"
                SolverAdd CellRef:="$P$" & k, Relation:=1, FormulaText:="$BL$3"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$C$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$P$" & k & ",$J$" & k & ",$L$" & k & ",$N$" & k, Engine:=3
                SolverAdd CellRef:="$B$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$D$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$F$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$H$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$J$" & k, Relation:=3, FormulaText:="$BE$3"
                SolverAdd CellRef:="$L$" & k, Relation:=3, FormulaText:="$BF$3"
                SolverAdd CellRef:="$N$" & k, Relation:=3, FormulaText:="$BG$3"
                SolverAdd CellRef:="$P$" & k, Relation:=3, FormulaText:="$BH$3"
                SolverAdd CellRef:="$J$" & k, Relation:=1, FormulaText:="$BI$3"
                SolverAdd CellRef:="$L$" & k, Relation:=1, FormulaText:="$BJ$3"
                SolverAdd CellRef:="$N$" & k, Relation:=1, FormulaText:="$BK$3"
                SolverAdd CellRef:="$P$" & k, Relation:=1, FormulaText:="$BL$3"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
    End Select
    
        hojUsu_Forecast.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 2).Value
        hojUsu_Forecast.Cells(c_ini, 10) = hojUsu_Summary.Cells(k, 4).Value
        hojUsu_Forecast.Cells(c_ini, 14) = hojUsu_Summary.Cells(k, 6).Value
        hojUsu_Forecast.Cells(c_ini, 18) = hojUsu_Summary.Cells(k, 8).Value
        hojUsu_Forecast.Cells(c_ini, 22) = hojUsu_Summary.Cells(k, 10).Value
        hojUsu_Forecast.Cells(c_ini, 26) = hojUsu_Summary.Cells(k, 12).Value
        hojUsu_Forecast.Cells(c_ini, 30) = hojUsu_Summary.Cells(k, 14).Value
        
        hojUsu_SetPricesWoodIndustry.Cells(c_ini, 3) = hojUsu_Summary.Cells(14, 3)
        hojUsu_SetPricesWoodIndustry.Cells(c_ini, 4) = hojUsu_Summary.Cells(15, 3)
        hojUsu_SetPricesWoodIndustry.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 10)
        hojUsu_SetPricesWoodIndustry.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 12)
        hojUsu_SetPricesWoodIndustry.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 14)
        hojUsu_SetPricesWoodIndustry.Cells(c_ini, 8) = hojUsu_Summary.Cells(k, 16)
        c_ini = c_ini + 1
    
    Next k
    
Application.ScreenUpdating = True

End Sub

Sub SOLVER_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRangeSolver")
b = hojUsu_SystemOptions.Range("FinalYearRangeSolver")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de engine y verificación de interacción de las Eq

metodoSolver = hojUsu_SystemOptions.Range("IterationMethod")
'hojUsu_SystemOptions.Range("SelectProcess") = 2
tipoSolver = hojUsu_SystemOptions.Range("VariablesSolver")
setPricesValue = hojUsu_SystemOptions.Range("OriginForVariablesTwo")

Call SUPPLY_FURNITURE_INDUSTRY
Call CONSUMPTION_FURNITURE_INDUSTRY
Call EXPORTS_FURNITURE_INDUSTRY
Call IMPORTS_FURNITURE_INDUSTRY
Call PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY
Call PRICE_OF_EXPORTS_FURNITURE_INDUSTRY
Call PRICE_OF_IMPORT_FURNITURE_INDUSTRY

For k = d_ini To d_fin

    'i es el número 8 de la hoja del mercado
    i = k - 31
    'j es el número 7 de la hoja del mercado
    j = i - 1
    'l es el n anterior de la hoja en summary
    l = k - 1
    'k es el año actual de la hoja summary
    contadorSetPrice = k - 37
    
    hojUsu_Summary.Cells(6, 5).Value = "=Summary!T" & k
    hojUsu_Summary.Cells(7, 5).Value = "=Summary!V" & k
    hojUsu_Summary.Cells(8, 5).Value = "=Summary!X" & k
    hojUsu_Summary.Cells(9, 5).Value = "=Summary!Z" & k
    hojUsu_Summary.Cells(10, 5).Value = "=Summary!AB" & k
    hojUsu_Summary.Cells(11, 5).Value = "=Summary!AD" & k
    hojUsu_Summary.Cells(12, 5).Value = "=Summary!AF" & k
    hojUsu_Summary.Cells(13, 5).Value = "=Summary!AH" & k
    
    If tipoSolver = 1 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(k, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
            
            Case 2
            
                hojUsu_Summary.Cells(k, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
                
            Case 3
            
                hojUsu_Summary.Cells(k, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
            
        End Select
    
    ElseIf tipoSolver = 2 Then
    
        Select Case setPricesValue
        
            Case 1
            
                hojUsu_Summary.Cells(l, 28) = hojUsu_Forecast.Cells(j, 49).Value
                hojUsu_Summary.Cells(l, 30) = hojUsu_Forecast.Cells(j, 53).Value
                hojUsu_Summary.Cells(l, 32) = hojUsu_Forecast.Cells(j, 57).Value
                hojUsu_Summary.Cells(l, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 28) = hojUsu_Forecast.Cells(i, 49).Value
                hojUsu_Summary.Cells(k, 30) = hojUsu_Forecast.Cells(i, 53).Value
                hojUsu_Summary.Cells(k, 32) = hojUsu_Forecast.Cells(i, 57).Value
                hojUsu_Summary.Cells(k, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
                
            Case 2
            
                hojUsu_Summary.Cells(l, 28) = hojUsu_Forecast.Cells(j, 47).Value
                hojUsu_Summary.Cells(l, 30) = hojUsu_Forecast.Cells(j, 51).Value
                hojUsu_Summary.Cells(l, 32) = hojUsu_Forecast.Cells(j, 55).Value
                hojUsu_Summary.Cells(l, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 28) = hojUsu_Forecast.Cells(i, 47).Value
                hojUsu_Summary.Cells(k, 30) = hojUsu_Forecast.Cells(i, 51).Value
                hojUsu_Summary.Cells(k, 32) = hojUsu_Forecast.Cells(i, 55).Value
                hojUsu_Summary.Cells(k, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
            
            Case 3
            
                hojUsu_Summary.Cells(l, 28) = hojUsu_Forecast.Cells(j, 48).Value
                hojUsu_Summary.Cells(l, 30) = hojUsu_Forecast.Cells(j, 52).Value
                hojUsu_Summary.Cells(l, 32) = hojUsu_Forecast.Cells(j, 56).Value
                hojUsu_Summary.Cells(l, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 28) = hojUsu_Forecast.Cells(i, 48).Value
                hojUsu_Summary.Cells(k, 30) = hojUsu_Forecast.Cells(i, 52).Value
                hojUsu_Summary.Cells(k, 32) = hojUsu_Forecast.Cells(i, 56).Value
                hojUsu_Summary.Cells(k, 34) = hojUsu_FurnitureIndustry.Cells(i, 15).Value
            
        End Select
    
    End If
    
    Select Case tipoSolver
    
        Case 1
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$E$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AH$" & k, Engine:=1
                SolverAdd CellRef:="$T$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$V$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$X$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$Z$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AB$" & k, Relation:=3, FormulaText:="$BE$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=3, FormulaText:="$BF$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=3, FormulaText:="$BG$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=3, FormulaText:="$BH$5"
                SolverAdd CellRef:="$AB$" & k, Relation:=1, FormulaText:="$BI$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=1, FormulaText:="$BJ$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=1, FormulaText:="$BK$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=1, FormulaText:="$BL$5"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$E$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AH$" & k, Engine:=2
                SolverAdd CellRef:="$T$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$V$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$X$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$Z$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AB$" & k, Relation:=3, FormulaText:="$BE$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=3, FormulaText:="$BF$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=3, FormulaText:="$BG$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=3, FormulaText:="$BH$5"
                SolverAdd CellRef:="$AB$" & k, Relation:=1, FormulaText:="$BI$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=1, FormulaText:="$BJ$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=1, FormulaText:="$BK$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=1, FormulaText:="$BL$5"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$E$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AH$" & k, Engine:=3
                SolverAdd CellRef:="$T$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$V$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$X$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$Z$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AB$" & k, Relation:=3, FormulaText:="$BE$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=3, FormulaText:="$BF$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=3, FormulaText:="$BG$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=3, FormulaText:="$BH$5"
                SolverAdd CellRef:="$AB$" & k, Relation:=1, FormulaText:="$BI$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=1, FormulaText:="$BJ$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=1, FormulaText:="$BK$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=1, FormulaText:="$BL$5"
            
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
        Case 2
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$E$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AB$" & k & ",$AD$" & k & ",$AF$" & k & ",$AH$" & k, Engine:=1
                SolverAdd CellRef:="$T$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$V$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$X$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$Z$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AB$" & k, Relation:=3, FormulaText:="$BE$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=3, FormulaText:="$BF$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=3, FormulaText:="$BG$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=3, FormulaText:="$BH$5"
                SolverAdd CellRef:="$AB$" & k, Relation:=1, FormulaText:="$BI$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=1, FormulaText:="$BJ$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=1, FormulaText:="$BK$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=1, FormulaText:="$BL$5"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$E$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AB$" & k & ",$AD$" & k & ",$AF$" & k & ",$AH$" & k, Engine:=2
                SolverAdd CellRef:="$T$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$V$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$X$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$Z$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AB$" & k, Relation:=3, FormulaText:="$BE$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=3, FormulaText:="$BF$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=3, FormulaText:="$BG$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=3, FormulaText:="$BH$5"
                SolverAdd CellRef:="$AB$" & k, Relation:=1, FormulaText:="$BI$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=1, FormulaText:="$BJ$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=1, FormulaText:="$BK$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=1, FormulaText:="$BL$5"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$E$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AB$" & k & ",$AD$" & k & ",$AF$" & k & ",$AH$" & k, Engine:=3
                SolverAdd CellRef:="$T$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$V$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$X$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$Z$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AB$" & k, Relation:=3, FormulaText:="$BE$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=3, FormulaText:="$BF$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=3, FormulaText:="$BG$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=3, FormulaText:="$BH$5"
                SolverAdd CellRef:="$AB$" & k, Relation:=1, FormulaText:="$BI$5"
                SolverAdd CellRef:="$AD$" & k, Relation:=1, FormulaText:="$BJ$5"
                SolverAdd CellRef:="$AF$" & k, Relation:=1, FormulaText:="$BK$5"
                SolverAdd CellRef:="$AH$" & k, Relation:=1, FormulaText:="$BL$5"
       
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
    End Select
    
        hojUsu_Forecast.Cells(c_ini, 34) = hojUsu_Summary.Cells(k, 20).Value
        hojUsu_Forecast.Cells(c_ini, 38) = hojUsu_Summary.Cells(k, 22).Value
        hojUsu_Forecast.Cells(c_ini, 42) = hojUsu_Summary.Cells(k, 24).Value
        hojUsu_Forecast.Cells(c_ini, 46) = hojUsu_Summary.Cells(k, 26).Value
        hojUsu_Forecast.Cells(c_ini, 50) = hojUsu_Summary.Cells(k, 28).Value
        hojUsu_Forecast.Cells(c_ini, 54) = hojUsu_Summary.Cells(k, 30).Value
        hojUsu_Forecast.Cells(c_ini, 58) = hojUsu_Summary.Cells(k, 32).Value
        
        hojUsu_SetPricesFurniture.Cells(c_ini, 3) = hojUsu_Summary.Cells(14, 5)
        hojUsu_SetPricesFurniture.Cells(c_ini, 4) = hojUsu_Summary.Cells(15, 5)
        hojUsu_SetPricesFurniture.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 28)
        hojUsu_SetPricesFurniture.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 30)
        hojUsu_SetPricesFurniture.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 32)
        hojUsu_SetPricesFurniture.Cells(c_ini, 8) = hojUsu_Summary.Cells(k, 34)
        c_ini = c_ini + 1
    
    Next k
    
Application.ScreenUpdating = True

End Sub
Sub SOLVER_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRangeSolver")
b = hojUsu_SystemOptions.Range("FinalYearRangeSolver")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de engine y verificación de interacción de las Eq

metodoSolver = hojUsu_SystemOptions.Range("IterationMethod")
'hojUsu_SystemOptions.Range("SelectProcess") = 2
tipoSolver = hojUsu_SystemOptions.Range("VariablesSolver")
setPricesValue = hojUsu_SystemOptions.Range("OriginForVariablesTwo")

Call SUPPLY_PULP_PAPER_INDUSTRY
Call CONSUMPTION_PULP_PAPER_INDUSTRY
Call EXPORTS_PULP_PAPER_INDUSTRY
Call IMPORTS_PULP_PAPER_INDUSTRY
Call PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY
Call PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY
Call PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY

For k = d_ini To d_fin

    'i es el número 8 de la hoja del mercado
    i = k - 31
    'j es el número 7 de la hoja del mercado
    j = i - 1
    'l es el n anterior de la hoja en summary
    l = k - 1
    'k es el año actual de la hoja summary
    contadorSetPrice = k - 37
    
    hojUsu_Summary.Cells(6, 7).Value = "=Summary!AL" & k
    hojUsu_Summary.Cells(7, 7).Value = "=Summary!AN" & k
    hojUsu_Summary.Cells(8, 7).Value = "=Summary!AP" & k
    hojUsu_Summary.Cells(9, 7).Value = "=Summary!AR" & k
    hojUsu_Summary.Cells(10, 7).Value = "=Summary!AT" & k
    hojUsu_Summary.Cells(11, 7).Value = "=Summary!AV" & k
    hojUsu_Summary.Cells(12, 7).Value = "=Summary!AX" & k
    hojUsu_Summary.Cells(13, 7).Value = "=Summary!AZ" & k
    
    If tipoSolver = 1 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(k, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
            
            Case 2
            
                hojUsu_Summary.Cells(k, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
                
            Case 3
            
                hojUsu_Summary.Cells(k, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
            
        End Select
    
    ElseIf tipoSolver = 2 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(l, 46) = hojUsu_Forecast.Cells(j, 77).Value
                hojUsu_Summary.Cells(l, 48) = hojUsu_Forecast.Cells(j, 81).Value
                hojUsu_Summary.Cells(l, 50) = hojUsu_Forecast.Cells(j, 85).Value
                hojUsu_Summary.Cells(l, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 46) = hojUsu_Forecast.Cells(i, 77).Value
                hojUsu_Summary.Cells(k, 48) = hojUsu_Forecast.Cells(i, 81).Value
                hojUsu_Summary.Cells(k, 50) = hojUsu_Forecast.Cells(i, 85).Value
                hojUsu_Summary.Cells(k, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
                
            Case 2
            
                hojUsu_Summary.Cells(l, 46) = hojUsu_Forecast.Cells(j, 75).Value
                hojUsu_Summary.Cells(l, 48) = hojUsu_Forecast.Cells(j, 79).Value
                hojUsu_Summary.Cells(l, 50) = hojUsu_Forecast.Cells(j, 83).Value
                hojUsu_Summary.Cells(l, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 46) = hojUsu_Forecast.Cells(i, 75).Value
                hojUsu_Summary.Cells(k, 48) = hojUsu_Forecast.Cells(i, 79).Value
                hojUsu_Summary.Cells(k, 50) = hojUsu_Forecast.Cells(i, 83).Value
                hojUsu_Summary.Cells(k, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
            
            Case 3
            
                hojUsu_Summary.Cells(l, 46) = hojUsu_Forecast.Cells(j, 76).Value
                hojUsu_Summary.Cells(l, 48) = hojUsu_Forecast.Cells(j, 80).Value
                hojUsu_Summary.Cells(l, 50) = hojUsu_Forecast.Cells(j, 84).Value
                hojUsu_Summary.Cells(l, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 46) = hojUsu_Forecast.Cells(i, 76).Value
                hojUsu_Summary.Cells(k, 48) = hojUsu_Forecast.Cells(i, 80).Value
                hojUsu_Summary.Cells(k, 50) = hojUsu_Forecast.Cells(i, 84).Value
                hojUsu_Summary.Cells(k, 52) = hojUsu_PulpPaperIndustry.Cells(i, 15).Value
            
        End Select
    
    End If
    
    Select Case tipoSolver
    
        Case 1
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$G$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AZ$" & k, Engine:=1
                SolverAdd CellRef:="$AL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AN$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AP$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AR$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AT$" & k, Relation:=3, FormulaText:="$BE$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=3, FormulaText:="$BF$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=3, FormulaText:="$BG$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=3, FormulaText:="$BH$7"
                SolverAdd CellRef:="$AT$" & k, Relation:=1, FormulaText:="$BI$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=1, FormulaText:="$BJ$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=1, FormulaText:="$BK$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=1, FormulaText:="$BL$7"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$G$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AZ$" & k, Engine:=2
                SolverAdd CellRef:="$AL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AN$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AP$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AR$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AT$" & k, Relation:=3, FormulaText:="$BE$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=3, FormulaText:="$BF$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=3, FormulaText:="$BG$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=3, FormulaText:="$BH$7"
                SolverAdd CellRef:="$AT$" & k, Relation:=1, FormulaText:="$BI$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=1, FormulaText:="$BJ$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=1, FormulaText:="$BK$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=1, FormulaText:="$BL$7"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$G$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AZ$" & k, Engine:=3
                SolverAdd CellRef:="$AL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AN$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AP$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AR$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AT$" & k, Relation:=3, FormulaText:="$BE$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=3, FormulaText:="$BF$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=3, FormulaText:="$BG$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=3, FormulaText:="$BH$7"
                SolverAdd CellRef:="$AT$" & k, Relation:=1, FormulaText:="$BI$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=1, FormulaText:="$BJ$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=1, FormulaText:="$BK$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=1, FormulaText:="$BL$7"
            
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
        Case 2
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$G$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AT$" & k & ",$AV$" & k & ",$AX$" & k & ",$AZ$" & k, Engine:=1
                SolverAdd CellRef:="$AL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AN$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AP$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AR$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AT$" & k, Relation:=3, FormulaText:="$BE$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=3, FormulaText:="$BF$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=3, FormulaText:="$BG$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=3, FormulaText:="$BH$7"
                SolverAdd CellRef:="$AT$" & k, Relation:=1, FormulaText:="$BI$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=1, FormulaText:="$BJ$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=1, FormulaText:="$BK$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=1, FormulaText:="$BL$7"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$G$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AT$" & k & ",$AV$" & k & ",$AX$" & k & ",$AZ$" & k, Engine:=2
                SolverAdd CellRef:="$AL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AN$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AP$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AR$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AT$" & k, Relation:=3, FormulaText:="$BE$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=3, FormulaText:="$BF$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=3, FormulaText:="$BG$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=3, FormulaText:="$BH$7"
                SolverAdd CellRef:="$AT$" & k, Relation:=1, FormulaText:="$BI$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=1, FormulaText:="$BJ$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=1, FormulaText:="$BK$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=1, FormulaText:="$BL$7"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$G$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$AT$" & k & ",$AV$" & k & ",$AX$" & k & ",$AZ$" & k, Engine:=3
                SolverAdd CellRef:="$AL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AN$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AP$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AR$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$AT$" & k, Relation:=3, FormulaText:="$BE$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=3, FormulaText:="$BF$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=3, FormulaText:="$BG$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=3, FormulaText:="$BH$7"
                SolverAdd CellRef:="$AT$" & k, Relation:=1, FormulaText:="$BI$7"
                SolverAdd CellRef:="$AV$" & k, Relation:=1, FormulaText:="$BJ$7"
                SolverAdd CellRef:="$AX$" & k, Relation:=1, FormulaText:="$BK$7"
                SolverAdd CellRef:="$AZ$" & k, Relation:=1, FormulaText:="$BL$7"
       
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
    End Select
    
        hojUsu_Forecast.Cells(c_ini, 62) = hojUsu_Summary.Cells(k, 38).Value
        hojUsu_Forecast.Cells(c_ini, 66) = hojUsu_Summary.Cells(k, 40).Value
        hojUsu_Forecast.Cells(c_ini, 70) = hojUsu_Summary.Cells(k, 42).Value
        hojUsu_Forecast.Cells(c_ini, 74) = hojUsu_Summary.Cells(k, 44).Value
        hojUsu_Forecast.Cells(c_ini, 78) = hojUsu_Summary.Cells(k, 46).Value
        hojUsu_Forecast.Cells(c_ini, 82) = hojUsu_Summary.Cells(k, 48).Value
        hojUsu_Forecast.Cells(c_ini, 86) = hojUsu_Summary.Cells(k, 50).Value
        
        hojUsu_SetPricesPulpPaper.Cells(c_ini, 3) = hojUsu_Summary.Cells(14, 7)
        hojUsu_SetPricesPulpPaper.Cells(c_ini, 4) = hojUsu_Summary.Cells(15, 7)
        hojUsu_SetPricesPulpPaper.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 46)
        hojUsu_SetPricesPulpPaper.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 48)
        hojUsu_SetPricesPulpPaper.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 50)
        hojUsu_SetPricesPulpPaper.Cells(c_ini, 8) = hojUsu_Summary.Cells(k, 52)
        c_ini = c_ini + 1
    
    Next k
    
Application.ScreenUpdating = True

End Sub
Sub SOLVER_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRangeSolver")
b = hojUsu_SystemOptions.Range("FinalYearRangeSolver")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de engine y verificación de interacción de las Eq

metodoSolver = hojUsu_SystemOptions.Range("IterationMethod")
'hojUsu_SystemOptions.Range("SelectProcess") = 2
tipoSolver = hojUsu_SystemOptions.Range("VariablesSolver")
setPricesValue = hojUsu_SystemOptions.Range("OriginForVariablesTwo")

Call SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
Call SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST
Call CONSUMPTION_WOOD_INDUSTRIAL
Call EXPORTS_WOOD_INDUSTRIAL
Call IMPORTS_WOOD_INDUSTRIAL
Call PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL
Call PRICE_OF_EXPORTS_WOOD_INDUSTRIAL
Call PRICE_OF_IMPORT_WOOD_INDUSTRIAL

For k = d_ini To d_fin

    'i es el número 8 de la hoja del mercado
    i = k - 31
    'j es el número 7 de la hoja del mercado
    j = i - 1
    'l es el n anterior de la hoja en summary
    l = k - 1
    'k es el año actual de la hoja summary
    contadorSetPrice = k - 37
    
    hojUsu_Summary.Cells(6, 21).Value = "=Summary!BD" & k & "+Summary!BF" & k
    hojUsu_Summary.Cells(7, 21).Value = "=Summary!BH" & k
    hojUsu_Summary.Cells(8, 21).Value = "=Summary!BJ" & k
    hojUsu_Summary.Cells(9, 21).Value = "=Summary!BL" & k
    hojUsu_Summary.Cells(10, 21).Value = "=Summary!BN" & k
    hojUsu_Summary.Cells(11, 21).Value = "=Summary!BP" & k
    hojUsu_Summary.Cells(12, 21).Value = "=Summary!BR" & k
    hojUsu_Summary.Cells(13, 21).Value = "=Summary!BT" & k
    
    If tipoSolver = 1 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(k, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
            
            Case 2
            
                hojUsu_Summary.Cells(k, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
                
            Case 3
            
                hojUsu_Summary.Cells(k, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
            
        End Select
    
    ElseIf tipoSolver = 2 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(l, 66) = hojUsu_Forecast.Cells(j, 109).Value
                hojUsu_Summary.Cells(l, 68) = hojUsu_Forecast.Cells(j, 113).Value
                hojUsu_Summary.Cells(l, 70) = hojUsu_Forecast.Cells(j, 117).Value
                hojUsu_Summary.Cells(l, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
        
                hojUsu_Summary.Cells(k, 66) = hojUsu_Forecast.Cells(i, 109).Value
                hojUsu_Summary.Cells(k, 68) = hojUsu_Forecast.Cells(i, 113).Value
                hojUsu_Summary.Cells(k, 70) = hojUsu_Forecast.Cells(i, 117).Value
                hojUsu_Summary.Cells(k, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
                
            Case 2
            
                hojUsu_Summary.Cells(l, 66) = hojUsu_Forecast.Cells(j, 107).Value
                hojUsu_Summary.Cells(l, 68) = hojUsu_Forecast.Cells(j, 111).Value
                hojUsu_Summary.Cells(l, 70) = hojUsu_Forecast.Cells(j, 115).Value
                hojUsu_Summary.Cells(l, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
        
                hojUsu_Summary.Cells(k, 66) = hojUsu_Forecast.Cells(i, 107).Value
                hojUsu_Summary.Cells(k, 68) = hojUsu_Forecast.Cells(i, 111).Value
                hojUsu_Summary.Cells(k, 70) = hojUsu_Forecast.Cells(i, 115).Value
                hojUsu_Summary.Cells(k, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
            
            Case 3
            
                hojUsu_Summary.Cells(l, 66) = hojUsu_Forecast.Cells(j, 108).Value
                hojUsu_Summary.Cells(l, 68) = hojUsu_Forecast.Cells(j, 112).Value
                hojUsu_Summary.Cells(l, 70) = hojUsu_Forecast.Cells(j, 116).Value
                hojUsu_Summary.Cells(l, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
        
                hojUsu_Summary.Cells(k, 66) = hojUsu_Forecast.Cells(i, 108).Value
                hojUsu_Summary.Cells(k, 68) = hojUsu_Forecast.Cells(i, 112).Value
                hojUsu_Summary.Cells(k, 70) = hojUsu_Forecast.Cells(i, 116).Value
                hojUsu_Summary.Cells(k, 72) = hojUsu_WoodIndustrial.Cells(i, 20).Value
            
        End Select
    
    End If
    
    Select Case tipoSolver
    
        Case 1
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$U$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$BT$" & k, Engine:=1
                SolverAdd CellRef:="$BD$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BF$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BH$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BJ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BN$" & k, Relation:=3, FormulaText:="$BE$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=3, FormulaText:="$BF$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=3, FormulaText:="$BG$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=3, FormulaText:="$BH$9"
                SolverAdd CellRef:="$BN$" & k, Relation:=1, FormulaText:="$BI$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=1, FormulaText:="$BJ$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=1, FormulaText:="$BK$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=1, FormulaText:="$BL$9"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$U$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$BT$" & k, Engine:=2
                SolverAdd CellRef:="$BD$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BF$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BH$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BJ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BN$" & k, Relation:=3, FormulaText:="$BE$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=3, FormulaText:="$BF$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=3, FormulaText:="$BG$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=3, FormulaText:="$BH$9"
                SolverAdd CellRef:="$BN$" & k, Relation:=1, FormulaText:="$BI$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=1, FormulaText:="$BJ$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=1, FormulaText:="$BK$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=1, FormulaText:="$BL$9"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$U$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$BT$" & k, Engine:=3
                SolverAdd CellRef:="$BD$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BF$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BH$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BJ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BN$" & k, Relation:=3, FormulaText:="$BE$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=3, FormulaText:="$BF$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=3, FormulaText:="$BG$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=3, FormulaText:="$BH$9"
                SolverAdd CellRef:="$BN$" & k, Relation:=1, FormulaText:="$BI$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=1, FormulaText:="$BJ$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=1, FormulaText:="$BK$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=1, FormulaText:="$BL$9"
            
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
        Case 2
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$U$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$BN$" & k & ",$BP$" & k & ",$BR$" & k & ",$BT$" & k, Engine:=1
                SolverAdd CellRef:="$BD$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BF$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BH$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BJ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BN$" & k, Relation:=3, FormulaText:="$BE$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=3, FormulaText:="$BF$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=3, FormulaText:="$BG$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=3, FormulaText:="$BH$9"
                SolverAdd CellRef:="$BN$" & k, Relation:=1, FormulaText:="$BI$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=1, FormulaText:="$BJ$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=1, FormulaText:="$BK$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=1, FormulaText:="$BL$9"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$U$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$BN$" & k & ",$BP$" & k & ",$BR$" & k & ",$BT$" & k, Engine:=2
                SolverAdd CellRef:="$BD$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BF$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BH$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BJ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BN$" & k, Relation:=3, FormulaText:="$BE$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=3, FormulaText:="$BF$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=3, FormulaText:="$BG$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=3, FormulaText:="$BH$9"
                SolverAdd CellRef:="$BN$" & k, Relation:=1, FormulaText:="$BI$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=1, FormulaText:="$BJ$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=1, FormulaText:="$BK$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=1, FormulaText:="$BL$9"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$U$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$BN$" & k & ",$BP$" & k & ",$BR$" & k & ",$BT$" & k, Engine:=3
                SolverAdd CellRef:="$BD$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BF$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BH$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BJ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BL$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BN$" & k, Relation:=3, FormulaText:="$BE$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=3, FormulaText:="$BF$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=3, FormulaText:="$BG$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=3, FormulaText:="$BH$9"
                SolverAdd CellRef:="$BN$" & k, Relation:=1, FormulaText:="$BI$9"
                SolverAdd CellRef:="$BP$" & k, Relation:=1, FormulaText:="$BJ$9"
                SolverAdd CellRef:="$BR$" & k, Relation:=1, FormulaText:="$BK$9"
                SolverAdd CellRef:="$BT$" & k, Relation:=1, FormulaText:="$BL$9"
            
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
    End Select
    
        hojUsu_Forecast.Cells(c_ini, 90) = hojUsu_Summary.Cells(k, 56).Value
        hojUsu_Forecast.Cells(c_ini, 94) = hojUsu_Summary.Cells(k, 58).Value
        hojUsu_Forecast.Cells(c_ini, 98) = hojUsu_Summary.Cells(k, 60).Value
        hojUsu_Forecast.Cells(c_ini, 102) = hojUsu_Summary.Cells(k, 62).Value
        hojUsu_Forecast.Cells(c_ini, 106) = hojUsu_Summary.Cells(k, 64).Value
        hojUsu_Forecast.Cells(c_ini, 110) = hojUsu_Summary.Cells(k, 66).Value
        hojUsu_Forecast.Cells(c_ini, 114) = hojUsu_Summary.Cells(k, 68).Value
        hojUsu_Forecast.Cells(c_ini, 118) = hojUsu_Summary.Cells(k, 70).Value
        
        hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 3) = hojUsu_Summary.Cells(14, 21)
        hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 4) = hojUsu_Summary.Cells(15, 21)
        hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 66)
        hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 68)
        hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 70)
        hojUsu_SetPricesWoodIndustrial.Cells(c_ini, 8) = hojUsu_Summary.Cells(k, 72)
        c_ini = c_ini + 1
    
    Next k
    
Application.ScreenUpdating = True

End Sub
Sub SOLVER_FIREWOOD()

Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_fin, d_fin, f As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast solo los valores)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRangeSolver")
b = hojUsu_SystemOptions.Range("FinalYearRangeSolver")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_fin = b - 1967
d_fin = b - 1936

'Tipo de engine y verificación de interacción de las Eq

metodoSolver = hojUsu_SystemOptions.Range("IterationMethod")
'hojUsu_SystemOptions.Range("SelectProcess") = 2
tipoSolver = hojUsu_SystemOptions.Range("VariablesSolver")
setPricesValue = hojUsu_SystemOptions.Range("OriginForVariablesTwo")

Call SUPPLY_FIREWOOD
Call CONSUMPTION_FIREWOOD
Call EXPORTS_FIREWOOD
Call IMPORTS_FIREWOOD
Call PRICE_OF_CONSUMPTION_FIREWOOD
Call PRICE_OF_EXPORTS_FIREWOOD
Call PRICE_OF_IMPORT_FIREWOOD

For k = d_ini To d_fin

    'i es el número 8 de la hoja del mercado
    i = k - 31
    'j es el número 7 de la hoja del mercado
    j = i - 1
    'l es el n anterior de la hoja en summary
    l = k - 1
    'k es el año actual de la hoja summary
    contadorSetPrice = k - 37
    
    hojUsu_Summary.Cells(6, 23).Value = "=Summary!BX" & k
    hojUsu_Summary.Cells(7, 23).Value = "=Summary!BZ" & k
    hojUsu_Summary.Cells(10, 23).Value = "=Summary!CF" & k
    
    If tipoSolver = 1 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(k, 90) = hojUsu_Firewood.Cells(i, 15).Value
            
            Case 2
            
                hojUsu_Summary.Cells(k, 90) = hojUsu_Firewood.Cells(i, 15).Value
                
            Case 3
            
                hojUsu_Summary.Cells(k, 90) = hojUsu_Firewood.Cells(i, 15).Value
            
        End Select
    
    ElseIf tipoSolver = 2 Then
    
        Select Case setPricesValue
        
            Case 1
        
                hojUsu_Summary.Cells(l, 84) = hojUsu_Forecast.Cells(j, 129).Value
                hojUsu_Summary.Cells(l, 90) = hojUsu_Firewood.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 84) = hojUsu_Forecast.Cells(i, 129).Value
                hojUsu_Summary.Cells(k, 90) = hojUsu_Firewood.Cells(i, 15).Value
                
            Case 2
            
                hojUsu_Summary.Cells(l, 84) = hojUsu_Forecast.Cells(j, 128).Value
                hojUsu_Summary.Cells(l, 90) = hojUsu_Firewood.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 84) = hojUsu_Forecast.Cells(i, 128).Value
                hojUsu_Summary.Cells(k, 90) = hojUsu_Firewood.Cells(i, 15).Value
            
            Case 3
            
                hojUsu_Summary.Cells(l, 84) = hojUsu_Forecast.Cells(j, 127).Value
                hojUsu_Summary.Cells(l, 90) = hojUsu_Firewood.Cells(i, 15).Value
        
                hojUsu_Summary.Cells(k, 84) = hojUsu_Forecast.Cells(i, 127).Value
                hojUsu_Summary.Cells(k, 90) = hojUsu_Firewood.Cells(i, 15).Value
            
        End Select
    
    End If
    
    Select Case tipoSolver
    
        Case 1
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$W$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$CL$" & k, Engine:=1
                SolverAdd CellRef:="$BX$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BZ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$CF$" & k, Relation:=3, FormulaText:="$BE$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=3, FormulaText:="$BH$11"
                SolverAdd CellRef:="$CF$" & k, Relation:=1, FormulaText:="$BI$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=1, FormulaText:="$BL$11"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$W$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$CL$" & k, Engine:=2
                SolverAdd CellRef:="$BX$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BZ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$CF$" & k, Relation:=3, FormulaText:="$BE$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=3, FormulaText:="$BH$11"
                SolverAdd CellRef:="$CF$" & k, Relation:=1, FormulaText:="$BI$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=1, FormulaText:="$BL$11"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$W$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$CL$" & k, Engine:=3
                SolverAdd CellRef:="$BX$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BZ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$CF$" & k, Relation:=3, FormulaText:="$BE$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=3, FormulaText:="$BH$11"
                SolverAdd CellRef:="$CF$" & k, Relation:=1, FormulaText:="$BI$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=1, FormulaText:="$BL$11"
            
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
        Case 2
        
            Select Case metodoSolver
            
                Case 1
            
                SolverReset
                
                SolverOk SetCell:="$W$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$CF$" & k & ",$CL$" & k, Engine:=1
                SolverAdd CellRef:="$BX$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BZ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$CF$" & k, Relation:=3, FormulaText:="$BE$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=3, FormulaText:="$BH$11"
                SolverAdd CellRef:="$CF$" & k, Relation:=1, FormulaText:="$BI$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=1, FormulaText:="$BL$11"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 2
            
                SolverReset
                
                SolverOk SetCell:="$W$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$CF$" & k & ",$CL$" & k, Engine:=2
                SolverAdd CellRef:="$BX$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BZ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$CF$" & k, Relation:=3, FormulaText:="$BE$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=3, FormulaText:="$BH$11"
                SolverAdd CellRef:="$CF$" & k, Relation:=1, FormulaText:="$BI$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=1, FormulaText:="$BL$11"
                
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
                
                Case 3
            
                SolverReset
                
                SolverOk SetCell:="$W$16", MaxMinVal:=3, ValueOf:=0, _
                    ByChange:="$CF$" & k & ",$CL$" & k, Engine:=3
                SolverAdd CellRef:="$BX$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$BZ$" & k, Relation:=3, FormulaText:="0"
                SolverAdd CellRef:="$CF$" & k, Relation:=3, FormulaText:="$BE$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=3, FormulaText:="$BH$11"
                SolverAdd CellRef:="$CF$" & k, Relation:=1, FormulaText:="$BI$11"
                SolverAdd CellRef:="$CL$" & k, Relation:=1, FormulaText:="$BL$11"
            
                SolverSolve True
                
                Application.Calculation = xlCalculationAutomatic
        
            End Select
        
    End Select
    
        hojUsu_Forecast.Cells(c_ini, 122) = hojUsu_Summary.Cells(k, 76).Value
        hojUsu_Forecast.Cells(c_ini, 126) = hojUsu_Summary.Cells(k, 78).Value
        hojUsu_Forecast.Cells(c_ini, 130) = hojUsu_Summary.Cells(k, 84).Value
        
        hojUsu_SetPricesFirewood.Cells(c_ini, 3) = hojUsu_Summary.Cells(14, 23)
        hojUsu_SetPricesFirewood.Cells(c_ini, 4) = hojUsu_Summary.Cells(15, 23)
        hojUsu_SetPricesFirewood.Cells(c_ini, 5) = hojUsu_Summary.Cells(k, 84)
'        hojUsu_SetPricesFirewood.Cells(c_ini, 6) = hojUsu_Summary.Cells(k, 86)
'        hojUsu_SetPricesFirewood.Cells(c_ini, 7) = hojUsu_Summary.Cells(k, 88)
        hojUsu_SetPricesFirewood.Cells(c_ini, 8) = hojUsu_Summary.Cells(k, 90)
        c_ini = c_ini + 1
    
    Next k
    
Application.ScreenUpdating = True

End Sub

Sub SOLVER_FINAL_RURAL_CONSUMPTION()



End Sub
