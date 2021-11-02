Attribute VB_Name = "Eq_Reset"
Sub RESET()

Application.ScreenUpdating = False

Dim market As String

market = hojUsu_SystemOptions.Range("MarketsInputs").Value
market = "All"

Range("SelectProcess") = ""
Range("InitialYearRange") = "1975"
If Range("FinalYearRange") = "" Then
    Range("FinalYearRange") = 2015
End If
Range("NegativeData") = ""
Range("Solver") = "No"
Range("VariablesSolver") = ""
Range("OriginForVariablesTwo") = ""
Range("IterationMethod") = ""
Range("InitialYearRangeSolver") = ""
Range("FinalYearRangeSolver") = ""
Range("Report_Export") = ""

Select Case market

    Case "Wood_Industry"
            
        Call RESET_WOOD_INDUSTRY
            
    Case "Furniture_Industry"
    
        Call RESET_FURNITURE_INDUSTRY
            
    Case "Pulp_Paper_Industry"
    
        Call RESET_PULP_PAPER_INDUSTRY
            
    Case "Wood_Industrial"
    
        Call RESET_WOOD_INDUSTRIAL
            
    Case "Firewood"
    
        Call RESET_FIREWOOD
            
    Case "All"
    
        Call RESET_WOOD_INDUSTRY
        Call RESET_FURNITURE_INDUSTRY
        Call RESET_PULP_PAPER_INDUSTRY
        Call RESET_WOOD_INDUSTRIAL
        Call RESET_FIREWOOD
                   
End Select

'años 1973 y 1974 = 1

'Comsumption of wood industry
    hojUsu_Summary.Cells(37, 4).Value = 1
    hojUsu_Summary.Cells(38, 4).Value = 1
'Comsumption price of wood industry
    hojUsu_Summary.Cells(37, 10).Value = 1
    hojUsu_Summary.Cells(38, 10).Value = 1
'Import price of wood industry
    hojUsu_Summary.Cells(37, 14).Value = 1
    hojUsu_Summary.Cells(38, 14).Value = 1
    
'Comsumption of furniture industry
    hojUsu_Summary.Cells(37, 22).Value = 1
    hojUsu_Summary.Cells(38, 22).Value = 1
'Comsumption price of furniture industry
    hojUsu_Summary.Cells(37, 28).Value = 1
    hojUsu_Summary.Cells(38, 28).Value = 1
'Import price of furniture industry
    hojUsu_Summary.Cells(37, 32).Value = 1
    hojUsu_Summary.Cells(38, 32).Value = 1

'Comsumption of pulp and paper industry
    hojUsu_Summary.Cells(37, 40).Value = 1
    hojUsu_Summary.Cells(38, 40).Value = 1
'Comsumption price of pulp and paper industry
    hojUsu_Summary.Cells(37, 46).Value = 1
    hojUsu_Summary.Cells(38, 46).Value = 1
'Import price of pulp and paper industry
    hojUsu_Summary.Cells(37, 50).Value = 1
    hojUsu_Summary.Cells(38, 50).Value = 1
    
'Comsumption of firewood
    hojUsu_Summary.Cells(37, 78).Value = 1
    hojUsu_Summary.Cells(38, 78).Value = 1
'Comsumption price of firewood
    hojUsu_Summary.Cells(37, 84).Value = 1
    hojUsu_Summary.Cells(38, 84).Value = 1

Application.ScreenUpdating = True

End Sub
Sub RESET_WOOD_INDUSTRY()

'Optimización de pantalla y gestión de errores
Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_end, d_end, columSummary, columForecast, columForecastAux As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast y SP)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_end = b - 1967
d_end = b - 1936

'Variables auxiliares
columForecastAux = 0

'Reset Summary
hojUsu_Summary.Activate

For columSummary = 2 To 14

    hojUsu_Summary.Range(Cells(d_ini, columSummary), Cells(d_end, columSummary)).ClearContents
    columSummary = columSummary + 1
    
Next columSummary

'Reset Forecast
hojUsu_Forecast.Activate
For columForecast = 4 To 30
    
    columForecastAux = columForecastAux + 1
    If columForecastAux <> 4 Then
    
        hojUsu_Forecast.Range(Cells(c_ini, columForecast), Cells(c_end, columForecast)).ClearContents
    
    Else
    
        columForecastAux = 0
    
    End If
    
Next columForecast

'Reset set prices
hojUsu_SetPricesWoodIndustry.Activate
hojUsu_SetPricesWoodIndustry.Range(Cells(c_ini, 3), Cells(c_end, 8)).ClearContents

'Reset PSt
Call Restart_PStw_Wood_Industry

hojUsu_SystemOptions.Activate

Application.ScreenUpdating = True

End Sub
Sub RESET_FURNITURE_INDUSTRY()

'Optimización de pantalla y gestión de errores
Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_end, d_end, columSummary, columForecast, columForecastAux As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast y SP)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_end = b - 1967
d_end = b - 1936

'Variables auxiliares
columForecastAux = 0

'Reset Summary
hojUsu_Summary.Activate
For columSummary = 20 To 32

    hojUsu_Summary.Range(Cells(d_ini, columSummary), Cells(d_end, columSummary)).ClearContents
    columSummary = columSummary + 1
    
Next columSummary

'Reset Forecast
hojUsu_Forecast.Activate
For columForecast = 32 To 58
    
    columForecastAux = columForecastAux + 1
    If columForecastAux <> 4 Then
    
        hojUsu_Forecast.Range(Cells(c_ini, columForecast), Cells(c_end, columForecast)).ClearContents
    
    Else
    
        columForecastAux = 0
    
    End If
    
Next columForecast

'Reset set prices
hojUsu_SetPricesFurniture.Activate
hojUsu_SetPricesFurniture.Range(Cells(c_ini, 3), Cells(c_end, 8)).ClearContents

'Reset PSt
Call Restart_PStw_Furniture_Industry

hojUsu_SystemOptions.Activate

Application.ScreenUpdating = True

End Sub
Sub RESET_PULP_PAPER_INDUSTRY()

'Optimización de pantalla y gestión de errores
Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_end, d_end, columSummary, columForecast, columForecastAux As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast y SP)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_end = b - 1967
d_end = b - 1936

'Variables auxiliares
columForecastAux = 0

'Reset Summary
hojUsu_Summary.Activate
For columSummary = 38 To 50

    hojUsu_Summary.Range(Cells(d_ini, columSummary), Cells(d_end, columSummary)).ClearContents
    columSummary = columSummary + 1
    
Next columSummary

'Reset Forecast
hojUsu_Forecast.Activate
For columForecast = 60 To 86
    
    columForecastAux = columForecastAux + 1
    If columForecastAux <> 4 Then
    
        hojUsu_Forecast.Range(Cells(c_ini, columForecast), Cells(c_end, columForecast)).ClearContents
    
    Else
    
        columForecastAux = 0
    
    End If
    
Next columForecast

'Reset set prices
hojUsu_SetPricesPulpPaper.Activate
hojUsu_SetPricesPulpPaper.Range(Cells(c_ini, 3), Cells(c_end, 8)).ClearContents

'Reset PSt
Call Restart_PStw_Pulp_Paper_Industry

hojUsu_SystemOptions.Activate

Application.ScreenUpdating = True

End Sub
Sub RESET_WOOD_INDUSTRIAL()

'Optimización de pantalla y gestión de errores
Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_end, d_end, columSummary, columForecast, columForecastAux As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast y SP)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_end = b - 1967
d_end = b - 1936

'Variables auxiliares
columForecastAux = 0

'Reset Summary
hojUsu_Summary.Activate
For columSummary = 56 To 70

    hojUsu_Summary.Range(Cells(d_ini, columSummary), Cells(d_end, columSummary)).ClearContents
    columSummary = columSummary + 1
    
Next columSummary

'Reset Forecast
hojUsu_Forecast.Activate
For columForecast = 88 To 118
    
    columForecastAux = columForecastAux + 1
    If columForecastAux <> 4 Then
    
        hojUsu_Forecast.Range(Cells(c_ini, columForecast), Cells(c_end, columForecast)).ClearContents
    
    Else
    
        columForecastAux = 0
    
    End If
    
Next columForecast

'Reset set prices
hojUsu_SetPricesWoodIndustrial.Activate
hojUsu_SetPricesWoodIndustrial.Range(Cells(c_ini, 3), Cells(c_end, 8)).ClearContents

'Reset PSt
Call Restart_PStw_Wood_Industrial

hojUsu_SystemOptions.Activate

Application.ScreenUpdating = True

End Sub
Sub RESET_FIREWOOD()

'Optimización de pantalla y gestión de errores
Application.ScreenUpdating = False

Dim a, b, c_ini, d_ini, c_end, d_end, columSummary, columForecast, columForecastAux As Integer

'a = año inicial a ser evaluado
'b = año final a ser evaluado
'c = años donde se depositan los resultados (forecast y SP)
'd = años de la hoja Summary

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
d_ini = a - 1936
c_end = b - 1967
d_end = b - 1936

'Variables auxiliares
columForecastAux = 0

'Reset Summary
hojUsu_Summary.Activate
For columSummary = 76 To 88

    hojUsu_Summary.Range(Cells(d_ini, columSummary), Cells(d_end, columSummary)).ClearContents
    columSummary = columSummary + 1
    
Next columSummary

'Reset Forecast
hojUsu_Forecast.Activate
For columForecast = 120 To 130
    
    columForecastAux = columForecastAux + 1
    If columForecastAux <> 4 Then
    
        hojUsu_Forecast.Range(Cells(c_ini, columForecast), Cells(c_end, columForecast)).ClearContents
    
    Else
    
        columForecastAux = 0
    
    End If
    
Next columForecast

'Reset set prices
hojUsu_SetPricesFirewood.Activate
hojUsu_SetPricesFirewood.Range(Cells(c_ini, 3), Cells(c_end, 8)).ClearContents

'Reset PSt
Call Restart_PStw_Firewood

hojUsu_SystemOptions.Activate

Application.ScreenUpdating = True

End Sub
Sub Restart_PStw_Firewood()

    hojUsu_Firewood.Range("O6:O51").Copy
    hojUsu_Summary.Range("CL37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub
Sub Restart_PStw_Furniture_Industry()

    hojUsu_FurnitureIndustry.Range("O6:O51").Copy
    hojUsu_Summary.Range("AH37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub
Sub Restart_PStw_Pulp_Paper_Industry()

    hojUsu_PulpPaperIndustry.Range("O6:O51").Copy
    hojUsu_Summary.Range("AZ37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
Sub Restart_PStw_Wood_Industrial()

    hojUsu_WoodIndustrial.Range("T6:T51").Copy
    hojUsu_Summary.Range("BT37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
    hojUsu_WoodIndustrial.Range("AE6:AE51").Copy
    hojUsu_Summary.Range("BV37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub
Sub Restart_PStw_Wood_Industry()

    hojUsu_WoodIndustry.Range("O6:O51").Copy
    hojUsu_Summary.Range("P37").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

