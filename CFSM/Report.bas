Attribute VB_Name = "Report"
Sub REPORT()

Dim market, equation As String

'configuración del sistema
processSelection = hojUsu_SystemOptions.Range("SelectProcess")
rangeYearIni = hojUsu_SystemOptions.Range("InitialYearRange")
rangeYearEnd = hojUsu_SystemOptions.Range("FinalYearRange")
negativeData = hojUsu_SystemOptions.Range("NegativeData")
exonerousVariables = hojUsu_SystemOptions.Range("VariablesSolver")
originForSetPrices = hojUsu_SystemOptions.Range("OriginForVariablesTwo")
iterationMethod = hojUsu_SystemOptions.Range("IterationMethod")
rangeYearSolverIni = hojUsu_SystemOptions.Range("InitialYearRangeSolver")
rangeYearSolverEnd = hojUsu_SystemOptions.Range("FinalYearRangeSolver")
market = hojUsu_SystemOptions.Range("MarketsInputs").Value
equation = hojUsu_SystemOptions.Range("EquationsInputs").Value

'configuración del sistema en el momento que se genera el reporte
Select Case processSelection
    Case 0
        hojUsu_Report.Cells(3, 1) = ""
        hojUsu_Report_MCC.Cells(3, 1) = ""
    Case 1
        hojUsu_Report.Cells(3, 1) = "Validation"
        hojUsu_Report_MCC.Cells(3, 1) = "Validation"
    Case 2
        hojUsu_Report.Cells(3, 1) = "Market Clearing Condition"
        hojUsu_Report_MCC.Cells(3, 1) = "Market Clearing Condition"
    Case 3
        hojUsu_Report.Cells(3, 1) = "Historical data"
        hojUsu_Report_MCC.Cells(3, 1) = "Historical data"
    Case 4
        hojUsu_Report.Cells(3, 1) = "Module Industrys - NPW"
        hojUsu_Report_MCC.Cells(3, 1) = "Module Industrys - NPW"
    Case 5
        hojUsu_Report.Cells(3, 1) = "Module NPW - Industries"
        hojUsu_Report_MCC.Cells(3, 1) = "Module NPW - Industries"
End Select

'Años reportados
hojUsu_Report.Cells(5, 1) = rangeYearIni
hojUsu_Report.Cells(6, 1) = rangeYearEnd
hojUsu_Report_MCC.Cells(5, 1) = rangeYearIni
hojUsu_Report_MCC.Cells(6, 1) = rangeYearEnd

'Comportamiento datos negativos
Select Case negativeData
    Case 0
        hojUsu_Report.Cells(8, 1) = ""
        hojUsu_Report_MCC.Cells(8, 1) = ""
    Case 1
        hojUsu_Report.Cells(8, 1) = "Historical data"
        hojUsu_Report_MCC.Cells(8, 1) = "Historical data"
    Case 2
        hojUsu_Report.Cells(8, 1) = "Value = 0"
        hojUsu_Report_MCC.Cells(8, 1) = "Value = 0"
    Case 3
        hojUsu_Report.Cells(8, 1) = "Raw data"
        hojUsu_Report_MCC.Cells(8, 1) = "Raw data"
End Select

Select Case exonerousVariables
    Case 0
        hojUsu_Report.Cells(11, 1) = ""
        hojUsu_Report_MCC.Cells(11, 1) = ""
    Case 1
        hojUsu_Report.Cells(11, 1) = "PSt"
        hojUsu_Report_MCC.Cells(11, 1) = "PSt"
    Case 2
        hojUsu_Report.Cells(11, 1) = "PCt, PXt, PMt, PSt"
        hojUsu_Report_MCC.Cells(11, 1) = "PCt, PXt, PMt, PSt"
End Select

Select Case originForSetPrices
    Case 0
        hojUsu_Report.Cells(13, 1) = ""
        hojUsu_Report_MCC.Cells(13, 1) = ""
    Case 1
        hojUsu_Report.Cells(13, 1) = "Using MCC solver, iteration 0"
        hojUsu_Report_MCC.Cells(13, 1) = "Using MCC solver, iteration 0"
    Case 2
        hojUsu_Report.Cells(13, 1) = "Historical data"
        hojUsu_Report_MCC.Cells(13, 1) = "Historical data"
    Case 3
        hojUsu_Report.Cells(13, 1) = "Estimated data using estimated equation"
        hojUsu_Report_MCC.Cells(13, 1) = "Estimated data using estimated equation"
End Select

Select Case iterationMethod
    Case 0
        hojUsu_Report.Cells(15, 1) = ""
        hojUsu_Report_MCC.Cells(15, 1) = ""
    Case 1
        hojUsu_Report.Cells(15, 1) = "GRG Nonlinear"
        hojUsu_Report_MCC.Cells(15, 1) = "GRG Nonlinear"
    Case 2
        hojUsu_Report.Cells(15, 1) = "Simplex LP"
        hojUsu_Report_MCC.Cells(15, 1) = "Simplex LP"
    Case 3
        hojUsu_Report.Cells(15, 1) = "Evolutionary"
        hojUsu_Report_MCC.Cells(15, 1) = "Evolutionary"
End Select

'años rango solver
hojUsu_Report.Cells(17, 1) = rangeYearSolverIni
hojUsu_Report.Cells(18, 1) = rangeYearSolverEnd

hojUsu_Report_MCC.Cells(17, 1) = rangeYearSolverIni
hojUsu_Report_MCC.Cells(18, 1) = rangeYearSolverEnd

'mercado y equation general
hojUsu_Report.Cells(20, 1) = market
hojUsu_Report.Cells(22, 1) = equation

hojUsu_Report_MCC.Cells(20, 1) = market

'resetear el informe anterior
If Range("FinalYearRange") <> "" Then
    hojUsu_Report.Activate
    rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

    hojUsu_Report.Range(Cells(4, 2), Cells(rowDataActu, 6)).Clear
End If

Select Case market

    Case "Wood_Industry"

        Select Case equation

            Case "Supply"

            Call REPORT_SUPPLY_WOOD_INDUSTRY

            Case "Consumption"

            Call REPORT_CONSUMPTION_WOOD_INDUSTRY

            Case "Exports"

            Call REPORT_EXPORTS_WOOD_INDUSTRY

            Case "Imports"

            Call REPORT_IMPORTS_WOOD_INDUSTRY

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRY

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRY

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRY

        End Select

    Case "Furniture_Industry"

        Select Case equation

            Case "Supply"

            Call REPORT_SUPPLY_FURNITURE_INDUSTRY

            Case "Consumption"

            Call REPORT_CONSUMPTION_FURNITURE_INDUSTRY

            Case "Exports"

            Call REPORT_EXPORTS_FURNITURE_INDUSTRY

            Case "Imports"

            Call REPORT_IMPORTS_FURNITURE_INDUSTRY

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_FURNITURE_INDUSTRY

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_FURNITURE_INDUSTRY

        End Select

    Case "Pulp_Paper_Industry"

        Select Case equation

            Case "Supply"

            Call REPORT_SUPPLY_PULP_PAPER_INDUSTRY

            Case "Consumption"

            Call REPORT_CONSUMPTION_PULP_PAPER_INDUSTRY

            Case "Exports"

            Call REPORT_EXPORTS_PULP_PAPER_INDUSTRY

            Case "Imports"

            Call REPORT_IMPORTS_PULP_PAPER_INDUSTRY

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY

        End Select

    Case "Wood_Industrial"

        Select Case equation

            Case "Supply forest plantations"
            
            Call REPORT_SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
            
            Case "Supply natural forest"
            
            Call REPORT_SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST

            Case "Consumption"

            Call REPORT_CONSUMPTION_WOOD_INDUSTRIAL

            Case "Exports"

            Call REPORT_EXPORTS_WOOD_INDUSTRIAL

            Case "Imports"

            Call REPORT_IMPORTS_WOOD_INDUSTRIAL

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRIAL

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRIAL

        End Select

    Case "Firewood"

        Select Case equation

            Case "Supply"

            Call REPORT_SUPPLY_FIREWOOD

            Case "Consumption"

            Call REPORT_CONSUMPTION_FIREWOOD

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_FIREWOOD

        End Select

    Case "Set_prices"

        Select Case equation

            Case "SP_Wood_Industry"

            Call REPORT_SP_WOOD_INDUSTRY

            Case "SP_Furniture_Industry"

            Call REPORT_SP_FURNITURE_INDUSTRY

            Case "SP_Pulp_Paper_Industry"

            Call REPORT_SP_PULP_PAPER_INDUSTRY

            Case "SP_Wood_Industrial"

            Call REPORT_SP_WOOD_INDUSTRIAL

            Case "SP_Firewood"

            Call REPORT_SP_FIREWOOD

        End Select

    Case "MCC"

        Select Case equation

            Case "MCC_Wood_Industry"

            Call REPORT_MCC_WOOD_INDUSTRY

            Case "MCC_Furniture_Industry"

            Call REPORT_MCC_FURNITURE_INDUSTRY

            Case "MCC_Pulp_Paper_Industry"

            Call REPORT_MCC_PULP_PAPER_INDUSTRY

            Case "MCC_Wood_Industrial"

            Call REPORT_MCC_INDUSTRIAL_WOOD

            Case "MCC_Firewood"

            Call REPORT_MCC_FIREWOOD
            
            Case "MCC_MWM"

            Call REPORT_MCC_MWM

            Case "MCC_UWM"

            Call REPORT_MCC_UWM

            Case "MCC_CFSM"

            Call REPORT_MCC_CFSM

        End Select
        
End Select

hojUsu_SystemOptions.Activate

End Sub
Sub REPORT_SUPPLY_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 3), Cells(3, 6)).Copy
hojUsu_Report.Activate

Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 3), Cells(c_end, 6)).Copy
hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of " & _
    "the wood industry - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_CONSUMPTION_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 7), Cells(3, 10)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 7), Cells(c_end, 10)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the wood industry - Ctw (2015 thousand million COP)"
    
'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_EXPORTS_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 11), Cells(3, 14)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 11), Cells(c_end, 14)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the wood industry - Xtw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_IMPORTS_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 15), Cells(3, 18)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 15), Cells(c_end, 18)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the wood industry - Mtw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 19), Cells(3, 22)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 19), Cells(c_end, 22)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption price of manufactured wood products of the wood industry - PCtw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 23), Cells(3, 26)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 23), Cells(c_end, 26)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Export price of manufactured wood products of the wood industry - PXtw (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 27), Cells(3, 30)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 27), Cells(c_end, 30)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Import price of manufactured wood products of the wood industry - PMtw (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SUPPLY_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 31), Cells(3, 34)).Copy
hojUsu_Report.Activate

Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 31), Cells(c_end, 34)).Copy
hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the furniture industry - Stf (Total gross production of the furniture industry, 2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_CONSUMPTION_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 35), Cells(3, 38)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 35), Cells(c_end, 38)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the furniture industry - Ctf (2015 thousand million COP of consumption of manufactured wood products of the furniture industry)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_EXPORTS_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 39), Cells(3, 42)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 39), Cells(c_end, 42)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the furniture industry - Xtf (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_IMPORTS_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 43), Cells(3, 46)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 43), Cells(c_end, 46)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the furniture industry - Mtf (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 47), Cells(3, 50)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 47), Cells(c_end, 50)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption price of manufactured wood products of the furniture industry - PCtf (Unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_EXPORTS_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 51), Cells(3, 54)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 51), Cells(c_end, 54)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Export price of manufactured wood products of the furniture industry - PXtf (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_IMPORT_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 55), Cells(3, 58)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 55), Cells(c_end, 58)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Import price of manufactured wood products of the furniture industry - PMtf (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SUPPLY_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 59), Cells(3, 62)).Copy
hojUsu_Report.Activate

Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 59), Cells(c_end, 62)).Copy
hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the pulp and paper industry - Stz (Total gross production of the pulp and paper industry, 2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_CONSUMPTION_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 63), Cells(3, 66)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 63), Cells(c_end, 66)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the pulp and paper industry - Ctz (2015 Thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_EXPORTS_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 67), Cells(3, 70)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 67), Cells(c_end, 70)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the pulp and paper industry - Xtz (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_IMPORTS_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 71), Cells(3, 74)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 71), Cells(c_end, 74)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the pulp and paper industry- Mtz (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 75), Cells(3, 78)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 75), Cells(c_end, 78)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption price of manufactured wood products of the pulp and paper industry - PCtz (Unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 79), Cells(3, 82)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 79), Cells(c_end, 82)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Export price of manufactured wood products of the pulp and paper industry - PXtz (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 83), Cells(3, 86)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 83), Cells(c_end, 86)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Import price of manufactured wood products of the pulp and paper industry - PMtz (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 87), Cells(3, 90)).Copy
hojUsu_Report.Activate

Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 87), Cells(c_end, 90)).Copy
hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Supply of unprocessed wood for the manufactured wood products industry - MW  from Forest Plantations - StMWfprw (Gross production MWrw, weighted by volumen source, 2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 91), Cells(3, 94)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 91), Cells(c_end, 94)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Supply of unprocessed wood for the manufactured wood products industry - MW  from natural forest - fn - StMWrw (Gross production MWrw, weighted by volumen source, 2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_CONSUMPTION_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 95), Cells(3, 98)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 95), Cells(c_end, 98)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption of unprocessed wood in the manufactured wood products industry - CtMWrw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_EXPORTS_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 99), Cells(3, 102)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 99), Cells(c_end, 102)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Exports of unprocessed wood for the manufactured wood products industry - XtMWrw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_IMPORTS_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 103), Cells(3, 106)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 103), Cells(c_end, 106)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Imports of unprocessed wood for the manufactured wood products industry - MtMWrw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 107), Cells(3, 110)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 107), Cells(c_end, 110)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption price of unprocessed wood in the manufactured wood products industry - PCtMWrw (Unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 111), Cells(3, 114)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 111), Cells(c_end, 114)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Export price of unprocessed wood for the manufactured wood products industry - PXtMWrw (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 115), Cells(3, 118)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 115), Cells(c_end, 118)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Import price of unprocessed wood for the manufactured wood products industry - PMtMWrw (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SUPPLY_FIREWOOD()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 119), Cells(3, 122)).Copy
hojUsu_Report.Activate

Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 119), Cells(c_end, 122)).Copy
hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Supply of unprocessed wood for firewood - FW from natural forest - fn - StFWnfrw (Total gross production of unprocessed wood for firewood, 2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_CONSUMPTION_FIREWOOD()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 123), Cells(3, 126)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 123), Cells(c_end, 126)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Consumption of unprocessed wood for firewood - CtFWrw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_PRICE_OF_CONSUMPTION_FIREWOOD()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 127), Cells(3, 130)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 127), Cells(c_end, 130)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Price of consumption of unprocessed wood in firewood - PCtFWrw (Unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 131), Cells(3, 134)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 131), Cells(c_end, 134)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Wood Industry (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

Call REPORT_MCC_WOOD_INDUSTRY_EXTENDED

End Sub
Sub REPORT_MCC_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 135), Cells(3, 138)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 135), Cells(c_end, 138)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Furniture Industry (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Call REPORT_MCC_FURNITURE_INDUSTRY_EXTENDED

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 139), Cells(3, 142)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 139), Cells(c_end, 142)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Pulp Paper Industry (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Call REPORT_MCC_PULP_PAPER_INDUSTRY_EXTENDED

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_INDUSTRIAL_WOOD()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 143), Cells(3, 146)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 143), Cells(c_end, 146)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Wood Industrial (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Call REPORT_MCC_INDUSTRIAL_WOOD_EXTENDED

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_FIREWOOD()

Application.ScreenUpdating = False

'titulo
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(1, 147), Cells(3, 150)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_Forecast.Activate
hojUsu_Forecast.Range(Cells(c_ini, 147), Cells(c_end, 150)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Firewood (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Call REPORT_MCC_FIREWOOD_EXTENDED

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_MWM()

Application.ScreenUpdating = False

'titulo
hojUsu_MCC.Activate
hojUsu_MCC.Range(Cells(1, 119), Cells(3, 122)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_MCC.Activate
hojUsu_MCC.Range(Cells(c_ini, 119), Cells(c_end, 122)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Manufactured Wood Products Markets - MWM (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Call REPORT_MCC_MWM_EXTENDED

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_UWM()

Application.ScreenUpdating = False

'titulo
hojUsu_MCC.Activate
hojUsu_MCC.Range(Cells(1, 143), Cells(3, 146)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_MCC.Activate
hojUsu_MCC.Range(Cells(c_ini, 143), Cells(c_end, 146)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Unprocessed Wood Markets - UWM (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Call REPORT_MCC_UWM_EXTENDED

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_CFSM()

Application.ScreenUpdating = False

'titulo
hojUsu_MCC.Activate
hojUsu_MCC.Range(Cells(1, 167), Cells(3, 170)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
hojUsu_MCC.Activate
hojUsu_MCC.Range(Cells(c_ini, 167), Cells(c_end, 170)).Copy

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Market Clearing Condition Colombian Forest Sector Model - CFSM (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Call REPORT_MCC_CFSM_EXTENDED

Application.ScreenUpdating = True

End Sub
Sub REPORT_SP_WOOD_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_SetPricesWoodIndustry.Activate
hojUsu_SetPricesWoodIndustry.Range(Cells(1, 5), Cells(3, 8)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
If hojUsu_SystemOptions.Range("SelectProcess") <> 3 Then
    hojUsu_SetPricesWoodIndustry.Activate
    hojUsu_SetPricesWoodIndustry.Range(Cells(c_ini, 5), Cells(c_end, 8)).Copy
Else
    hojUsu_SetPricesWoodIndustry.Activate
    hojUsu_SetPricesWoodIndustry.Range(Cells(c_ini, 12), Cells(c_end, 15)).Copy
End If

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Set prices of wood industry (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SP_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_SetPricesFurniture.Activate
hojUsu_SetPricesFurniture.Range(Cells(1, 5), Cells(3, 8)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
If hojUsu_SystemOptions.Range("SelectProcess") <> 3 Then
    hojUsu_SetPricesFurniture.Activate
    hojUsu_SetPricesFurniture.Range(Cells(c_ini, 5), Cells(c_end, 8)).Copy
Else
    hojUsu_SetPricesFurniture.Activate
    hojUsu_SetPricesFurniture.Range(Cells(c_ini, 12), Cells(c_end, 15)).Copy
End If

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Set prices of furniture industry (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SP_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

'titulo
hojUsu_SetPricesPulpPaper.Activate
hojUsu_SetPricesPulpPaper.Range(Cells(1, 5), Cells(3, 8)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
If hojUsu_SystemOptions.Range("SelectProcess") <> 3 Then
    hojUsu_SetPricesPulpPaper.Activate
    hojUsu_SetPricesPulpPaper.Range(Cells(c_ini, 5), Cells(c_end, 8)).Copy
Else
    hojUsu_SetPricesPulpPaper.Activate
    hojUsu_SetPricesPulpPaper.Range(Cells(c_ini, 12), Cells(c_end, 15)).Copy
End If

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Set prices of pulp and paper industry (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SP_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

'titulo
hojUsu_SetPricesWoodIndustrial.Activate
hojUsu_SetPricesWoodIndustrial.Range(Cells(1, 5), Cells(3, 8)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
If hojUsu_SystemOptions.Range("SelectProcess") <> 3 Then
    hojUsu_SetPricesWoodIndustrial.Activate
    hojUsu_SetPricesWoodIndustrial.Range(Cells(c_ini, 5), Cells(c_end, 8)).Copy
Else
    hojUsu_SetPricesWoodIndustrial.Activate
    hojUsu_SetPricesWoodIndustrial.Range(Cells(c_ini, 12), Cells(c_end, 15)).Copy
End If

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Set prices of wood industrial (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_SP_FIREWOOD()

Application.ScreenUpdating = False

'titulo
hojUsu_SetPricesFirewood.Activate
hojUsu_SetPricesFirewood.Range(Cells(1, 5), Cells(3, 8)).Copy

hojUsu_Report.Activate
Range("C1:F1").Select
ActiveSheet.Paste

'Rango años evaluados

a = hojUsu_SystemOptions.Range("InitialYearRange")
b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

c_ini = a - 1967
c_end = b - 1967

'años
Call YearsSimpleReport

'datos
If hojUsu_SystemOptions.Range("SelectProcess") <> 3 Then
    hojUsu_SetPricesFirewood.Activate
    hojUsu_SetPricesFirewood.Range(Cells(c_ini, 5), Cells(c_end, 8)).Copy
Else
    hojUsu_SetPricesFirewood.Activate
    hojUsu_SetPricesFirewood.Range(Cells(c_ini, 12), Cells(c_end, 15)).Copy
End If

hojUsu_Report.Activate
hojUsu_Report.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "Set prices of firewood (unitless, 2015 = 1)"

'copiar formato celdas
rowDataActu = hojUsu_Report.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report.Range(Cells(3, 2), Cells(3, 6)).Copy
hojUsu_Report.Range(Cells(3, 2), Cells(rowDataActu, 6)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

hojUsu_Report.ChartObjects("Gráfico 2").Activate
    ActiveChart.SetSourceData Source:=Range(Cells(2, 2), Cells(rowDataActu, 6))

hojUsu_Report.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
