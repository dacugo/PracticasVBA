Attribute VB_Name = "Report_extended"
Sub REPORT_MCC_WOOD_INDUSTRY_EXTENDED()

Application.ScreenUpdating = False

hojUsu_Report_MCC.Activate

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 3), Cells(3, 6)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste
    
    'demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 151), Cells(3, 154)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste
    
    'consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 7), Cells(3, 10)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste
    
    'exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 11), Cells(3, 14)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 15), Cells(1, 18)).Select
    ActiveSheet.Paste
    
    'importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 15), Cells(3, 18)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 19), Cells(1, 22)).Select
    ActiveSheet.Paste
    
    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 19), Cells(3, 22)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 23), Cells(1, 26)).Select
    ActiveSheet.Paste
    
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 23), Cells(3, 26)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 27), Cells(1, 30)).Select
    ActiveSheet.Paste
    
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 27), Cells(3, 30)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 31), Cells(1, 34)).Select
    ActiveSheet.Paste
    
    'precio de la Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 171), Cells(3, 172)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 35), Cells(1, 36)).Select
    ActiveSheet.Paste
    
    'Market Clearing condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 131), Cells(3, 134)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados
    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados
    c_ini = a - 1967
    c_end = b - 1967
    
'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear

'años
Call YearsMCCReport

'datos
    'Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 3), Cells(c_end, 6)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 151), Cells(c_end, 154)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 7), Cells(c_end, 10)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 11), Cells(c_end, 14)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 15).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 15), Cells(c_end, 18)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 19).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 19), Cells(c_end, 22)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 23).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 23), Cells(c_end, 26)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 27).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 27), Cells(c_end, 30)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 31).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 171), Cells(c_end, 172)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 35).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Market Clearing Condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 131), Cells(c_end, 134)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Wood Industry at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Wood Industry (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Prices of Wood Industry at final iteration for Market Clearing Condition MCC (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Wood Industry - Historical Data (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Wood Industry - Estimated data using estimated equations (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Wood Industry at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Wood Industry (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the Wood Industry at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the Wood Industry - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Wood Industry at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Wood Industry - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Wood Industry at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Wood Industry - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the Wood Industry at final iteration for Market Clearing Condition MCC - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the Wood Industry - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the Wood Industry at final iteration for Market Clearing Condition MCC - Mtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the Wood Industry - Mtw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AD$2:$AD$" & rowDataActu & _
    ",'Report MCC'!$AH$2:$AH$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AA$2:$AA$" & rowDataActu & _
    ",'Report MCC'!$AE$2:$AE$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AB$2:$AB$" & rowDataActu & _
    ",'Report MCC'!$AF$2:$AF$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)

hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_FURNITURE_INDUSTRY_EXTENDED()

Application.ScreenUpdating = False

hojUsu_Report_MCC.Activate

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 31), Cells(3, 34)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste
    
    'demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 155), Cells(3, 158)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste
    
    'consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 35), Cells(3, 38)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste
    
    'exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 39), Cells(3, 42)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 15), Cells(1, 18)).Select
    ActiveSheet.Paste
    
    'importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 43), Cells(3, 46)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 19), Cells(1, 22)).Select
    ActiveSheet.Paste
    
    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 47), Cells(3, 50)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 23), Cells(1, 26)).Select
    ActiveSheet.Paste
    
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 51), Cells(3, 54)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 27), Cells(1, 30)).Select
    ActiveSheet.Paste
    
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 55), Cells(3, 58)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 31), Cells(1, 34)).Select
    ActiveSheet.Paste
    
    'precio de la Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 173), Cells(3, 174)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 35), Cells(1, 36)).Select
    ActiveSheet.Paste
    
    'Market Clearing condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 135), Cells(3, 138)).Copy

    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados

    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

    c_ini = a - 1967
    c_end = b - 1967
    
'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear
    
'años
    Call YearsMCCReport

'datos
    'Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 31), Cells(c_end, 34)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

    'Demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 155), Cells(c_end, 158)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 35), Cells(c_end, 38)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 39), Cells(c_end, 42)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 15).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 43), Cells(c_end, 46)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 19).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 47), Cells(c_end, 50)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 23).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 51), Cells(c_end, 54)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 27).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 55), Cells(c_end, 58)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 31).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 173), Cells(c_end, 174)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 35).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
                
    'Market Clearing Condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 135), Cells(c_end, 138)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Furniture Industry at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Furniture Industry (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Prices of Furniture Industry at final iteration for Market Clearing Condition MCC (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Furniture Industry - Historical Data (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Furniture Industry - Estimated data using estimated equations (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Furniture Industry at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Furniture Industry (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the Furniture Industry at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the Furniture Industry - Stf (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Furniture Industry at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Furniture Industry - Dtf (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Furniture Industry at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Furniture Industry - Ctf (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the Furniture Industry at final iteration for Market Clearing Condition MCC - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the Furniture Industry - Xtf (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the Furniture Industry at final iteration for Market Clearing Condition MCC - Mtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the Furniture Industry - Mtf (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AD$2:$AD$" & rowDataActu & _
    ",'Report MCC'!$AH$2:$AH$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AA$2:$AA$" & rowDataActu & _
    ",'Report MCC'!$AE$2:$AE$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AB$2:$AB$" & rowDataActu & _
    ",'Report MCC'!$AF$2:$AF$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_PULP_PAPER_INDUSTRY_EXTENDED()

Application.ScreenUpdating = False

hojUsu_Report_MCC.Activate

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 59), Cells(3, 62)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste
    
    'demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 159), Cells(3, 162)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste
    
    'consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 63), Cells(3, 66)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste
    
    'exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 67), Cells(3, 70)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 15), Cells(1, 18)).Select
    ActiveSheet.Paste
    
    'importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 71), Cells(3, 74)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 19), Cells(1, 22)).Select
    ActiveSheet.Paste

    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 75), Cells(3, 78)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 23), Cells(1, 26)).Select
    ActiveSheet.Paste
    
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 79), Cells(3, 82)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 27), Cells(1, 30)).Select
    ActiveSheet.Paste
    
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 83), Cells(3, 86)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 31), Cells(1, 34)).Select
    ActiveSheet.Paste
    
    'precio de la Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 175), Cells(3, 176)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 35), Cells(1, 36)).Select
    ActiveSheet.Paste

    'Market Clearing condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 139), Cells(2, 142)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados
    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados
    c_ini = a - 1967
    c_end = b - 1967

'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear

'años
    Call YearsMCCReport

'datos
    'Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 59), Cells(c_end, 62)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'Demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 159), Cells(c_end, 162)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 63), Cells(c_end, 66)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 67), Cells(c_end, 70)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 15).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 71), Cells(c_end, 74)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 19).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 75), Cells(c_end, 78)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 23).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 79), Cells(c_end, 82)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 27).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 83), Cells(c_end, 86)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 31).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 175), Cells(c_end, 176)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 35).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Market Clearing Condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 139), Cells(c_end, 142)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Pulp and paper Industry at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Pulp and paper Industry (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Prices of Pulp and paper Industry at final iteration for Market Clearing Condition MCC (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Pulp and paper Industry - Historical Data (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Pulp and paper Industry - Estimated data using estimated equations (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Pulp and paper Industry at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Pulp and paper Industry (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the Pulp and paper Industry at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the Pulp and paper Industry - Stz (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Pulp and paper Industry at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Pulp and paper Industry - Dtz (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Pulp and paper Industry at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Pulp and paper Industry - Ctz (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the Pulp and paper Industry at final iteration for Market Clearing Condition MCC - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "Exports of manufactured wood products of the Pulp and paper Industry - Xtz (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the Pulp and paper Industry at final iteration for Market Clearing Condition MCC - Mtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "Imports of manufactured wood products of the Pulp and paper Industry - Mtz (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AD$2:$AD$" & rowDataActu & _
    ",'Report MCC'!$AH$2:$AH$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AA$2:$AA$" & rowDataActu & _
    ",'Report MCC'!$AE$2:$AE$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AB$2:$AB$" & rowDataActu & _
    ",'Report MCC'!$AF$2:$AF$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)

hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_INDUSTRIAL_WOOD_EXTENDED()

Application.ScreenUpdating = False

hojUsu_Report_MCC.Activate

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 181), Cells(3, 184)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste
    
    'demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 163), Cells(3, 166)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste
    
    'consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 95), Cells(3, 98)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste
    
    'exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 99), Cells(3, 102)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 15), Cells(1, 18)).Select
    ActiveSheet.Paste
    
    'importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 103), Cells(3, 106)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 19), Cells(1, 22)).Select
    ActiveSheet.Paste

    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 107), Cells(3, 110)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 23), Cells(1, 26)).Select
    ActiveSheet.Paste
    
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 111), Cells(3, 114)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 27), Cells(1, 30)).Select
    ActiveSheet.Paste
    
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 115), Cells(3, 118)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 31), Cells(1, 34)).Select
    ActiveSheet.Paste
    
    'precio de la Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 177), Cells(3, 178)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 35), Cells(1, 36)).Select
    ActiveSheet.Paste

    'Market Clearing condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 143), Cells(3, 146)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados
    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados
    c_ini = a - 1967
    c_end = b - 1967

'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear

'años
    Call YearsMCCReport

'datos

    'Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 181), Cells(c_end, 184)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'Demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 163), Cells(c_end, 166)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 95), Cells(c_end, 98)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 99), Cells(c_end, 102)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 15).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 103), Cells(c_end, 106)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 19).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 107), Cells(c_end, 110)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 23).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la exportación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 111), Cells(c_end, 114)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 27).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la importación
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 115), Cells(c_end, 118)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 31).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 177), Cells(c_end, 178)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 35).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Market Clearing Condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 143), Cells(c_end, 146)).Copy

    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Industrial Wood at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Industrial Wood (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Prices of Industrial Wood at final iteration for Market Clearing Condition MCC (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Industrial Wood - Historical Data (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Industrial Wood - Estimated data using estimated equations (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Unprocessed Wood of the Industrial Wood at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Unprocessed Wood of the Industrial Wood (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of Unprocessed Wood of the Industrial Wood at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of Unprocessed Wood of the Industrial Wood - StMWrw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of Unprocessed Wood of the Industrial Wood at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of Unprocessed Wood of the Industrial Wood - DtMWrw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Unprocessed Wood of the Industrial Wood at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Unprocessed Wood of the Industrial Wood - CtMWrw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Exports of Unprocessed Wood of the Industrial Wood at final iteration for Market Clearing Condition MCC - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "Exports of Unprocessed Wood of the Industrial Wood - XtMWrw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Imports of Unprocessed Wood of the Industrial Wood at final iteration for Market Clearing Condition MCC - Mtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "Imports of Unprocessed Wood of the Industrial Wood - MtMWrw (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AD$2:$AD$" & rowDataActu & _
    ",'Report MCC'!$AH$2:$AH$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AA$2:$AA$" & rowDataActu & _
    ",'Report MCC'!$AE$2:$AE$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AB$2:$AB$" & rowDataActu & _
    ",'Report MCC'!$AF$2:$AF$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)

hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_FIREWOOD_EXTENDED()

Application.ScreenUpdating = False

hojUsu_Report_MCC.Activate

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 119), Cells(3, 122)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste
    
    'demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 167), Cells(3, 170)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste
    
    'consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 123), Cells(3, 126)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste

    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 127), Cells(3, 130)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 23), Cells(1, 26)).Select
    ActiveSheet.Paste
    
    'precio de la Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 179), Cells(3, 180)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 35), Cells(1, 36)).Select
    ActiveSheet.Paste

    'Market Clearing condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(1, 147), Cells(3, 150)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados

    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

    c_ini = a - 1967
    c_end = b - 1967

'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear

'años
    Call YearsMCCReport

'datos

    'Oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 119), Cells(c_end, 122)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Demanda
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 167), Cells(c_end, 170)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 123), Cells(c_end, 126)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio del consumo
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 127), Cells(c_end, 130)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 23).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'precio de la oferta
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 179), Cells(c_end, 180)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 35).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Market Clearing Condition
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 147), Cells(c_end, 150)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Firewood at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Firewood - FWrw(2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Prices of Firewood at final iteration for Market Clearing Condition MCC (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Firewood - Historical Data (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "Prices of Firewood - Estimated data using estimated equations (Uniless, 2015 = 1)"
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Firewood at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of manufactured wood products of the Firewood (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of manufactured wood products of the Firewood at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of natural forest for the Firewood - StFWnfrw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Firewood at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of manufactured wood products of the Firewood - DtFWrw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Firewood at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of manufactured wood products of the Firewood - CtFWrw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "."

'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)

hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

'Borrar contenido de columnas
    Columns("O:O").Select
    Selection.Clear
    
    Columns("S:S").Select
    Selection.Clear
    
    Columns("AA:AA").Select
    Selection.Clear
    
    Columns("AE:AE").Select
    Selection.Clear

'esconder columnas sin datos
    Columns("O:V").EntireColumn.Hidden = True
    Columns("AA:AH").EntireColumn.Hidden = True
    Columns("CX:DQ").EntireColumn.Hidden = True
Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_MWM_EXTENDED()

Application.ScreenUpdating = False

hojUsu_Report_MCC.Activate

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 99), Cells(3, 102)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste
    
    'demanda
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 115), Cells(3, 118)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste
    
    'consumo
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 103), Cells(3, 106)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste
    
    'exportación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 107), Cells(3, 110)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 15), Cells(1, 18)).Select
    ActiveSheet.Paste
    
    'importación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 111), Cells(3, 114)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 19), Cells(1, 22)).Select
    ActiveSheet.Paste
    
    'Market Clearing condition
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 119), Cells(3, 122)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados
    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados
    c_ini = a - 1967
    c_end = b - 1967

'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear

'años
    Call YearsMCCReport
    
'datos
    'Oferta
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 99), Cells(c_end, 102)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Demanda
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 115), Cells(c_end, 118)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 103), Cells(c_end, 106)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Exportación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 107), Cells(c_end, 110)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 15).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Importación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 111), Cells(c_end, 114)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 19).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Market Clearing Condition
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 119), Cells(c_end, 122)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Manufactured Wood Products Markets at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Manufactured Wood Products Markets (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Manufactured Wood Products Markets at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Manufactured Wood Products Markets (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of Manufactured Wood Products Markets at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of Manufactured Wood Products Markets - StMWM (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of Manufactured Wood Products Markets at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of Manufactured Wood Products Markets - DtMWM (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Manufactured Wood Products Markets at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Manufactured Wood Products Markets - CtMWM (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Exports of Manufactured Wood Products Markets at final iteration for Market Clearing Condition MCC - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "Exports of Manufactured Wood Products Markets - XtMWM (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Imports of Manufactured Wood Products Markets at final iteration for Market Clearing Condition MCC - Mtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "Imports of Manufactured Wood Products Markets - MtMWM (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

'Borrar contenido de columnas
    Columns("W:W").Select
    Selection.Clear
    
    Columns("AA:AA").Select
    Selection.Clear
    
    Columns("AE:AE").Select
    Selection.Clear
    
    Columns("AI:AI").Select
    Selection.Clear
    
'esconder columnas sin datos
    Columns("W:AJ").EntireColumn.Hidden = True
    Columns("AZ:BI").EntireColumn.Hidden = True

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_UWM_EXTENDED()

Application.ScreenUpdating = False

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 123), Cells(3, 126)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste

    'demanda
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 139), Cells(3, 142)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste

    'consumo
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 127), Cells(3, 130)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste

    'exportación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 131), Cells(3, 134)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 15), Cells(1, 18)).Select
    ActiveSheet.Paste

    'importación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 135), Cells(3, 138)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 19), Cells(1, 22)).Select
    ActiveSheet.Paste

    'Market Clearing condition
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 143), Cells(3, 146)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados
    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados
    c_ini = a - 1967
    c_end = b - 1967
    
'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear

'años
    Call YearsMCCReport

'datos
    'Oferta
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 123), Cells(c_end, 126)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Demanda
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 139), Cells(c_end, 142)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 127), Cells(c_end, 130)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Exportación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 131), Cells(c_end, 134)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 15).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Importación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 135), Cells(c_end, 138)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 19).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Market Clearing Condition
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 143), Cells(c_end, 146)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Unprocessed Wood Markets at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Unprocessed Wood Markets (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Unprocessed Wood Markets at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Unprocessed Wood Markets (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of Unprocessed Wood Markets at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of Unprocessed Wood Markets - StUWM (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of Unprocessed Wood Markets at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of Unprocessed Wood Markets - DtUWM (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Unprocessed Wood Markets at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Unprocessed Wood Markets - CtUWM (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Exports of Unprocessed Wood Markets at final iteration for Market Clearing Condition MCC - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "Exports of Unprocessed Wood Markets - XtUWM (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Imports of Unprocessed Wood Markets at final iteration for Market Clearing Condition MCC - Mtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "Imports of Unprocessed Wood Markets - MtUWM (2015 thousand million COP)"
    
'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

'Borrar contenido de columnas
    Columns("W:W").Select
    Selection.Clear
    
    Columns("AA:AA").Select
    Selection.Clear
    
    Columns("AE:AE").Select
    Selection.Clear
    
    Columns("AI:AI").Select
    Selection.Clear
    
'esconder columnas sin datos
    Columns("W:AJ").EntireColumn.Hidden = True
    Columns("AZ:BI").EntireColumn.Hidden = True

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_CFSM_EXTENDED()

Application.ScreenUpdating = False

hojUsu_Report_MCC.Activate

'activar todas las columnas
    Columns("A:DR").EntireColumn.Hidden = False

'titulos
    'oferta
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 147), Cells(3, 150)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 3), Cells(1, 6)).Select
    ActiveSheet.Paste

    'demanda
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 163), Cells(3, 166)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 7), Cells(1, 10)).Select
    ActiveSheet.Paste

    'consumo
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 151), Cells(3, 154)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 11), Cells(1, 14)).Select
    ActiveSheet.Paste

    'exportación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 155), Cells(3, 158)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 15), Cells(1, 18)).Select
    ActiveSheet.Paste

    'importación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 159), Cells(3, 162)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 19), Cells(1, 22)).Select
    ActiveSheet.Paste

    'Market Clearing condition
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(1, 167), Cells(3, 170)).Copy
    
    hojUsu_Report_MCC.Activate
    Range(Cells(1, 37), Cells(1, 40)).Select
    ActiveSheet.Paste

'Rango años evaluados

    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años procesos y resultados

    c_ini = a - 1967
    c_end = b - 1967

'Reset informe anterior
    hojUsu_Report_MCC.Activate
    DataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row
    
    hojUsu_Report_MCC.Range(Cells(4, 2), Cells(DataActu, 40)).Clear

'años
    Call YearsMCCReport

'datos
    'Oferta
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 147), Cells(c_end, 150)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Demanda
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 163), Cells(c_end, 166)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Consumo
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 151), Cells(c_end, 154)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 11).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Exportación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 155), Cells(c_end, 158)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 15).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Importación
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 159), Cells(c_end, 162)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 19).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Market Clearing Condition
    hojUsu_MCC.Activate
    hojUsu_MCC.Range(Cells(c_ini, 167), Cells(c_end, 170)).Copy
    
    hojUsu_Report_MCC.Activate
    hojUsu_Report_MCC.Cells(3, 37).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       
'cambiar nombre gráfica
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Colombian Forest Sector Model at final iteration for Market Clearing Condition MCC (2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.ChartTitle.Text = "Supply minus Demand of Colombian Forest Sector Model (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.ChartTitle.Text = "."
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Colombian Forest Sector Model at final iteration for Market Clearing Condition MCC (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.ChartTitle.Text = "Demand vs Supply of Colombian Forest Sector Model (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Supply of Colombian Forest Sector Model at final iteration for Market Clearing Condition MCC - Stw (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.ChartTitle.Text = "Supply of Colombian Forest Sector Model - StCFSM (Total gross production of the wood industry, 2015 thousand million COP)"

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Demand of Colombian Forest Sector Model at final iteration for Market Clearing Condition MCC - Dtw (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.ChartTitle.Text = "Demand of Colombian Forest Sector Model - DtCFSM (Total gross production of the wood industry, 2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Colombian Forest Sector Model at final iteration for Market Clearing Condition MCC - Ctw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.ChartTitle.Text = "Consumption of Colombian Forest Sector Model - CtCFSM (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Exports of Colombian Forest Sector Model at final iteration for Market Clearing Condition MCC - Xtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.ChartTitle.Text = "Exports of Colombian Forest Sector Model - XtCFSM (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.ChartTitle.Text = "Imports of Colombian Forest Sector Model at final iteration for Market Clearing Condition MCC - Mtw (2015 thousand million COP)"
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.ChartTitle.Text = "Imports of Colombian Forest Sector Model - MtCFSM (2015 thousand million COP)"

'copiar formato celdas
rowDataActu = hojUsu_Report_MCC.Cells(Rows.Count, "B").End(xlUp).Row

hojUsu_Report_MCC.Range(Cells(3, 2), Cells(3, 40)).Copy
hojUsu_Report_MCC.Range(Cells(3, 2), Cells(rowDataActu, 40)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

'datos de las gráficas
hojUsu_Report_MCC.ChartObjects("G_MCC_SolverFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AN$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_MCC").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$AK$2:$AN$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$Z$2:$Z$" & rowDataActu & _
    ",'Report MCC'!$AJ$2:$AJ$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesHistoricalData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$W$2:$W$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_SetPricesEstimatedData").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$X$2:$X$" & rowDataActu & _
    ",'Report MCC'!$AI$2:$AI$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_D_S_FinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_D_S").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)
'
hojUsu_Report_MCC.ChartObjects("G_SupplyFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$F$2:$F$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Supply").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$C$2:$F$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_DemandFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$J$2:$J$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Demand").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$G$2:$J$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ConsumptionFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$N$2:$N$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Consumption").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$K$2:$N$" & rowDataActu)

hojUsu_Report_MCC.ChartObjects("G_ExportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$R$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Exports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$O$2:$R$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_ImportsFinalIteration").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$V$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.ChartObjects("G_Imports").Activate
    ActiveChart.SetSourceData Source:=Range("'Report MCC'!$B$2:$B$" & rowDataActu & _
    ",'Report MCC'!$S$2:$V$" & rowDataActu)
    
hojUsu_Report_MCC.Cells(1, 3).Select
Application.CutCopyMode = False

'Borrar contenido de columnas
    Columns("W:W").Select
    Selection.Clear
    
    Columns("AA:AA").Select
    Selection.Clear
    
    Columns("AE:AE").Select
    Selection.Clear
    
    Columns("AI:AI").Select
    Selection.Clear

'esconder columnas sin datos
    Columns("W:AJ").EntireColumn.Hidden = True
    Columns("AZ:BI").EntireColumn.Hidden = True

Application.ScreenUpdating = True

End Sub
