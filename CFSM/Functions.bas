Attribute VB_Name = "Functions"
Sub YearsSimpleReport()

'Rango años evaluados
    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años para insertar en el informe
    c_ini = a - 1967
    c_end = b - 1967

'años
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 2), Cells(c_end, 2)).Copy
    
    hojUsu_Report.Activate
        Range("B3").Select
        ActiveSheet.Paste
    
'formato centrado
    Columns("B:B").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

End Sub
Sub YearsMCCReport()

'Rango años evaluados
    a = hojUsu_SystemOptions.Range("InitialYearRange")
    b = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos años para insertar en el informe
    c_ini = a - 1967
    c_end = b - 1967

'años
    hojUsu_Forecast.Activate
    hojUsu_Forecast.Range(Cells(c_ini, 2), Cells(c_end, 2)).Copy
    
    hojUsu_Report_MCC.Activate
        Range("B3").Select
        ActiveSheet.Paste
    
'formato centrado
    Columns("B:B").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

End Sub
