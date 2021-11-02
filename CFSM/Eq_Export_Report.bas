Attribute VB_Name = "Eq_Export_Report"
Sub EXPORT_REPORT()
Attribute EXPORT_REPORT.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim results As Object
Dim rutaArchivo, nombreArchivo, nombreLibroAplicativo As String

'ruta del aplicativo y nombre del libro donde se exportan los resultados
nombreLibroAplicativo = ActiveWorkbook.Name
'nombreLibroAplicativo = Left(nombreLibroAplicativo, Len(nombreLibroAplicativo) - 5)
rutaArchivo = ActiveWorkbook.Path
nombreArchivo = "Results.xlsx"
rutaNuevoArchivo = rutaArchivo & "\" & nombreArchivo

'método seleccionado
metodo = hojUsu_SystemOptions.Range("ReportType")

Select Case metodo
    Case "New"
        'si el libro Results esta cerrado, se crea el documento en la ruta actual del archivo y se le anexa la primera hoja
        Set results = Workbooks.Add()
    
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs Filename:=rutaArchivo & "\" & nombreArchivo
        
        Workbooks(nombreLibroAplicativo).Activate
        Sheets("Report").Select
        
        Sheets("Report").Copy Before:=Workbooks("Results.xlsx").Sheets(1)
    
        Sheets("Hoja1").Delete
        Application.DisplayAlerts = True
    Case "Continue"
        'si el libro Results esta abierto, se copia la hoja directamente
        Workbooks(nombreLibroAplicativo).Activate
        Sheets("Report").Select
        Sheets("Report").Copy Before:=Workbooks("Results.xlsx").Sheets(1)
End Select

hojUsu_SystemOptions.Activate

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True

End Sub
Sub EXPORT_REPORT_MCC()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim results As Object
Dim rutaArchivo, nombreArchivo, nombreLibroAplicativo As String

'ruta del aplicativo y nombre del libro donde se exportan los resultados
nombreLibroAplicativo = ActiveWorkbook.Name
'nombreLibroAplicativo = Left(nombreLibroAplicativo, Len(nombreLibroAplicativo) - 5)
rutaArchivo = ActiveWorkbook.Path
nombreArchivo = "Results.xlsx"
rutaNuevoArchivo = rutaArchivo & "\" & nombreArchivo

'método seleccionado
metodo = hojUsu_SystemOptions.Range("ReportType")

Select Case metodo
    Case "New"
        'si el libro Results esta cerrado, se crea el documento en la ruta actual del archivo y se le anexa la primera hoja
        Set results = Workbooks.Add()
    
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs Filename:=rutaArchivo & "\" & nombreArchivo
        
        Workbooks(nombreLibroAplicativo).Activate
        Sheets("Report MCC").Select
        
        Sheets("Report MCC").Copy Before:=Workbooks("Results.xlsx").Sheets(1)
    
        Sheets("Hoja1").Delete
        Application.DisplayAlerts = True
    Case "Continue"
        'si el libro Results esta abierto, se copia la hoja directamente
        Workbooks(nombreLibroAplicativo).Activate
        Sheets("Report MCC").Select
        Sheets("Report MCC").Copy Before:=Workbooks("Results.xlsx").Sheets(1)
End Select

hojUsu_SystemOptions.Activate
Application.Calculation = xlCalculationAutomatic

Application.ScreenUpdating = True

End Sub
Sub EXPORT_EQUATION()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim market, equation As String

'configuración del sistema
market = hojUsu_SystemOptions.Range("MarketsOutputs").Value
equation = hojUsu_SystemOptions.Range("EquationsOutputs").Value

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
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Stw"

            Case "Consumption"

            Call REPORT_CONSUMPTION_WOOD_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Ctw"

            Case "Exports"

            Call REPORT_EXPORTS_WOOD_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Xtw"

            Case "Imports"

            Call REPORT_IMPORTS_WOOD_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Mtw"

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PCtw"

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PXtw"

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PMtw"

            Case "All"

            Call REPORT_ALL_WOOD_INDUSTRY

        End Select

    Case "Furniture_Industry"

        Select Case equation

            Case "Supply"

            Call REPORT_SUPPLY_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Stf"

            Case "Consumption"

            Call REPORT_CONSUMPTION_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Ctf"

            Case "Exports"

            Call REPORT_EXPORTS_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Xtf"

            Case "Imports"

            Call REPORT_IMPORTS_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Mtf"

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PCtf"

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PXtf"

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PMtf"

            Case "All"

            Call REPORT_ALL_FURNITURE_INDUSTRY

        End Select

    Case "Pulp_Paper_Industry"

        Select Case equation

            Case "Supply"

            Call REPORT_SUPPLY_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Stz"

            Case "Consumption"

            Call REPORT_CONSUMPTION_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Ctz"

            Case "Exports"

            Call REPORT_EXPORTS_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Xtz"

            Case "Imports"

            Call REPORT_IMPORTS_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "Mtz"

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PCtz"

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PXtz"

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PMtz"

            Case "All"

            Call REPORT_ALL_PULP_PAPER_INDUSTRY

        End Select

    Case "Wood_Industrial"

        Select Case equation

            Case "Supply forest plantations"

            Call REPORT_SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "StMWfprw"

            Case "Supply natural forest"

            Call REPORT_SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "StMWnfrw"

            Case "Consumption"

            Call REPORT_CONSUMPTION_WOOD_INDUSTRIAL
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "CtMWrw"

            Case "Exports"

            Call REPORT_EXPORTS_WOOD_INDUSTRIAL
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "XtMWrw"

            Case "Imports"

            Call REPORT_IMPORTS_WOOD_INDUSTRIAL
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "MtMWrw"

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PCtMWrw"

            Case "Price deflator of exports"

            Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRIAL
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PXtMWrw"

            Case "Price deflator of imports"

            Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRIAL
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PMtMWrw"

            Case "All"

            Call REPORT_ALL_WOOD_INDUSTRIAL

        End Select

    Case "Firewood"

        Select Case equation

            Case "Supply"

            Call REPORT_SUPPLY_FIREWOOD
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "StFWnfrw"

            Case "Consumption"

            Call REPORT_CONSUMPTION_FIREWOOD
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "CtFWrw"

            Case "Price deflator of consumption"

            Call REPORT_PRICE_OF_CONSUMPTION_FIREWOOD
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "PCtFWrw"

            Case "All"

            Call REPORT_ALL_FIREWOOD

        End Select

    Case "Set_prices"

        Select Case equation

            Case "SP_Wood_Industry"

            Call REPORT_SP_WOOD_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "SP_Wood_Industry"

            Case "SP_Furniture_Industry"

            Call REPORT_SP_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "SP_Furniture_Industry"

            Case "SP_Pulp_Paper_Industry"

            Call REPORT_SP_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "SP_Pulp_Paper_Industry"

            Case "SP_Wood_Industrial"

            Call REPORT_SP_WOOD_INDUSTRIAL
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "SP_Wood_Industrial"

            Case "SP_Firewood"

            Call REPORT_SP_FIREWOOD
            Call EXPORT_REPORT
            Workbooks("Results.xlsx").Sheets(1).Name = "SP_Firewood"

            Case "All"

            Call REPORT_SP_ALL

        End Select
    
    Case "MCC"
        
        Select Case equation

            Case "MCC_Wood_Industry"

            Call REPORT_MCC_WOOD_INDUSTRY
            Call EXPORT_REPORT
            hojUsu_SystemOptions.Range("ReportType") = "Continue"
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Wood_Industry"
            Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Wood_Industry_S"

            Case "MCC_Furniture_Industry"

            Call REPORT_MCC_FURNITURE_INDUSTRY
            Call EXPORT_REPORT
            hojUsu_SystemOptions.Range("ReportType") = "Continue"
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Furniture_Industry"
            Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Furniture_Industry_S"

            Case "MCC_Pulp_Paper_Industry"

            Call REPORT_MCC_PULP_PAPER_INDUSTRY
            Call EXPORT_REPORT
            hojUsu_SystemOptions.Range("ReportType") = "Continue"
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Pulp_Paper_Industry"
            Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Pulp_Paper_Industry_S"

            Case "MCC_Wood_Industrial"

            Call REPORT_MCC_INDUSTRIAL_WOOD
            Call EXPORT_REPORT
            hojUsu_SystemOptions.Range("ReportType") = "Continue"
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Wood_Industrial"
            Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Wood_Industrial_S"

            Case "MCC_Firewood"

            Call REPORT_MCC_FIREWOOD
            Call EXPORT_REPORT
            hojUsu_SystemOptions.Range("ReportType") = "Continue"
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Firewood"
            Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Firewood_S"
            
            Case "MCC_MWM"

            Call REPORT_MCC_MWM_EXTENDED
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_MWM"
            
            Case "MCC_UWM"

            Call REPORT_MCC_UWM_EXTENDED
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_UWM"
            
            Case "MCC_CFSM"

            Call REPORT_MCC_CFSM_EXTENDED
            Call EXPORT_REPORT_MCC
            Workbooks("Results.xlsx").Sheets(1).Name = "MCC_CFSM"

            Case "All"

            Call REPORT_MCC_ALL

        End Select

    Case "All"

        Select Case equation

            Case "All"

            Call REPORT_ALL_REPORT

        End Select

End Select

hojUsu_SystemOptions.Activate

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
Sub REPORT_ALL_WOOD_INDUSTRY()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_SUPPLY_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Stw"

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call REPORT_CONSUMPTION_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Ctw"

Call REPORT_EXPORTS_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Xtw"

Call REPORT_IMPORTS_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Mtw"

Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtw"

Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtw"

Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtw"

Application.ScreenUpdating = True


End Sub
Sub REPORT_ALL_FURNITURE_INDUSTRY()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_SUPPLY_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Stf"

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call REPORT_CONSUMPTION_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Ctf"

Call REPORT_EXPORTS_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Xtf"

Call REPORT_IMPORTS_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Mtf"

Call REPORT_PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtf"

Call REPORT_PRICE_OF_EXPORTS_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtf"

Call REPORT_PRICE_OF_IMPORT_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtf"

Application.ScreenUpdating = True

End Sub
Sub REPORT_ALL_PULP_PAPER_INDUSTRY()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_SUPPLY_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Stz"

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call REPORT_CONSUMPTION_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Ctz"

Call REPORT_EXPORTS_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Xtz"

Call REPORT_IMPORTS_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Mtz"

Call REPORT_PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtz"

Call REPORT_PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtz"

Call REPORT_PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtz"

Application.ScreenUpdating = True

End Sub
Sub REPORT_ALL_WOOD_INDUSTRIAL()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "StMWfprw"

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call REPORT_SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "StMWnfrw"

Call REPORT_CONSUMPTION_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "CtMWrw"

Call REPORT_EXPORTS_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "XtMWrw"

Call REPORT_IMPORTS_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "MtMWrw"

Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtMWrw"

Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtMWrw"

Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtMWrw"

Application.ScreenUpdating = True

End Sub
Sub REPORT_ALL_FIREWOOD()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_SUPPLY_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "StFWnfrw"

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call REPORT_CONSUMPTION_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "CtFWrw"

Call REPORT_PRICE_OF_CONSUMPTION_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtFWrw"

Application.ScreenUpdating = True

End Sub
Sub REPORT_SP_ALL()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_SP_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Wood_Industry"

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call REPORT_SP_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Furniture_Industry"

Call REPORT_SP_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Pulp_Paper_Industry"

Call REPORT_SP_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Wood_Industrial"

Call REPORT_SP_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Firewood"

Application.ScreenUpdating = True

End Sub
Sub REPORT_MCC_ALL()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_MCC_WOOD_INDUSTRY
Call EXPORT_REPORT

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Wood_Industry"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Wood_Industry_S"


Call REPORT_MCC_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Furniture_Industry"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Furniture_Industry_S"

Call REPORT_MCC_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Pulp_Paper_Industry"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Pulp_Paper_Industry_S"

Call REPORT_MCC_INDUSTRIAL_WOOD
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Wood_Industrial"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Wood_Industrial_S"

Call REPORT_MCC_FIREWOOD
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Firewood"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Firewood_S"

Call REPORT_MCC_MWM_EXTENDED
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_MWM"

Call REPORT_MCC_UWM_EXTENDED
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_UWM"

Call REPORT_MCC_CFSM_EXTENDED
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_CFSM"

Application.ScreenUpdating = True

End Sub
Sub REPORT_ALL_REPORT()

Application.ScreenUpdating = False

hojUsu_SystemOptions.Range("ReportType") = "New"

Call REPORT_SUPPLY_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Stw"

hojUsu_SystemOptions.Range("ReportType") = "Continue"

Call REPORT_CONSUMPTION_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Ctw"

Call REPORT_EXPORTS_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Xtw"

Call REPORT_IMPORTS_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Mtw"

Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtw"

Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtw"

Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtw"

Call REPORT_SUPPLY_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Stf"

Call REPORT_CONSUMPTION_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Ctf"

Call REPORT_EXPORTS_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Xtf"

Call REPORT_IMPORTS_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Mtf"

Call REPORT_PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtf"

Call REPORT_PRICE_OF_EXPORTS_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtf"

Call REPORT_PRICE_OF_IMPORT_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtf"

Call REPORT_SUPPLY_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Stz"

Call REPORT_CONSUMPTION_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Ctz"

Call REPORT_EXPORTS_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Xtz"

Call REPORT_IMPORTS_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "Mtz"

Call REPORT_PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtz"

Call REPORT_PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtz"

Call REPORT_PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtz"

Call REPORT_SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "StMWfprw"

Call REPORT_SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "StMWnfrw"

Call REPORT_CONSUMPTION_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "CtMWrw"

Call REPORT_EXPORTS_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "XtMWrw"

Call REPORT_IMPORTS_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "MtMWrw"

Call REPORT_PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtMWrw"

Call REPORT_PRICE_OF_EXPORTS_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PXtMWrw"

Call REPORT_PRICE_OF_IMPORT_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PMtMWrw"

Call REPORT_SUPPLY_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "StFWnfrw"

Call REPORT_CONSUMPTION_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "CtFWrw"

Call REPORT_PRICE_OF_CONSUMPTION_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "PCtFWrw"

Call REPORT_SP_WOOD_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Wood_Industry"

Call REPORT_SP_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Furniture_Industry"

Call REPORT_SP_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Pulp_Paper_Industry"

Call REPORT_SP_WOOD_INDUSTRIAL
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Wood_Industrial"

Call REPORT_SP_FIREWOOD
Call EXPORT_REPORT
Workbooks("Results.xlsx").Sheets(1).Name = "SP_Firewood"

Call REPORT_MCC_WOOD_INDUSTRY
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Wood_Industry"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Wood_Industry_S"

Call REPORT_MCC_FURNITURE_INDUSTRY
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Furniture_Industry"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Furniture_Industry_S"

Call REPORT_MCC_PULP_PAPER_INDUSTRY
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Pulp_Paper_Industry"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Pulp_Paper_Industry_S"

Call REPORT_MCC_INDUSTRIAL_WOOD
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Wood_Industrial"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Wood_Industrial_S"

Call REPORT_MCC_FIREWOOD
Call EXPORT_REPORT
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_Firewood"
Workbooks("Results.xlsx").Sheets(2).Name = "MCC_Firewood_S"

Call REPORT_MCC_MWM_EXTENDED
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_MWM"

Call REPORT_MCC_UWM_EXTENDED
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_UWM"

Call REPORT_MCC_CFSM_EXTENDED
Call EXPORT_REPORT_MCC
Workbooks("Results.xlsx").Sheets(1).Name = "MCC_CFSM"

Application.ScreenUpdating = True

End Sub
