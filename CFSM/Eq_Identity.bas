Attribute VB_Name = "Eq_Identity"
Sub ALL_WOOD_INDUSTRY()

Call SUPPLY_WOOD_INDUSTRY

Call PRICE_OF_CONSUMPTION_WOOD_INDUSTRY
Call PRICE_OF_EXPORTS_WOOD_INDUSTRY
Call PRICE_OF_IMPORT_WOOD_INDUSTRY

Call CONSUMPTION_WOOD_INDUSTRY
Call EXPORTS_WOOD_INDUSTRY
Call IMPORTS_WOOD_INDUSTRY

End Sub
Sub ALL_FURNITURE_INDUSTRY()

Call SUPPLY_FURNITURE_INDUSTRY
Call PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY
Call PRICE_OF_EXPORTS_FURNITURE_INDUSTRY
Call PRICE_OF_IMPORT_FURNITURE_INDUSTRY
Call CONSUMPTION_FURNITURE_INDUSTRY
Call EXPORTS_FURNITURE_INDUSTRY
Call IMPORTS_FURNITURE_INDUSTRY

End Sub
Sub ALL_PULP_PAPER_INDUSTRY()

Call SUPPLY_PULP_PAPER_INDUSTRY
Call PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY
Call PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY
Call PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY
Call CONSUMPTION_PULP_PAPER_INDUSTRY
Call EXPORTS_PULP_PAPER_INDUSTRY
Call IMPORTS_PULP_PAPER_INDUSTRY

End Sub
Sub ALL_WOOD_INDUSTRIAL()

Call SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
Call SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST
Call PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL
Call PRICE_OF_EXPORTS_WOOD_INDUSTRIAL
Call PRICE_OF_IMPORT_WOOD_INDUSTRIAL
Call CONSUMPTION_WOOD_INDUSTRIAL
Call EXPORTS_WOOD_INDUSTRIAL
Call IMPORTS_WOOD_INDUSTRIAL

End Sub
Sub ALL_FIREWOOD()

Call SUPPLY_FIREWOOD
Call PRICE_OF_CONSUMPTION_FIREWOOD
Call PRICE_OF_EXPORTS_FIREWOOD
Call PRICE_OF_IMPORT_FIREWOOD
Call CONSUMPTION_FIREWOOD
Call EXPORTS_FIREWOOD
Call IMPORTS_FIREWOOD

End Sub
Sub ALL_FINAL_RURAL_CONSUMPTION()



End Sub
Sub CALL_EQUATIONS()

'Permitir el año 1970 para visualizar los datos historicos, de lo contrario colocar 1975
If hojUsu_SystemOptions.Range("SelectProcess") <> 3 Then
    hojUsu_SystemOptions.Range("InitialYearRange") = 1975
End If

Dim market, equation As String

market = hojUsu_SystemOptions.Range("MarketsInputs").Value
equation = hojUsu_SystemOptions.Range("EquationsInputs").Value

Select Case market

    Case "Wood_Industry"
    
        Select Case equation
        
            Case "Supply"
            
            Call SUPPLY_WOOD_INDUSTRY
            
            Case "Consumption"
            
            Call CONSUMPTION_WOOD_INDUSTRY
            
            Case "Exports"
            
            Call EXPORTS_WOOD_INDUSTRY
            
            Case "Imports"
            
            Call IMPORTS_WOOD_INDUSTRY
            
            Case "Price deflator of consumption"
            
            Call PRICE_OF_CONSUMPTION_WOOD_INDUSTRY
            
            Case "Price deflator of exports"
            
            Call PRICE_OF_EXPORTS_WOOD_INDUSTRY
            
            Case "Price deflator of imports"
            
            Call PRICE_OF_IMPORT_WOOD_INDUSTRY
            
            Case "All"
            
            Call ALL_WOOD_INDUSTRY
        
        End Select

    Case "Furniture_Industry"
    
        Select Case equation
        
            Case "Supply"
            
            Call SUPPLY_FURNITURE_INDUSTRY
            
            Case "Consumption"
            
            Call CONSUMPTION_FURNITURE_INDUSTRY
            
            Case "Exports"
            
            Call EXPORTS_FURNITURE_INDUSTRY
            
            Case "Imports"
            
            Call IMPORTS_FURNITURE_INDUSTRY
            
            Case "Price deflator of consumption"
            
            Call PRICE_OF_CONSUMPTION_FURNITURE_INDUSTRY
            
            Case "Price deflator of exports"
            
            Call PRICE_OF_EXPORTS_FURNITURE_INDUSTRY
            
            Case "Price deflator of imports"
            
            Call PRICE_OF_IMPORT_FURNITURE_INDUSTRY
            
            Case "All"
            
            Call ALL_FURNITURE_INDUSTRY
        
        End Select

    Case "Pulp_Paper_Industry"
    
        Select Case equation
        
            Case "Supply"
            
            Call SUPPLY_PULP_PAPER_INDUSTRY
            
            Case "Consumption"
            
            Call CONSUMPTION_PULP_PAPER_INDUSTRY
            
            Case "Exports"
            
            Call EXPORTS_PULP_PAPER_INDUSTRY
            
            Case "Imports"
            
            Call IMPORTS_PULP_PAPER_INDUSTRY
            
            Case "Price deflator of consumption"
            
            Call PRICE_OF_CONSUMPTION_PULP_PAPER_INDUSTRY
            
            Case "Price deflator of exports"
            
            Call PRICE_OF_EXPORTS_PULP_PAPER_INDUSTRY
            
            Case "Price deflator of imports"
            
            Call PRICE_OF_IMPORT_PULP_PAPER_INDUSTRY
            
            Case "All"
            
            Call ALL_PULP_PAPER_INDUSTRY
        
        End Select

    Case "Wood_Industrial"
    
        Select Case equation
        
            Case "Supply forest plantations"
            
            Call SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
            Call SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST
            
            Case "Supply natural forest"
            
            Call SUPPLY_WOOD_INDUSTRIAL_FOREST_PLANTATIONS
            Call SUPPLY_WOOD_INDUSTRIAL_NATURAL_FOREST
            
            Case "Consumption"
            
            Call CONSUMPTION_WOOD_INDUSTRIAL
            
            Case "Exports"
            
            Call EXPORTS_WOOD_INDUSTRIAL
            
            Case "Imports"
            
            Call IMPORTS_WOOD_INDUSTRIAL
            
            Case "Price deflator of consumption"
            
            Call PRICE_OF_CONSUMPTION_WOOD_INDUSTRIAL
            
            Case "Price deflator of exports"
            
            Call PRICE_OF_EXPORTS_WOOD_INDUSTRIAL
            
            Case "Price deflator of imports"
            
            Call PRICE_OF_IMPORT_WOOD_INDUSTRIAL
            
            Case "All"
            
            Call ALL_WOOD_INDUSTRIAL
        
        End Select

    Case "Firewood"
    
        Select Case equation
        
            Case "Supply"
            
            Call SUPPLY_FIREWOOD
            
            Case "Consumption"
            
            Call CONSUMPTION_FIREWOOD
            
            Case "Exports"
            
            Call EXPORTS_FIREWOOD
            
            Case "Imports"
            
            Call IMPORTS_FIREWOOD
            
            Case "Price deflator of consumption"
            
            Call PRICE_OF_CONSUMPTION_FIREWOOD
            
            Case "Price deflator of exports"
            
            Call PRICE_OF_EXPORTS_FIREWOOD
            
            Case "Price deflator of imports"
            
            Call PRICE_OF_IMPORT_FIREWOOD
            
            Case "All"
            
            Call ALL_FIREWOOD
        
        End Select

    Case "All"
    
        Select Case equation
        
            Case "All"
            
            Call ALL_WOOD_INDUSTRY
            Call ALL_FURNITURE_INDUSTRY
            Call ALL_PULP_PAPER_INDUSTRY
            Call ALL_WOOD_INDUSTRIAL
            Call ALL_FIREWOOD
                   
        End Select

End Select

End Sub
