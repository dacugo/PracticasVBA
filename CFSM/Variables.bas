Attribute VB_Name = "Variables"
Option Explicit

Public initialYear As Integer
Public finalYear As Integer
Public c_ini As Integer
Public c_fin As Integer
Public d_ini As Integer
Public d_fin As Integer
Public selectProcess As Integer
Public negativeData As Integer
Public k As Integer
Public i As Integer
Public j As Integer
Public l As Integer
Public m As Integer

Sub AsignacionVariablesOpcionesUsuario()
    
'año inicial asignado por el usuario
    initialYear = hojUsu_SystemOptions.Range("InitialYearRange")

'año final asignado por el usuario
    finalYear = hojUsu_SystemOptions.Range("FinalYearRange")

'Rangos de años utilizados en los procesos y los resultados de cada ecuación
    c_ini = initialYear - 1967
    d_ini = initialYear - 1936
    c_fin = finalYear - 1967
    d_fin = finalYear - 1936

'Tipo de proceso

'seleccionar si se desea ver la validación del sistema o el Market Clearing Condition
    selectProcess = hojUsu_SystemOptions.Range("SelectProcess").Value
'uso de los datos que resultan negativos (usar datos original o como lo genera la ecuación
    negativeData = hojUsu_SystemOptions.Range("NegativeData").Value

End Sub
Sub AsignacionVariablesProcesos()

    'k es el año actual de la hoja summary

        'i es el número de la celda en la hoja del mercado
        i = k - 31
        
        'j es el número de la celda en la hoja del mercado
        j = i - 1
        
        'l es el n anterior de la hoja en summary
        l = k - 1
        
        'm es el n dos años anteriores al año actual
        m = l - 1
        
End Sub
