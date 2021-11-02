Attribute VB_Name = "Test"
Sub TestHistoricalData()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba para re
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2018

'aplica el reset a todo el sistema
Call RESET

'asigna los años de la prueba para el reporte
hojUsu_SystemOptions.Range("InitialYearRange") = 1970
hojUsu_SystemOptions.Range("FinalYearRange") = 2018

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'se configura para realizar la validation
hojUsu_SystemOptions.Range("SelectProcess") = 1

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'se ejecuta el sistema
Call CALL_EQUATIONS

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF1_IM1()

'empieza aquí

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 1
hojUsu_SystemOptions.Range("IterationMethod") = 1

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF2_IM1()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 2
hojUsu_SystemOptions.Range("IterationMethod") = 1

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF3_IM1()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 3
hojUsu_SystemOptions.Range("IterationMethod") = 1

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF1_IM2()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 1
hojUsu_SystemOptions.Range("IterationMethod") = 2

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF2_IM2()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 2
hojUsu_SystemOptions.Range("IterationMethod") = 2

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF3_IM2()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 3
hojUsu_SystemOptions.Range("IterationMethod") = 2

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF1_IM3()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 1
hojUsu_SystemOptions.Range("IterationMethod") = 3

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF2_IM3()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 2
hojUsu_SystemOptions.Range("IterationMethod") = 3

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV1_OF3_IM3()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 3
hojUsu_SystemOptions.Range("IterationMethod") = 3

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF1_IM1()

'empieza aquí

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 1
hojUsu_SystemOptions.Range("IterationMethod") = 1

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF2_IM1()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 2
hojUsu_SystemOptions.Range("IterationMethod") = 1

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF3_IM1()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 3
hojUsu_SystemOptions.Range("IterationMethod") = 1

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF1_IM2()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 1
hojUsu_SystemOptions.Range("IterationMethod") = 2

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF2_IM2()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 2
hojUsu_SystemOptions.Range("IterationMethod") = 2

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF3_IM2()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 3
hojUsu_SystemOptions.Range("IterationMethod") = 2

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF1_IM3()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 1
hojUsu_SystemOptions.Range("IterationMethod") = 3

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF2_IM3()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 2
hojUsu_SystemOptions.Range("IterationMethod") = 3

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeRawData_EV2_OF3_IM3()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 3

'configuración solver
hojUsu_SystemOptions.Range("VariablesSolver") = 1
hojUsu_SystemOptions.Range("OriginForVariablesTwo") = 3
hojUsu_SystemOptions.Range("IterationMethod") = 3

'se ejecuta el sistema
Call RUN_SOLVER

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeHistoricalData()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'se configura para realizar la validation
hojUsu_SystemOptions.Range("SelectProcess") = 2

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 1

'se ejecuta el sistema
Call CALL_EQUATIONS

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
Sub TestValidationData_NegativeDataZero()

'Selecciona todos los mercados
hojUsu_SystemOptions.Range("MarketsInputs") = "All"
hojUsu_SystemOptions.Range("MarketsOutputs") = "All"

'selecciona todas las ecuaciones
hojUsu_SystemOptions.Range("EquationsInputs") = "All"
hojUsu_SystemOptions.Range("EquationsOutputs") = "All"

'asigna los años de la prueba
hojUsu_SystemOptions.Range("InitialYearRange") = 1975
hojUsu_SystemOptions.Range("FinalYearRange") = 2015

'aplica el reset a todo el sistema
Call RESET

'se configura para realizar la validation
hojUsu_SystemOptions.Range("SelectProcess") = 2

'Negative data
hojUsu_SystemOptions.Range("NegativeData") = 2

'se ejecuta el sistema
Call CALL_EQUATIONS

'se exporta el reporte
Call EXPORT_EQUATION

End Sub
