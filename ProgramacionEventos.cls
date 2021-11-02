VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)



'        If Target.Column = 1 Then
'        ThisRow = Target.Row
'            If Target.Value > 100 Then
'                Range("B" & ThisRow).Interior.ColorIndex = 3
'                Rows("7:13").EntireRow.Hidden = True
'            Else
'                Range("B" & ThisRow).Interior.ColorIndex = xlColorIndexNone
'                Rows("6:14").EntireRow.Hidden = False
'            End If
'        End If

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
