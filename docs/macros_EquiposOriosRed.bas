Sub Borrar()
Attribute Borrar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Borrar Macro
'
   
End Sub
Sub Actualizar()
Attribute Actualizar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Actualizar Macro
'

'
    Range("B4").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("A2:N2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Range("EquiposOrios[#All]").RemoveDuplicates Columns:=6, Header _
        :=xlYes
    Range("A2").Select
    Sheets("P2").Select
    Range("A8").Select
    ActiveSheet.PivotTables("PivotTable3").PivotCache.Refresh
    Range("C2").Select
End Sub
