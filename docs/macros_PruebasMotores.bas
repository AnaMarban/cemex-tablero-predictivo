' Macros del libro 'GDL - Pruebas Motores - Prueba_v2.xlsm' (el archivo en data/raw se guardó sin macros).

' ===== Sheet1.cls
' (Cada hoja de motor tiene este mismo botón; se muestra uno)
Private Sub CommandButton1_Click()

End Sub

' ===== Module2.bas
Sub NuevaEntrada()
    
    ' Generacion de variables de uso interno
    
    Pagina = Sheets("Indice").Range("AI4").Value + 10
    PaginaHiper = "'" & Pagina - 10 & "'!A1"
    Celda = "C" & Pagina
    CeldaO = "O" & Pagina
    CeldaP = "P" & Pagina
    CeldaQ = "Q" & Pagina
    CeldaR = "R" & Pagina
    CeldaS = "S" & Pagina
    CeldaT = "T" & Pagina
    CeldaU = "U" & Pagina
    CeldaV = "V" & Pagina
    CeldaW = "W" & Pagina
    CeldaX = "X" & Pagina
    
    ' Genera una copia de la hoja Blank y la coloca al final con el nombre del consecutivo
    
    Sheets("Blank").Copy After:=Sheets((ActiveWorkbook.Sheets.Count) - 1)
    ActiveSheet.Name = Sheets("Indice").Range("AI4").Value

    ' Cambia las formulas de los encabezados

    Rows("1:12").Select
    Selection.Replace What:="20", Replacement:=Pagina
    ', LookAt:=xlPart, _
    '    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    '    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    Range("A1").Select
    
    ' Genera los hypervinculos a la nueva hoja
    
    Sheets("Indice").Select
    Range([Celda]).Select
    Range([Celda]).Value = Pagina - 10

    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        PaginaHiper

    Range("A1").Select
    
    ' Genera las formulas de la nueva entrada
    
    Range([CeldaO]).Select
    Range([CeldaO]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",MAX('" & Pagina - 10 & "'!$J$19:$J$35))"
    
    '=IF(OR(ISBLANK(D11),ISBLANK(E11)),"",MAX('1'!$J$19:$J$35))
    
    Range([CeldaP]).Select
    Range([CeldaP]).Value = "=IF(ISERROR(VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AH$35,20)),"""",VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AH$35,20))"

    '=IF(ISERROR(VLOOKUP(O11,'1'!$J$19:$AH$35,20)),"",VLOOKUP(O11,'1'!$J$19:$AH$35,20))

    Range([CeldaQ]).Select
    Range([CeldaQ]).Value = "=IF(ISERROR(VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AH$35,24)),"""",VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AH$35,24))"

    '=IF(ISERROR(VLOOKUP(O11,'1'!$J$19:$AH$35,24)),"",VLOOKUP(O11,'1'!$J$19:$AH$35,24))

    Range([CeldaR]).Select
    Range([CeldaR]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",IF(NOT(ISNUMBER(P" & Pagina & ")),""A"",IF(P" & Pagina & ">='" & Pagina - 10 & "'!$AC$13,""N"",IF(AND(P" & Pagina & "<'" & Pagina - 10 & "'!$AC$13,P" & Pagina & ">'" & Pagina - 10 & "'!$AC$14),""P"",""A""))))"
    
    '=IF(OR(ISBLANK(D11),ISBLANK(E11)),"",IF(P11>='1'!$AC$13,"N",IF(AND(P11<'1'!$AC$13,P11>'1'!$AC$14),"P","A")))
    
    Range([CeldaS]).Select
    Range([CeldaS]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",IF(F" & Pagina & "=""JA"","""",MAX('" & Pagina - 10 & "'!AT19:AT35)))"

    '=IF(OR(ISBLANK(D11),ISBLANK(E11)),"",IF(F11="JA","",MAX('1'!AT19:AT35)))

    Range([CeldaT]).Select
    Range([CeldaT]).Value = "=IF(ISERROR(IF(F" & Pagina & "=""JA"","""",VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BR$35,20))),"""",IF(F" & Pagina & "=""JA"","""",VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BR$35,20)))"

    '=IF(ISERROR(IF(F11="JA","",VLOOKUP(S11,'1'!$AT$19:$BR$35,20))),"",IF(F11="JA","",VLOOKUP(S11,'1'!$AT$19:$BR$35,20)))

    Range([CeldaU]).Select
    Range([CeldaU]).Value = "=IF(ISERROR(IF(F" & Pagina & "=""JA"","""",VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BR$35,24))),"""",IF(F" & Pagina & "=""JA"","""",VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BR$35,24)))"

    '=IF(ISERROR(IF(F11="JA","",VLOOKUP(S11,'1'!$AT$19:$BR$35,24))),"",IF(F11="JA","",VLOOKUP(S11,'1'!$AT$19:$BR$35,24)))

    Range([CeldaV]).Select
    Range([CeldaV]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & "),F" & Pagina & "=""JA""),"""",IF(NOT(ISNUMBER(P" & Pagina & ")),""A"",IF(F" & Pagina & "=""JA"","""",IF(T" & Pagina & ">='" & Pagina - 10 & "'!$BM$13,""N"",IF(AND(T" & Pagina & "<'" & Pagina - 10 & "'!$BM$13,T" & Pagina & ">'" & Pagina - 10 & "'!$BM$14),""P"",""A"")))))"

    '=IF(OR(ISBLANK(D11),ISBLANK(E11)),"",IF(F11="JA","",IF(T11>='1'!$BM$13,"N",IF(AND(T11<'1'!$BM$13,T11>'1'!$BM$14),"P","A"))))

    Range([CeldaW]).Select
    Range([CeldaW]).Value = "=CONCATENATE(IF(ISERROR(VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AJ$35,26)),"""",VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AJ$35,26)),"", "",IF(ISERROR(VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BT$35,26)),"""",VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BT$35,26)))"

    Range([CeldaX]).Select
    Range([CeldaX]).Value = "=CONCATENATE(IF(ISERROR(VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AJ$35,27)),"""",VLOOKUP(O" & Pagina & ",'" & Pagina - 10 & "'!$J$19:$AJ$35,27)),"", "",IF(ISERROR(VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BT$35,27)),"""",VLOOKUP(S" & Pagina & ",'" & Pagina - 10 & "'!$AT$19:$BT$35,27)))"

    Range("M2").Select
    
End Sub
Sub CambiaCeldas()
Attribute CambiaCeldas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CambiaCeldas Macro
'

    Celda = Sheets("Indice").Range("AI4").Value + 10
    Rows("1:12").Select
    Selection.Replace What:="21", Replacement:=Celda, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Sub Hipervinculo()
Attribute Hipervinculo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Hipervinculo Macro
'

'
    ActiveCell.FormulaR1C1 = "51"
    Range("C61").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'51'!A1"
End Sub
Sub CambioHoja()
Attribute CambioHoja.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CambioHoja Macro
'

'
    Sheets("51").Select
End Sub
