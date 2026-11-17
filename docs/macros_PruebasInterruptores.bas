' Código VBA del libro "GDL - Pruebas Interruptores" (.xlsm). El libro en data/raw se guardó sin macros para que
' el antivirus no bloquee el proyecto; este es el código original, solo como referencia.

' ===== Sheet100.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Module1.bas =====
Sub NuevaEntrada()
    
    ' Generacion de variables de uso interno
    
    Pagina = Sheets("Indice").Range("C1").Value + 16
    PaginaHiper = "'" & Pagina - 16 & "'!A1"
    celda = "C" & Pagina
    CeldaQ = "Q" & Pagina
    CeldaR = "R" & Pagina
    CeldaS = "S" & Pagina
    CeldaT = "T" & Pagina
    CeldaU = "U" & Pagina
    CeldaV = "V" & Pagina
    CeldaW = "W" & Pagina
    CeldaX = "X" & Pagina
    CeldaY = "Y" & Pagina
    CeldaZ = "Z" & Pagina
    CeldaAA = "AA" & Pagina
    CeldaAB = "AB" & Pagina
    CeldaAC = "AC" & Pagina
    CeldaAD = "AD" & Pagina
    CeldaAE = "AE" & Pagina
    CeldaAF = "AF" & Pagina
    CeldaAG = "AG" & Pagina
    CeldaAH = "AH" & Pagina
    CeldaAI = "AI" & Pagina
    CeldaAJ = "AJ" & Pagina
    CeldaAK = "AK" & Pagina
    CeldaAL = "AL" & Pagina
    CeldaAM = "AM" & Pagina
    
    ' Genera una copia de la hoja Blank y la coloca al final con el nombre del consecutivo
    
    Sheets("Blank").Copy After:=Sheets((ActiveWorkbook.Sheets.Count) - 1)
    ActiveSheet.Name = Sheets("Indice").Range("C1").Value

    ' Cambia las formulas de los encabezados

    Columns("G:J").Select
    Selection.Replace What:="16", Replacement:=Pagina
    ', LookAt:=xlPart, _
    '    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    '    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    Range("A1").Select
    
    ' Genera los hypervinculos a la nueva hoja
    
    Sheets("Indice").Select
    Range([celda]).Select
    Range([celda]).Value = Pagina - 16

    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        PaginaHiper

    Range("A1").Select
    
    
    ' Genera las formulas de la nueva entrada
    
    ' Prueba Condición física
        
    Range([CeldaT]).Select
    Range([CeldaT]).Value = "=MAX(" & Pagina - 16 & "!U16:U32)"

    Range([CeldaU]).Select
    Range([CeldaU]).Value = "=IF(ISERROR(VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,16)),"""",VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,16))"
' =IF(ISERROR(VLOOKUP(S16,EJEMPLO!U16:AM32,16)),"",VLOOKUP(S16,EJEMPLO!U16:AM32,16))

    Range([CeldaV]).Select
    Range([CeldaV]).Value = "=IF(ISERROR(VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,17)),"""",VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,17))"
'=IF(ISERROR(VLOOKUP(S16,EJEMPLO!U16:AM32,17)),"",VLOOKUP(S16,EJEMPLO!U16:AM32,17))

    Range([CeldaW]).Select
    Range([CeldaW]).Value = "=IF(ISERROR(VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,18)),"""",VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,18))"

    Range([CeldaX]).Select
    Range([CeldaX]).Value = "=IF(ISERROR(VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,19)),"""",VLOOKUP(T" & Pagina & "," & Pagina - 16 & "!U16:AM32,19))"


    ' Prueba Operacion
    
    Range([CeldaY]).Select
    Range([CeldaY]).Value = "=MAX(" & Pagina - 16 & "!AU16:AU32)"
    
    Range([CeldaZ]).Select
    Range([CeldaZ]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,15)),"""",VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,15))"
    
    Range([CeldaAA]).Select
    Range([CeldaAA]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,16)),"""",VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,16))"
    
    Range([CeldaAB]).Select
    Range([CeldaAB]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,17)),"""",VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,17))"
    
    Range([CeldaAC]).Select
    Range([CeldaAC]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,18)),"""",VLOOKUP(Y" & Pagina & "," & Pagina - 16 & "!AU16:BL32,18))"
    
    
    ' Prueba Electrica
    
    Range([CeldaAD]).Select
    Range([CeldaAD]).Value = "=MAX(" & Pagina - 16 & "!BT16:BT32)"
    
    Range([CeldaAE]).Select
    Range([CeldaAE]).Value = "=IF(ISERROR(VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,23)),"""",VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,23))"
    
    Range([CeldaAF]).Select
    Range([CeldaAF]).Value = "=IF(ISERROR(VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,24)),"""",VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,24))"
    
    Range([CeldaAG]).Select
    Range([CeldaAG]).Value = "=IF(ISERROR(VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,25)),"""",VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,25))"
    
    Range([CeldaAH]).Select
    Range([CeldaAH]).Value = "=IF(ISERROR(VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,26)),"""",VLOOKUP(AD" & Pagina & "," & Pagina - 16 & "!BT16:CS32,26))"


    ' Prueba Relevador
    
    Range([CeldaAI]).Select
    Range([CeldaAI]).Value = "=MAX(" & Pagina - 16 & "!DA16:DA32)"
    
    Range([CeldaAJ]).Select
    Range([CeldaAJ]).Value = "=IF(ISERROR(VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,13)),"""",VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,13))"

    Range([CeldaAK]).Select
    Range([CeldaAK]).Value = "=IF(ISERROR(VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,14)),"""",VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,14))"

    Range([CeldaAL]).Select
    Range([CeldaAL]).Value = "=IF(ISERROR(VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,15)),"""",VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,15))"

    Range([CeldaAM]).Select
    Range([CeldaAM]).Value = "=IF(ISERROR(VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,16)),"""",VLOOKUP(AI" & Pagina & "," & Pagina - 16 & "!DA16:DP32,16))"


    Range("A1").Select


'    Range([CeldaQ]).Select
'    Range([CeldaQ]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",MAX('" & Pagina - 16 & "'!$U$16:$U$32))"
     
'    Range([CeldaR]).Select
'    Range([CeldaR]).Value = "=IF(ISERROR(VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,16)),"""",VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,16))"

'    Range([CeldaS]).Select
'    Range([CeldaS]).Value = "=IF(ISERROR(VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,18)),"""",VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,18))"

'    Range([CeldaT]).Select
'    Range([CeldaT]).Value = "=IF(ISERROR(VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,19)),"""",VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,19))"


    ' Prueba Aceite Cromatografia
    
'    Range([CeldaU]).Select
'    Range([CeldaU]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",MAX('" & Pagina - 16 & "'!$AU$16:$AU$32))"
     
'    Range([CeldaV]).Select
'    Range([CeldaV]).Value = "=IF(ISERROR(VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,15)),"""",VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,15))"

'    Range([CeldaW]).Select
'    Range([CeldaW]).Value = "=IF(ISERROR(VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,16)),"""",VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,16))"

'    Range([CeldaX]).Select
'    Range([CeldaX]).Value = "=IF(ISERROR(VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,18)),"""",VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,18))"


    ' Prueba Electricas
    
'    Range([CeldaY]).Select
'    Range([CeldaY]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",MAX('" & Pagina - 16 & "'!$BT$16:$BT$32))"
     
'    Range([CeldaZ]).Select
'    Range([CeldaZ]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,22)),"""",VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,22))"

'    Range([CeldaAA]).Select
'    Range([CeldaAA]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,23)),"""",VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,23))"

'    Range([CeldaAB]).Select
'    Range([CeldaAB]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,25)),"""",VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,25))"


'    Range([celda]).Select
    
    
End Sub

Function ColorFondoIgual(celda As Range, Rojo As Integer, Verde As Integer, Azul As Integer) As Boolean

If celda(1, 1).Interior.Color = RGB(Rojo, Verde, Azul) Then
    ColorFondoIgual = True
Else
    ColorFondoIgual = False
End If

End Function
Function obtenercolor2(celda As Range) As String

Dim C As Long
Dim R As Long
Dim G As Long
Dim B As Long
C = celda.Interior.Color
R = C Mod 256
G = C \ 256 Mod 256
B = C \ 65536 Mod 256
obtenercolor2 = "R=" & R & ", G=" & G & ", B=" & B

End Function

' ===== Sheet60.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet61.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet62.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet63.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet64.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet65.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet66.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet67.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet68.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet69.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet70.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet71.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet72.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet73.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet74.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet75.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet76.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet77.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet78.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet79.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet80.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet81.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet82.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet83.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet84.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet85.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet86.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet87.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet88.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet89.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet90.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet91.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet92.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet93.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet94.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet95.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet96.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet97.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet98.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet99.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet102.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet103.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet104.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet105.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet106.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet107.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet108.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet109.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet110.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet111.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet112.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet113.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet114.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet115.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet116.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet117.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet118.cls =====
Private Sub CommandButton1_Click()

End Sub

' ===== Sheet119.cls =====
Private Sub CommandButton1_Click()

End Sub
