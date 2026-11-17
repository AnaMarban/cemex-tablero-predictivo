' Código VBA del libro "GDL - Pruebas Transformadores" (.xlsm). El libro en data/raw se guardó sin macros para que
' el antivirus no bloquee el proyecto; este es el código original, solo como referencia.

' ===== Sheet1.cls =====
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
    
    ' Prueba Aceite Fisico Quimicas
    
    Range([CeldaQ]).Select
    Range([CeldaQ]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",MAX('" & Pagina - 16 & "'!$U$16:$U$32))"
     
    Range([CeldaR]).Select
    Range([CeldaR]).Value = "=IF(ISERROR(VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,16)),"""",VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,16))"

    Range([CeldaS]).Select
    Range([CeldaS]).Value = "=IF(ISERROR(VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,18)),"""",VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,18))"

    Range([CeldaT]).Select
    Range([CeldaT]).Value = "=IF(ISERROR(VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,19)),"""",VLOOKUP(Q" & Pagina & ",'" & Pagina - 16 & "'!$U$16:$AM$32,19))"


    ' Prueba Aceite Cromatografia
    
    Range([CeldaU]).Select
    Range([CeldaU]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",MAX('" & Pagina - 16 & "'!$AU$16:$AU$32))"
     
    Range([CeldaV]).Select
    Range([CeldaV]).Value = "=IF(ISERROR(VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,15)),"""",VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,15))"

    Range([CeldaW]).Select
    Range([CeldaW]).Value = "=IF(ISERROR(VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,16)),"""",VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,16))"

    Range([CeldaX]).Select
    Range([CeldaX]).Value = "=IF(ISERROR(VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,18)),"""",VLOOKUP(U" & Pagina & ",'" & Pagina - 16 & "'!$AU$16:$BL$32,18))"


    ' Prueba Electricas
    
    Range([CeldaY]).Select
    Range([CeldaY]).Value = "=IF(OR(ISBLANK(D" & Pagina & "),ISBLANK(E" & Pagina & ")),"""",MAX('" & Pagina - 16 & "'!$BT$16:$BT$32))"
     
    Range([CeldaZ]).Select
    Range([CeldaZ]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,22)),"""",VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,22))"

    Range([CeldaAA]).Select
    Range([CeldaAA]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,23)),"""",VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,23))"

    Range([CeldaAB]).Select
    Range([CeldaAB]).Value = "=IF(ISERROR(VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,25)),"""",VLOOKUP(Y" & Pagina & ",'" & Pagina - 16 & "'!$BT$16:$CR$32,25))"


    Range([celda]).Select
    
    
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

' ===== Sheet59.cls =====
Private Sub CommandButton1_Click()

End Sub

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
