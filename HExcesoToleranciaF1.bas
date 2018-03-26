Attribute VB_Name = "HExcesoToleranciaF1"
Sub Excesos_Tolerancia_F1()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Call Obtengo_Datos
    Call CopiaDatos
    Sheets("Incidencias").Select
End Sub
Private Sub Obtengo_Datos()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Dotacion Ofisis").Select
    Worksheets("Dotacion Ofisis").Copy After:=Sheets(Sheets.Count)
    Sheets("Dotacion Ofisis (2)").Name = "Exc_Tol_1"
    Sheets("Exc_Tol_1").Select
    Cells.EntireColumn.Hidden = False
    'Todos losa datos a valores
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    'Obtengo fila y columna
    Dim Conteo, i As Integer
    Range("Q1").End(xlDown).Select
    Dim NroFila, NroColumna As Integer
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    Conteo = Range(Cells(2, 17), Cells(NroFila, NroColumna)).Count
    Range("AK:AW").ClearContents
    'Elimino las filas sin datos
    Range("AJ2").Select
    ActiveCell.Select
    For i = 1 To Conteo
        If ActiveCell.Value = "" Then
            Selection.EntireRow.Delete
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next i
    'Elimino columnas innecesarias
    Range("A:P,S:S,U:V,Z:AI").Columns.Delete
    Range("A1:L1").AutoFilter
    
    Range("A1").FormulaLocal = "CODIGO"
    Range("B1").FormulaLocal = "DNI"
    Range("C1").FormulaLocal = "NOMBRE"
    Range("D1").FormulaLocal = "TIPO"
    Range("E1").FormulaLocal = "TIENDA"
    Range("F1").FormulaLocal = "DPTO"
    Range("G1").FormulaLocal = "FECHA"
    Range("H1").FormulaLocal = "EVENTO"
    Range("I1").FormulaLocal = "Plan"
    Range("J1").FormulaLocal = "Real"
    Range("K1").FormulaLocal = "Dif"
    Range("L1").FormulaLocal = "OBS"
    
    Range("A2").Select
    'Declaracion de variables para validación
    Dim a, b, z As Integer
    a = 1
    b = 1
    z = 0
    'Validación de existencia de registros
    Do While a <= 3
        If Cells(b + 1, 1).Value <> "" Then
            z = z + 1
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    'Seleccion por cantidad de registros
    Select Case z
        Dim NroFila2, NroColumna2, Conteo2 As Integer
        Case 0
        Case 1
            NroFila2 = 1
            NroColumna2 = 1
            Range("E2").Formula = "=Incidencias!E11"
            Range("F2").Formula = "=IFERROR(VLOOKUP(B2,PareoMarcajes!B:F,5,0),"""")"
            Range("E2:F2").Copy
            Range(Cells(2, 5), Cells(NroFila2, NroColumna2 + 5)).PasteSpecial Paste:=xlPasteValues
            Range("H2").FormulaLocal = "Entrada"
            Range("L2").FormulaLocal = "Exc. Tol. Ingreso"
        Case Else
            Range("A1").End(xlDown).Select
            NroFila2 = ActiveCell.Row
            NroColumna2 = ActiveCell.Column
            Conteo2 = Range(Cells(2, 1), Cells(NroFila2, NroColumna2)).Count
            
            Application.Calculation = xlCalculationManual
            Range("E2").Formula = "=Incidencias!E11"
            Range("F2").Formula = "=IFERROR(VLOOKUP(B2,PareoMarcajes!B:F,5,0),"""")"
            Range("E2:F2").Copy
            Range(Cells(3, 5), Cells(NroFila2, NroColumna2 + 5)).PasteSpecial Paste:=xlPasteFormulas
            Application.Calculation = xlCalculationAutomatic
            Application.Calculate
            
            Range(Cells(2, 5), Cells(NroFila2, NroColumna2 + 5)).Copy
            Range(Cells(2, 5), Cells(NroFila2, NroColumna2 + 5)).PasteSpecial Paste:=xlPasteValues
            
            Range("H2").FormulaLocal = "Entrada"
            Range("L2").FormulaLocal = "Exc. Tol. Ingreso"
            
            Range("H2:L2").Copy
            Range(Cells(3, 8), Cells(NroFila2, NroColumna2 + 11)).PasteSpecial Paste:=xlPasteValues
    End Select
    Range("A2").Select
End Sub
Private Sub CopiaDatos()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Exc_Tol_1").Select
    Range("A2").Select
    'Declaración de variables de validación
    Dim a, b, z As Integer
    a = 1
    b = 1
    z = 0
    'Validación conteo de datos a copiar
    Do While a <= 3
        If Cells(b + 1, 1).Value <> "" Then
            z = z + 1
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    'Seleccion por cantidad de registros
    Select Case z
        'Declaración de variables
        Dim Nrofila3, Nrocolumna3 As Integer
        Case 0
        Case 1
            'Selecciono los datos a copiar
            Range(Cells(2, 1), Cells(2, 12)).Copy
            'Conteo de datos en incidencias
            Sheets("Incidencias").Select
            Range("A10").End(xlDown).Select
            Nrofila3 = ActiveCell.Row
            Nrocolumna3 = ActiveCell.Column
            'Pego los datos copiados
            Cells(Nrofila3 + 1, Nrocolumna3).Select
            Selection.PasteSpecial Paste:=xlPasteValues
            'Formato a registro copiado
            Range("A11:L11").Copy
            Cells(Nrofila3 + 1, Nrocolumna3).Select
            Selection.PasteSpecial Paste:=xlPasteFormats
        Case Else
            Dim cantDatos As Integer
            Range("A2").End(xlDown).Select
            Dim NroFila4, NroColumna4 As Integer
            NroFila4 = ActiveCell.Row
            NroColumna4 = ActiveCell.Column
            cantDatos = Range(Cells(2, 1), Cells(NroFila4, NroColumna4)).Count
            Range(Cells(2, 1), Cells(NroFila4, NroColumna4 + 11)).Copy
            'Conteo de datos en incidencias
            Sheets("Incidencias").Select
            Range("A10").End(xlDown).Select
            Nrofila3 = ActiveCell.Row
            Nrocolumna3 = ActiveCell.Column
            'Pego los datos copiados
            Cells(Nrofila3 + 1, Nrocolumna3).Select
            Selection.PasteSpecial Paste:=xlPasteValues
            'Formatos a registros copiados
            Dim c As Integer
            For c = 1 To cantDatos
                If c Mod 2 = 0 Then
                    Range("A11:L11").Copy
                    Cells(Nrofila3 + c, Nrocolumna3).Select
                    Selection.PasteSpecial Paste:=xlPasteFormats
                Else
                    Range("A12:L12").Copy
                    Cells(Nrofila3 + c, Nrocolumna3).Select
                    Selection.PasteSpecial Paste:=xlPasteFormats
                End If
            Next c
    End Select
    'Borrar la hoja "Excesos"
    Application.DisplayAlerts = False
    Sheets("Exc_Tol_1").Select
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
End Sub
