Attribute VB_Name = "DTextoYFormato"
Sub DNI_aTexto_Incidencias()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Incidencias").Select
    'Agrego zoon a la hoja
    ActiveWindow.Zoom = 90
    'Eliminación de otros datos
    Dim NroFila, NroColumna, Conteo As Integer
    Range("L11").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    'Selecciona todo el rango
    Conteo = Range(Cells(11, 12), Cells(NroFila, NroColumna)).Count
    Dim i As Integer
    'Elimino filas distintas de Ent. Atrasada y Ausencia
    Range("L11").Select
    ActiveCell.Select
    For i = 1 To Conteo
        If (ActiveCell.Value = "Ent. Atrasada" Or ActiveCell.Value = "Ausencia" Or ActiveCell.Value = "Refrigerio Largo") Then
            ActiveCell.Offset(1, 0).Select
        Else
            Selection.EntireRow.Delete
        End If
    Next i
    'Declaración de variables de validación
    Dim a, b As Integer
    a = 1
    b = 0
    'Validación de existencia de registros
    Do While a <= 3
        If Cells(11 + b, 2).Value <> "" Then
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    'Seleccion por existencia de registros
    Select Case b
        Dim NroFila1, NroColumna1 As Integer
        Case 0
        Case 1
            Range("B11").Select
            Call Formato_Texto(11)
        Case Else
            Range("B11").End(xlDown).Select
            NroFila1 = ActiveCell.Row
            NroColumna1 = ActiveCell.Column
            Range(Cells(11, 2), Cells(NroFila1, NroColumna1)).Select
            Call Formato_Texto(11)
    End Select
    Range("B11").Select
End Sub
Private Sub Formato_Texto(num As Integer)
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    'Datos a texto
    Selection.TextToColumns Destination:=Range("B" & num), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
End Sub
Sub DNI_aTexto_PareoMarcajes()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("PareoMarcajes").Select
    'Agrego zoon a la hoja
    ActiveWindow.Zoom = 85
    'Columna DNI a texto - Selecciono datos
    Range("B12").Select
    'Declaración de variables de validación
    Dim a, b As Integer
    a = 1
    b = 0
    'Validación de existencia de registros
    Do While a <= 3
        If Cells(11 + b, 2).Value <> "" Then
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    'Seleccion por existencia de registros
    Select Case b
        Dim NroFila, NroColumna As Integer
        Case 0
        Case 1
            Range("B12").Select
            Call Formato_Texto(12)
        Case Else
            Range("B12").End(xlDown).Select
            NroFila = ActiveCell.Row
            NroColumna = ActiveCell.Column
            Range(Cells(12, 2), Cells(NroFila, NroColumna)).Select
            Call Formato_Texto(12)
    End Select
    Range("B12").Select
End Sub
Sub Formato_Dotacion_Ofisis()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Dotacion Ofisis").Select
    'Agrego zoon a la hoja
    ActiveWindow.Zoom = 90
    'Seleccionar todas las celdas y auto-ajuste(ancho)
    Cells.Select
    Cells.RowHeight = 15
    Cells.Font.Name = "Calibri"
    
    With Range("A1:P1")
        .Font.Size = 9
        .Font.Name = "Arial"
        .RowHeight = 40
        .ColumnWidth = 11
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Interior.ColorIndex = 40
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.ColorIndex = 1
    End With
    
    Cells.EntireColumn.AutoFit
    Range("A:D,H:I,K:L,N:P,U:U").EntireColumn.Hidden = True
    
    'Agrego titulo a la celda
    Range("Q1").FormulaLocal = "DNI"
    Range("R1").FormulaLocal = "TRABAJADOR"
    Range("S1").FormulaLocal = "APELLIDOS_NOMBRES"
    Range("T1").FormulaLocal = "PLANILLA"
    Range("U1").FormulaLocal = "DESCRIPCION"
    
    'Agrego formula
    Range("Q2").Formula = "=IFERROR(MID(M2,7,8),""-"")"
    Range("R2").Formula = "=E2"
    Range("S2").Formula = "=F2"
    Range("T2").Formula = "=G2"
    Range("U2").Formula = "=J2"
    
    'Conteo de datos
    Dim NroFilaFin, NroColumnaFin As Integer
    Range("M2").End(xlDown).Select
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    
    Application.Calculation = xlCalculationManual
    'Copio formula
    Range("Q2:U2").Copy Range(Cells(3, 17), Cells(NroFilaFin, NroColumnaFin + 8))
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
    Call Ordena_DotacionOfisis
    
    Range("Q:U").EntireColumn.Hidden = True
    'Dato "OK" para validar el Icono: Impresion
    Range("AZ1").FormulaLocal = "OK"
    Range("AZ1").Font.ColorIndex = 2
    'Inmovilizar paneles
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    'Fin del proceso Formato_Dotacion_Ofisis
    Range("E1").Select
End Sub
Sub Formato_Control_Disciplinario()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Control Disciplinario").Select
    'Agrego zoon a la hoja
    ActiveWindow.Zoom = 90
    Cells.Select
    Cells.RowHeight = 15
    Cells.Font.Name = "Calibri"
    
    With Range("A1:R1")
        .Font.Size = 9
        .Font.Name = "Arial"
        .RowHeight = 40
        .ColumnWidth = 10
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Interior.ColorIndex = 37
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.ColorIndex = 1
    End With
    
    Cells.EntireColumn.AutoFit
    
    Columns("K:K").Select
    Selection.Replace What:="2", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="8", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="9", Replacement:="C", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="1", Replacement:="D", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="7", Replacement:="E", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Dim NroFila, NroColumna As Integer
    Range("D2").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    
    Application.Calculation = xlCalculationManual
    Range("E2").Formula = "=IFERROR(VLOOKUP(C2,'Dotacion Ofisis'!E:Q,13,0),"""")"
    Range("F2").Formula = "=CONCATENATE(E2,MID(I2,1,4),K2)"
    Range("E2:F2").Copy Range(Cells(3, 5), Cells(NroFila, NroColumna + 2))
    Range("S1").FormulaLocal = "DIA"
    Range("S2").Formula = "=DAY(J2)"
    Range("T1").FormulaLocal = "MES"
    Range("T2").Formula = "=MONTH(J2)"
    Range("U1").FormulaLocal = "AÑO"
    Range("U2").Formula = "=YEAR(J2)"
    Range("S2:U2").Copy Range(Cells(2, 19), Cells(NroFila, NroColumna + 17))
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
    Call Ordena_ControlDisciplinario
    
    Range("S:U").ClearContents
    Range("A:B,E:H,M:Q").EntireColumn.Hidden = True
    'Ancho fijo a columnas
    Range("R:R").ColumnWidth = 70
    'Dato "OK" para validar el Icono: Impresion
    Range("AZ1").FormulaLocal = "OK"
    Range("AZ1").Font.ColorIndex = 2
    'Inmovilizar paneles
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    'Fin del proceso Formato_Control_Disciplinario
    Range("C1").Select
End Sub
