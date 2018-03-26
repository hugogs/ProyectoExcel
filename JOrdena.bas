Attribute VB_Name = "jOrdena"
Sub Ordena_Incidencias()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Incidencias").Select
    'Validacion
    Range("B11").Select
    'Declaración de variables de validación
    Dim a, b As Integer
    a = 1
    b = 0
    'Validación de registros
    Do While a <= 3
        If Cells(11 + b, 1).Value <> "" Then
            'Cells(11 + b, 1).Select
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    If (b <> 0) Then
        Sheets("Incidencias").Select
        'Quito Autofiltro
        Range("A10:L10").AutoFilter
        'Seleciona la ultima fila de la columna K
        Range("L10").End(xlDown).Select
        'Obtengo la fila y columna
        Dim NroFila, NroColumna As Integer
        NroFila = ActiveCell.Row
        NroColumna = ActiveCell.Column
        'Ancho de columnas y alto de filas
        Range(Cells(11, 8), Cells(NroFila, NroColumna - 4)).ColumnWidth = 11
        Columns("L:L").Select
        Selection.ColumnWidth = 17
        Range(Cells(11, 1), Cells(NroFila, NroColumna)).RowHeight = 15
        'Selecciona todo el rango a ordenar
        Range(Cells(10, 1), Cells(NroFila, NroColumna)).Select
        'Ordena por fecha:
        Selection.Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        'Ordena por apellido:
        Selection.Sort Key1:=Range("C10"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        'Agrego filtro
        Range("A10:L10").AutoFilter
        Range("L11").Select
        ActiveWorkbook.Save
    Else
    
    End If
End Sub
Sub Ordena_DotacionOfisis()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Dotacion Ofisis").Select
    'Validacion
    Range("A2").Select
    'Declaración de variables de validación
    Dim a, b As Integer
    a = 1
    b = 0
    'Validación de registros
    Do While a <= 3
        If Cells(2 + b, 1).Value <> "" Then
            'Cells(11 + b, 1).Select
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    If (b <> 0) Then
        Sheets("Dotacion Ofisis").Select
        'Quito Autofiltro
        Range("A1:P1").AutoFilter
        'Seleciona la ultima fila de la columna K
        Range("P1").End(xlDown).Select
        'Obtengo la fila y columna
        Dim NroFila1, NroColumna1 As Integer
        NroFila1 = ActiveCell.Row
        NroColumna1 = ActiveCell.Column
        'Selecciona todo el rango a ordenar
        Range(Cells(1, 1), Cells(NroFila1, NroColumna1)).Select
        'Ordena por apellido:
        Selection.Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        'Agrego filtro
        Range("A1:P1").AutoFilter
        Range("E2").Select
        ActiveWorkbook.Save
    Else
    
    End If
End Sub
Sub Ordena_ControlDisciplinario()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Control Disciplinario").Select
    'Validacion
    Range("A2").Select
    'Declaración de variables de validación
    Dim a, b As Integer
    a = 1
    b = 0
    'Validación de registros
    Do While a <= 3
        If Cells(2 + b, 1).Value <> "" Then
            'Cells(11 + b, 1).Select
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    If (b <> 0) Then
        Sheets("Control Disciplinario").Select
        'Quito Autofiltro
        Range("A1:R1").AutoFilter
        'Obtengo la fila y columna
        Dim NroFila, NroColumna As Integer
        Range("D2").End(xlDown).Select
        NroFila = ActiveCell.Row
        NroColumna = ActiveCell.Column
        'Selecciona todo el rango a ordenar
        Range(Cells(1, 1), Cells(NroFila, NroColumna + 17)).Select
        'Ordena por DIA:
        Selection.Sort Key1:=Range("S1"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        'Ordena por MES:
        Selection.Sort Key1:=Range("T1"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        'Ordena por AÑO:
        Selection.Sort Key1:=Range("U1"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        'Ordena por Apellido:
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        'Agrego filtro
        Range("A1:R1").AutoFilter
        Range("C2").Select
        ActiveWorkbook.Save
    Else
    
    End If
End Sub
Sub Borra_Info_Tolerancia()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("PareoMarcajes").Visible = xlSheetVisible
    Sheets("PareoMarcajes").Select
    
    If (Range("L1").Value = "BendicemeDios") Then
        Sheets("IMPRESION").Visible = xlSheetVisible
        Sheets("IMPRESION").Select
        Range("A:D,H:I,K:L,N:P").EntireColumn.Hidden = False
        'No borra las fórmulas
    Else
        Sheets("PareoMarcajes").Select
        Range("AJ:AM").ClearContents
        Sheets("Dotacion Ofisis").Select
        Range("Q:AW").ClearContents
        Range("A:D,H:I,K:L,N:U").EntireColumn.Hidden = True
        Range("E2").Select
        Sheets("IMPRESION").Visible = xlSheetVeryHidden
    End If
    'Este paso es "IMPORTANTE" para los datos de fin de semana
    Sheets("Dotacion Ofisis").Select
    'Agrego titulo a la celda
    Range("Q1").FormulaLocal = "DNI"
    'Agrego formula
    Range("Q2").Formula = "=IFERROR(MID(M2,7,8),""-"")"
    
    'Conteo de datos
    Dim NroFilaFin, NroColumnaFin As Integer
    Range("M2").End(xlDown).Select
    NroFilaFin = ActiveCell.Row
    NroColumnaFin = ActiveCell.Column
    
    Application.Calculation = xlCalculationManual
    'Copio formula
    Range("Q2").Copy Range(Cells(3, 17), Cells(NroFilaFin, NroColumnaFin + 4))
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Range("Q:Q").EntireColumn.Hidden = True
End Sub
