Attribute VB_Name = "AFormatoEca"
Sub FormatoEca(Control As IRibbonControl)
'Sub FormatoEca()
    'Dedicado a mi Señor Todopoderoso
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Call Validacion_Incidencia
    Call Validacion_PareoMarcaje
    Call Excesos_Colacion
    Call Ordena_Incidencias
    On Error Resume Next
    Call Validacion_ResumenHoras
    Call Validacion_ResumenHorasDetalle
    'Guarda los cambios del archivo
    ActiveWorkbook.Save
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Private Sub Validacion_Incidencia()
    'Dedicado a mi Señor Todopoderoso
    '---INCIDENCIAS---
    Sheets("Incidencias").Visible = xlSheetVisible
    Sheets("Incidencias").Select
    'Validacion
    Range("A11").Select
    'Declaración de variables de validación
    Dim a, b As Integer
    a = 1
    b = 0
    'Validación de registros
    Do While a <= 5
        If Cells(11 + b, 1).Value = "" Then
            'Cells(11 + b, 1).Select
            a = a + 1
            b = b + 1
        Else
            Exit Do
        End If
    Loop
    'MsgBox ("El valor de 'valor':" & " es " & b)
    If (b = 0) Then
        Call Formato_Incidencia
    Else
        Sheets("Incidencias").Select
        MsgBox "Hoja 'Incidencias' no contiene datos.", vbOKOnly + vbCritical + vbDefaultButton1, "Productividad"
    End If
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
End Sub
Private Sub Formato_Incidencia()
    'Dedicado a mi Señor Todopoderoso
    Sheets("Incidencias").Select
    'Seleccionar todas las celdas y acomoda
    Cells.Select
    Cells.RowHeight = 15
    Cells.EntireColumn.AutoFit
    'Ocultar la fila 8, alto a fila 9 y alto fila 10
    Rows("8:8").EntireRow.Hidden = True
    Rows(9).RowHeight = 5
    Rows(10).RowHeight = 24.75
    'Agregar filtro a la fila 10
    Rows("10:10").AutoFilter
    'Inmovilizar los paneles
    Rows("11:11").Select
    ActiveWindow.FreezePanes = True
    'Activar el Zoom de la hoja
    ActiveWindow.Zoom = 90
    'Declaro variables para obtener Nro fila y columna
    Dim NroFila, NroColumna As Integer
    'Obtengo la fila y columna
    Range("A10").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    'Declaro variable para almacenar la cantidad de datos
    Dim Conteo As Integer
    'Obtengo la cantidad de datos
    Conteo = Range(Cells(10, 1), Cells(NroFila, NroColumna)).Count
    'Declaro variable para bucle
    Dim i As Integer
    Range("L10").Select
    ActiveCell.Select
    For i = 1 To Conteo
        If ActiveCell.Value = "Permenencia menor a lo planificado" Or ActiveCell.Value = "Permenencia mayor a lo planificado" Or ActiveCell.Value = "Muchas horas seguidas" Then
            Selection.EntireRow.Delete
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next i
    'Ancho fijo a columnas
    Range("A:A").ColumnWidth = 13.3
    Range("B:B").ColumnWidth = 8.3
    Range("D:D").ColumnWidth = 7.3
    Range("G:G").ColumnWidth = 9.2
    Range("H:H").ColumnWidth = 6.5
    Range("I:I").ColumnWidth = 9.3
    Range("J:J").ColumnWidth = 4.8
    Range("K:K").ColumnWidth = 8.6
    Range("L:L").ColumnWidth = 17
    'Fin de proceso
    Range("L11").Select
End Sub
Private Sub Validacion_PareoMarcaje()
    'Dedicado a mi Señor Todopoderoso
    '---PAREO MARCAJES---
    Sheets("PareoMarcajes").Visible = xlSheetVisible
    Sheets("PareoMarcajes").Select
    'Validacion
    Range("A12").Select
    'Declaración de variables de validación
    Dim c, d As Integer
    c = 1
    d = 0
    'Validación de registros
    Do While c <= 5
        If Cells(12 + d, 1).Value = "" Then
            'Cells(12 + d, 1).Select
            c = c + 1
            d = d + 1
        Else
            Exit Do
        End If
    Loop
    If (d = 0) Then
        Call Formato_PareoMarcaje
    Else
        Sheets("PareoMarcajes").Select
        MsgBox "Hoja 'PareoMarcajes' no contiene datos.", vbOKOnly + vbCritical + vbDefaultButton1, "Productividad"
    End If
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
End Sub
Private Sub Formato_PareoMarcaje()
    'Dedicado a mi Señor Todopoderoso
    Sheets("PareoMarcajes").Select
    Cells.Select
    Cells.RowHeight = 15
    Cells.EntireColumn.AutoFit
    'Ocultar la fila 8, alto a fila 9 y alto fila 11
    Rows("8:8").EntireRow.Hidden = True
    Rows(9).RowHeight = 5
    Rows(11).RowHeight = 24.75
    'Agrego filtro y zoom
    Rows("11:11").AutoFilter
    ActiveWindow.Zoom = 85
    'Oculto columnas
    Columns("AC:AI").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
    'Desconbinar celdas
    Range("A1:K9").UnMerge
    'Ocultar columna G
    Columns("G:G").Select
    Range("G:G").ColumnWidth = 9.2
    Selection.EntireColumn.Hidden = True
    'Combinar celdas
    Range("A1:E1").Merge
    Range("F1:K1").Merge
    Range("A2:K2").Merge
    Range("A3:K3").Merge
    Range("A4:K4").Merge
    Range("A5:E5").Merge
    Range("F5:K5").Merge
    Range("A6:E6").Merge
    Range("F6:K6").Merge
    Range("A7:E7").Merge
    Range("F7:K7").Merge
    Range("A8:E8").Merge
    Range("F8:K8").Merge
    Range("A9:K9").Merge
    'Inmovilizo paneles
    Range("E12").Select
    ActiveWindow.FreezePanes = True
    'Celdas con ancho fijo
    Range("A:A").ColumnWidth = 13.3
    Range("B:B").ColumnWidth = 8.3
    Range("D:D").ColumnWidth = 7.3
    'Range("G:G").ColumnWidth = 9.2
    Range("H:H").ColumnWidth = 13.3
    
    Range("I:I").ColumnWidth = 4.8
    Range("J:J").ColumnWidth = 4.8
    Range("K:K").ColumnWidth = 5.3
    Range("L:L").ColumnWidth = 10.8
    
    Range("M:M").ColumnWidth = 4.8
    Range("N:N").ColumnWidth = 4.8
    Range("O:O").ColumnWidth = 5.3
    Range("P:P").ColumnWidth = 10.8
    
    Range("Q:Q").ColumnWidth = 6.8
    
    Range("R:R").ColumnWidth = 4.8
    Range("S:S").ColumnWidth = 4.8
    Range("T:T").ColumnWidth = 5.3
    Range("U:U").ColumnWidth = 10.8
    
    Range("V:V").ColumnWidth = 4.8
    Range("W:W").ColumnWidth = 4.8
    Range("X:X").ColumnWidth = 5.3
    Range("Y:Y").ColumnWidth = 10.8
    
    Range("Z:Z").ColumnWidth = 4.8
    Range("AA:AA").ColumnWidth = 4.8
    Range("AB:AB").ColumnWidth = 5.3
    'Fin del proceso
    Range("I12").Select
End Sub
Private Sub Validacion_ResumenHoras()
    'Dedicado a mi Señor Todopoderoso
    '---RESUMEN HORAS---
    Sheets("ResumenHoras").Visible = xlSheetVisible
    Sheets("ResumenHoras").Select
    'Validacion
    Range("A13").Select
    'Declaración de variables de validación
    Dim c, d As Integer
    c = 1
    d = 0
    'Validación de registros
    Do While c <= 5
        If Cells(13 + d, 1).Value = "" Then
            'Cells(13 + d, 1).Select
            c = c + 1
            d = d + 1
        Else
            Exit Do
        End If
    Loop
    If (d = 0) Then
        Formato_ResumenHoras
    Else
        Application.DisplayAlerts = False
        Sheets("ResumenHoras").Select
        ActiveWindow.SelectedSheets.Delete
        'Application.DisplayAlerts = True
    End If
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
End Sub
Private Sub Formato_ResumenHoras()
    'Dedicado a mi Señor Todopoderoso
    Sheets("ResumenHoras").Select
    'Formato a 02 decimales
    Range("C13").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.NumberFormat = "0.00"
    'Agrego Autoajuste de ancho a todas las columnas
    Cells.Select
    Cells.RowHeight = 15
    Cells.EntireColumn.AutoFit
    'Oculto la fila 8, alto a fila 9 e elimino las filas 10 y 11
    Rows("8:8").EntireRow.Hidden = True
    Rows(9).RowHeight = 5
    Rows("10:11").EntireRow.Hidden = True
    'Alto de la nueva fila 10
    Rows("12:12").RowHeight = 25
    'Agrego el filtro
    Range("A12:U12").AutoFilter
    'Inmovilizar los paneles
    Rows("13:13").Select
    ActiveWindow.FreezePanes = True
    'Alineo a la izquierda los nombres de Dpto.
    Range("A13:B13").Select
    Range(Selection, Selection.End(xlDown)).HorizontalAlignment = xlLeft
    'Color Rojo a la celda horas extras PT
    Range("F12").Interior.Color = 255
    'Agrego zoon a la hoja
    ActiveWindow.Zoom = 90
    'Agrego titulo a la celda
    Range("O12").FormulaR1C1 = "% Horas No Trabajadas"
    
    'Declaro variables
    Dim NroFila, NroColumna As Integer
    'Obtengo el total de datos (filas y columnas) a trabajar
    Range("O13").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    'Agrego formula
    Application.Calculation = xlCalculationManual
    Range("O13").Formula = "=IFERROR(IF(L13/J13<0,""-"",L13/J13),""-"")"
    'Copio formula
    Cells(13, 15).Copy
    Range(Cells(14, 15), Cells(NroFila, NroColumna)).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Range(Cells(13, 15), Cells(NroFila, NroColumna)).Select
    With Selection
        .Style = "Percent"
        .NumberFormat = "0.00%"
    End With

    'Copio y pego en valores
    Range(Cells(13, 15), Cells(NroFila, NroColumna)).Copy
    Range(Cells(13, 15), Cells(NroFila, NroColumna)).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    'Agrego estilo "-" a todos los ceros,excepto la columna de %
    Range(Cells(13, 3), Cells(NroFila, NroColumna - 1)).Select
    Selection.Style = "Comma"
    Range(Cells(13, 16), Cells(NroFila, NroColumna + 6)).Select
    Selection.Style = "Comma"

    'Celdas con ancho fijo
    Range("C:C").ColumnWidth = 8.3
    
    Range("D:D").ColumnWidth = 8.2
    Range("E:E").ColumnWidth = 8.2
    Range("F:F").ColumnWidth = 7.9
    Range("G:G").ColumnWidth = 6.2
    Range("H:H").ColumnWidth = 6.6
    
    Range("I:I").ColumnWidth = 8
    Range("J:J").ColumnWidth = 10.3
    Range("K:K").ColumnWidth = 9
    Range("L:L").ColumnWidth = 9
    
    Range("M:M").ColumnWidth = 6.9
    Range("N:N").ColumnWidth = 10.3
    Range("O:O").ColumnWidth = 9.5
    Range("P:P").ColumnWidth = 8.8
    Range("Q:Q").ColumnWidth = 7.3
    Range("R:R").ColumnWidth = 9.6
    Range("S:S").ColumnWidth = 7
    Range("T:T").ColumnWidth = 8.2
    Range("U:U").ColumnWidth = 9.9
    'Fin del proceso
    Range("C13").Select
End Sub
Private Sub Validacion_ResumenHorasDetalle()
    'Dedicado a mi Señor Todopoderoso
    '---RESUMEN HORAS DETALLE---
    Sheets("ResumenHorasDetalle").Visible = xlSheetVisible
    Sheets("ResumenHorasDetalle").Select
    'Validacion
    Range("A13").Select
    'Declaración de variables de validación
    Dim c, d As Integer
    c = 1
    d = 0
    'Validación de registros
    Do While c <= 5
        If Cells(13 + d, 1).Value = "" Then
            'Cells(13 + d, 1).Select
            c = c + 1
            d = d + 1
        Else
            Exit Do
        End If
    Loop
    If (d = 0) Then
        Formato_ResumenHorasDetalle
    Else
        Application.DisplayAlerts = False
        Sheets("ResumenHorasDetalle").Select
        ActiveWindow.SelectedSheets.Delete
        'Application.DisplayAlerts = True
    End If
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
End Sub
Private Sub Formato_ResumenHorasDetalle()
    'Dedicado a mi Señor Todopoderoso
    Sheets("ResumenHorasDetalle").Select
    'Formato a 02 decimales
    Range("G13").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.NumberFormat = "0.00"
    'Agrego Autoajuste de ancho a todas las columnas
    Cells.Select
    Cells.RowHeight = 15
    Cells.EntireColumn.AutoFit
    'Oculto la fila 8, alto de fila 9 e elimino las filas 10 y 11
    Rows("8:8").EntireRow.Hidden = True
    Rows(9).RowHeight = 5
    Rows("10:11").EntireRow.Hidden = True
    'Alto de la nueva fila 10
    Rows("12:12").RowHeight = 25
    'Agrego el filtro
    Range("A12:Z12").AutoFilter
    'Inmovilizo paneles
    Range("E13").Select
    ActiveWindow.FreezePanes = True
    'Alineo a la izquierda los nombres de Dpto.
    Range("C13:F13").Select
    Range(Selection, Selection.End(xlDown)).HorizontalAlignment = xlLeft
    'Agrego zoon a la hoja
    ActiveWindow.Zoom = 85
    'Agrego titulo a la celda
    Range("T12").FormulaR1C1 = "% Horas No Trabajadas"
    
    'Declaro variables
    Dim NroFila, NroColumna As Integer
    'Obtengo el total de datos (filas y columnas) a trabajar
    Range("T13").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    'Agrego formula
    Application.Calculation = xlCalculationManual
    Range("T13").Formula = "=IFERROR(IF(Q13/O13<0,""-"",Q13/O13),""-"")"
    'Copio formula
    Cells(13, 20).Copy
    Range(Cells(14, 20), Cells(NroFila, NroColumna)).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Range(Cells(13, 20), Cells(NroFila, NroColumna)).Select
    With Selection
        .Style = "Percent"
        .NumberFormat = "0.00%"
    End With
    
    'Copio y pego en valores
    Range(Cells(13, 20), Cells(NroFila, NroColumna)).Copy
    Range(Cells(13, 20), Cells(NroFila, NroColumna)).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    'Agrego estilo "-" a todos los ceros,excepto la columna de %
    Range(Cells(13, 7), Cells(NroFila, NroColumna - 1)).Select
    Selection.Style = "Comma"
    Range(Cells(13, 21), Cells(NroFila, NroColumna + 6)).Select
    Selection.Style = "Comma"
    
    'Celdas con ancho fijo
    Range("B:B").ColumnWidth = 8.3
    Range("D:D").ColumnWidth = 7.3
    
    Range("G:G").ColumnWidth = 9.6
    Range("H:H").ColumnWidth = 10.2
    
    Range("I:I").ColumnWidth = 8.2
    Range("J:J").ColumnWidth = 8.2
    Range("K:K").ColumnWidth = 7.9
    Range("L:L").ColumnWidth = 6.2
    
    Range("M:M").ColumnWidth = 7.8
    Range("N:N").ColumnWidth = 10.2
    Range("O:O").ColumnWidth = 10.3
    Range("P:P").ColumnWidth = 9
    Range("Q:Q").ColumnWidth = 9
    Range("R:R").ColumnWidth = 7.2
    Range("S:S").ColumnWidth = 10.3
    Range("T:T").ColumnWidth = 9.5
    
    Range("U:U").ColumnWidth = 8.8
    Range("V:V").ColumnWidth = 7.3
    Range("W:W").ColumnWidth = 9.6
    Range("X:X").ColumnWidth = 7
    Range("Y:Y").ColumnWidth = 8.2
    Range("Z:Z").ColumnWidth = 9.9
    
    'Fin del Proceso
    Range("G11").Select
End Sub
