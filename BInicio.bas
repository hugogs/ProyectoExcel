Attribute VB_Name = "BInicio"
Sub Inicio(Control As IRibbonControl)
'Sub INICIO()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 02.01.2018
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Llamo al proceso para Guardar
    Call Guardar_En_Escritorio_Usuario
    
    'Declaracion de variables
    Dim h1, h2, h3, h4 As Boolean
    h1 = False
    h2 = False
    h3 = False
    h4 = False
    
    'Formato a Hoja "Dotacion Ofisis"
    For i = 1 To Worksheets.Count
        If (Worksheets(i).Name = "Dotacion Ofisis") Then
            h1 = True
            Exit For
        End If
    Next i
    If (h1 = True) Then
        'MsgBox ("La hoja ya existe")
    Else
        Sheets.Add(After:=Sheets("PareoMarcajes")).Name = "Dotacion Ofisis"
        Sheets("Dotacion Ofisis").Visible = xlSheetVisible
        Call Encabezado_Dotacion_Ofisis
    End If
    
    'Formato a Hoja "Control Disciplinario"
    For j = 1 To Worksheets.Count
        If (Worksheets(j).Name = "Control Disciplinario") Then
            h2 = True
            Exit For
        End If
    Next j
    If (h2 = True) Then
        'MsgBox ("La hoja ya existe")
    Else
        Sheets.Add(Before:=Sheets("Dotacion Ofisis")).Name = "Control Disciplinario"
        Sheets("Control Disciplinario").Visible = xlSheetVisible
        Call Encabezado_Control_Disciplinario
    End If
    
    'Elimino Hoja "ResumenHoras"
    For x = 1 To Worksheets.Count
        If (Worksheets(x).Name = "ResumenHoras") Then
            h3 = True
            Exit For
        End If
    Next x
    If (h3 = True) Then
        Application.DisplayAlerts = False
        Sheets("ResumenHoras").Select
        ActiveWindow.SelectedSheets.Delete
        'Application.DisplayAlerts = True
    Else
        'MsgBox ("La hoja no existe")
    End If
    
    'Elimino Hoja "ResumenHorasDetalle"
    For y = 1 To Worksheets.Count
        If (Worksheets(y).Name = "ResumenHorasDetalle") Then
            h4 = True
            Exit For
        End If
    Next y
    If (h4 = True) Then
        Application.DisplayAlerts = False
        Sheets("ResumenHorasDetalle").Select
        ActiveWindow.SelectedSheets.Delete
        'Application.DisplayAlerts = True
    Else
        'MsgBox ("La hoja no existe")
    End If
    
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
    'Guarda los cambios del archivo
    ActiveWorkbook.Save
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Sub DISCIPLINA(Control As IRibbonControl)
'Sub DISCIPLINA()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 02.01.2018
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim Hoja1, Hoja2 As Boolean
    Hoja1 = False
    Hoja2 = False
    For i = 1 To Worksheets.Count
        If (Worksheets(i).Name = "Control Disciplinario") Then
            Hoja1 = True
            Exit For
        End If
    Next i
    For j = 1 To Worksheets.Count
        If (Worksheets(j).Name = "Dotacion Ofisis") Then
            Hoja2 = True
            Exit For
        End If
    Next j
    If (Hoja1 = False Or Hoja2 = False) Then
        MsgBox "Empezar el proceso por el Icono 'Inicio'", vbOKOnly + vbCritical + vbDefaultButton1, "Macro Tardanzas"
    Else
        Call Validacion
    End If
    'Guarda los cambios del archivo
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Sub IMPRIMIR(Control As IRibbonControl)
'Sub IMPRIMIR()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 02.01.2018
    Dim dato1, dato2 As String
    dato1 = ""
    dato2 = ""
    On Error Resume Next
    Sheets("Control Disciplinario").Select
    dato1 = Range("AZ1").Value
    Sheets("Dotacion Ofisis").Select
    dato2 = Range("AZ1").Value
    If (dato1 = "" And dato2 = "") Then
        MsgBox "Faltan los datos del Icono 'Procesar'", vbOKOnly + vbCritical + vbDefaultButton1, "Macro Tardanzas"
    Else
        Call Impresion
        UserForm1.Show
    End If
    'Guarda los cambios del archivo
    ActiveWorkbook.Save
End Sub
Private Sub Encabezado_Control_Disciplinario()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 26.01.2018
    Sheets("Control Disciplinario").Visible = xlSheetVisible
    Sheets("Control Disciplinario").Select
    Range("A:Z").EntireColumn.Hidden = False
    ActiveWindow.Zoom = 90
    Cells.Select
    Selection.ClearContents
    Range("A1").FormulaLocal = "EMPRESA"
    Range("B1").FormulaLocal = "DESCRIPCION"
    Range("C1").FormulaLocal = "TRABAJADOR"
    Range("D1").FormulaLocal = "APELLIDOS_NOMBRES"
    Range("E1").FormulaLocal = "SITUACION_TRABAJADOR"
    Range("F1").FormulaLocal = "CORRELATIVO"
    Range("G1").FormulaLocal = "SITUACION"
    Range("H1").FormulaLocal = "FALTA"
    Range("I1").FormulaLocal = "DESCRIPCION"
    Range("J1").FormulaLocal = "FECHA_FALTA"
    Range("K1").FormulaLocal = "SANCION"
    Range("L1").FormulaLocal = "DESCRIPCION"
    Range("M1").FormulaLocal = "FECHA_INICIO"
    Range("N1").FormulaLocal = "FECHA_FINAL"
    Range("O1").FormulaLocal = "ARCHIVO"
    Range("P1").FormulaLocal = "TRABAJADOR_INFORMA"
    Range("Q1").FormulaLocal = "APELLIDOS_NOMBRES"
    Range("R1").FormulaLocal = "OBSERVACIONES"
    With Range("A1:R1")
        .Font.Size = 9
        .Font.Name = "Arial"
        .RowHeight = 40
        .ColumnWidth = 10
        .Font.Color = RGB(0, 0, 0)
        '.Font.Bold = True
        '.Interior.ColorIndex = 37
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        '.Borders.ColorIndex = 1
    End With
    Range("A2").Select
    Cells.Select
    'Sheets("Control Disciplinario").Visible = xlSheetVeryHidden
End Sub
Private Sub Encabezado_Dotacion_Ofisis()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 26.01.2018
    Sheets("Dotacion Ofisis").Visible = xlSheetVisible
    Sheets("Dotacion Ofisis").Select
    Range("A:Z").EntireColumn.Hidden = False
    ActiveWindow.Zoom = 90
    Cells.Select
    Selection.ClearContents
    Range("A1").FormulaLocal = "EMPRESA"
    Range("B1").FormulaLocal = "NOMBRE"
    Range("C1").FormulaLocal = "UNIDAD"
    Range("D1").FormulaLocal = "DESCRIPCION"
    Range("E1").FormulaLocal = "TRABAJADOR"
    Range("F1").FormulaLocal = "APELLIDOS_NOMBRES"
    Range("G1").FormulaLocal = "PLANILLA"
    Range("H1").FormulaLocal = "DESCRIPCION"
    Range("I1").FormulaLocal = "PUESTO_TRABAJO"
    Range("J1").FormulaLocal = "DESCRIPCION"
    Range("K1").FormulaLocal = "CALIFICACION_TRABAJADOR"
    Range("L1").FormulaLocal = "DESCRIPCION"
    Range("M1").FormulaLocal = "DOCUMENTO_IDENTIDAD"
    Range("N1").FormulaLocal = "FECHA_INGRESO"
    Range("O1").FormulaLocal = "FECHA_CESE"
    Range("P1").FormulaLocal = "SITUACION_TRABAJADOR"
    With Range("A1:P1")
        .Font.Size = 9
        .Font.Name = "Arial"
        .RowHeight = 40
        .ColumnWidth = 11
        .Font.Color = RGB(0, 0, 0)
        '.Font.Bold = True
        '.Interior.ColorIndex = 37
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        '.Borders.ColorIndex = 1
    End With
    Range("A2").Select
    Cells.Select
    'Sheets("Dotacion Ofisis").Visible = xlSheetVeryHidden
End Sub
Private Sub Guardar_En_Escritorio_Usuario()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 02.01.2018
    'Declaro variables
    Dim mydesk, nombre As String
    On Error Resume Next
    'Obtengo los datos
    mydesk = CreateObject("wscript.shell").specialfolders("desktop") & "\"
    nombre = ActiveWorkbook.Name
    'Valido si existe la carpeta "Disciplina Asistencia", caso contrario se crea
    Path = mydesk & "Disciplina Asistencia"
    If Dir(Path, vbDirectory) = "" Then
        MkDir Path
    End If
    'Valido si existe la carpeta "nombre" dentro de la carpeta "Disciplina Asistencia", caso contrario se crea
    Path1 = mydesk & "Disciplina Asistencia" & "\" & nombre
    If Dir(Path1, vbDirectory) = "" Then
        MkDir Path1
    End If
    'Ubico el directorio en donde se guardara el archivo "nombre"
    ChDir mydesk & "Disciplina Asistencia" & "\" & nombre
    'Guardo sin preguntar el archivo
    ActiveWorkbook.SaveAs mydesk & "Disciplina Asistencia" & "\" & nombre & "\" & nombre
End Sub
