Attribute VB_Name = "EExcesoColacion"
Sub Excesos_Colacion()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Incidencias").Select
    'Valido la cantidad en "dato"
    Dim dato As Integer
    Range("M1").NumberFormat = "General"
    Range("M1").Formula = "=COUNTIF(L:L,""Refrigerio Largo"")"
    dato = Range("M1").Value
    Range("M1").ClearContents
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
    
    If (dato > 0 Or b <> 0) Then
        'Tiene datos refrigerio largo o no tiene datos incidencias
    Else
        Call ObtengoDatos
        Call CopiaDatos
    End If
    'Fin del proceso
    Range("L11").Select
End Sub
Private Sub ObtengoDatos()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("PareoMarcajes").Select
    'Copio a una nueva hoja
    Worksheets("PareoMarcajes").Copy After:=Sheets(Sheets.Count)
    'Cambio de nombre de hoja
    Sheets("PareoMarcajes (2)").Name = "Excesos"
    Sheets("Excesos").Select
    Rows("1:10").Select
    Selection.EntireRow.Delete
    'Obtengo la fila y columna
    Dim NroFila, NroColumna, Counter As Integer
    Range("A2").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    'Selecciona todo el rango
    Counter = Range(Cells(2, 1), Cells(NroFila, NroColumna)).Count
    Dim i As Integer
    'Elimino filas distintas de Refrigerio Largo o Refrigerio Muy Largo
    Range("Q2").Select
    ActiveCell.Select
        For i = 1 To Counter
            If ActiveCell.Value <= "01:06" Then
                Selection.EntireRow.Delete
            Else
                ActiveCell.Offset(1, 0).Select
            End If
        Next i
    'Elimino columnas innecesarias
    Range("G:G,I:K,N:P,S:X,Z:AI").Delete
    Range("A2").Select
    'Declaracion de variables para validación
    Dim a, b, z As Integer
    a = 1
    b = 1
    z = 0
    'Validación de existencia de registros
    Do While a <= 5
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
        Dim NroFila2, NroColumna2 As Integer
        Case 0
        Case 1
            NroFila2 = 1
            NroColumna2 = 1
            Call datos_formato
            Range("H1").Copy Range("H2")
            Range("I1").Copy Range("I2")
            Range("L1").Copy Range("L2")
            Range("K2").Formula = "=CONCATENATE(IF(LEN(HOUR(J2-I2))=1,CONCATENATE(""0"",HOUR(J2-I2)),HOUR(J2-I2)),"":"",IF(LEN(MINUTE(J2-I2))=1,CONCATENATE(""0"",MINUTE(J2-I2)),MINUTE(J2-I2)))"
            Range("A2").Select
        Case Else
            'Obtengo datos de ultima fila y columna
            Range("A2").End(xlDown).Select
            NroFila2 = ActiveCell.Row
            NroColumna2 = ActiveCell.Column
            Call datos_formato
            'Copio datos que se repiten
            Application.Calculation = xlCalculationManual
            Range("H1").Copy Range(Cells(2, 8), Cells(NroFila2, NroColumna2 + 7))
            Range("I1").Copy Range(Cells(2, 9), Cells(NroFila2, NroColumna2 + 8))
            Range("L1").Copy Range(Cells(2, 12), Cells(NroFila2, NroColumna2 + 11))
            Range("K2").Formula = "=CONCATENATE(IF(LEN(HOUR(J2-I2))=1,CONCATENATE(""0"",HOUR(J2-I2)),HOUR(J2-I2)),"":"",IF(LEN(MINUTE(J2-I2))=1,CONCATENATE(""0"",MINUTE(J2-I2)),MINUTE(J2-I2)))"
            Range("K2").Copy Range(Cells(2 + 1, 11), Cells(NroFila2, NroColumna2 + 10))
            Application.Calculation = xlCalculationAutomatic
            Application.Calculate
            Range("A2").Select
    End Select
End Sub
Private Sub datos_formato()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Excesos").Select
    With Range("H1:L1")
        .Interior.ColorIndex = 2
        .Font.Bold = False
        .Font.ColorIndex = 1
    End With
    Range("H1").FormulaLocal = "Dur. Refrigerio"
    Range("I1").Formula = "=CONCATENATE(""01"","":"",""00"")"
    Range("J1").FormulaR1C1 = "Dura. Real"
    Range("K1").FormulaR1C1 = "Dif."
    Range("L1").FormulaR1C1 = "Refrigerio Largo"
    'Selecciono inicio registro de datos
    Range("K2").Select
End Sub
Private Sub CopiaDatos()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Excesos").Select
    Range("A2").Select
    'Declaración de variables de validación
    Dim a, b, z As Integer
    a = 1
    b = 1
    z = 0
    'Validación conteo de datos a copiar
    Do While a <= 5
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
            Range("L11").Select
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
            Range("L11").Select
    End Select
    'Borrar la hoja "Excesos"
    Application.DisplayAlerts = False
    Sheets("Excesos").Select
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
End Sub
