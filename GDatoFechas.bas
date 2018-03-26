Attribute VB_Name = "HDatoFechas"
Sub Dato_fechas()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 02.01.2018
    Sheets("PareoMarcajes").Select
    'Declaro las variables
    Dim f1, f2, f3, f4, f5, f6, f7, f8 As String
    Dim i As Integer
    'Asignacion de valores
    f1 = Cells(12, 8).Value
    f2 = Cells(13, 8).Value
    f3 = Cells(14, 8).Value
    f4 = Cells(15, 8).Value
    f5 = Cells(16, 8).Value
    f6 = Cells(17, 8).Value
    f7 = Cells(18, 8).Value
    f8 = Cells(19, 8).Value
    i = 0
    'Comparacion de datos
    Do
        If (f1 = f2) Then
            'MsgBox "Es igual f1 a f2"
            'Llamo a proceso para eliminar fechas
            Call Elimina_Fechas(f1, "", "")
            Exit Do
        End If
        If (f1 = f3) Then
            'MsgBox "Es igual f1 a f3"
            'Llamo a proceso para eliminar fechas
            Call Elimina_Fechas(f2, "", "")
            Exit Do
        End If
        If (f1 = f4) Then
            'MsgBox "Es igual f1 a f4"
            'Llamo a proceso para eliminar fechas
            Call Elimina_Fechas(f3, "", "")
            Exit Do
        End If
        If (f1 = f5) Then
            'MsgBox "Es igual f1 a f5"
            'Llamo a proceso para eliminar fechas
            Call Elimina_Fechas(f4, "", "")
            Exit Do
        End If
        If (f1 = f6) Then
            'MsgBox "Es igual f1 a f6"
            'Llamo a proceso para eliminar fechas
            Call Elimina_Fechas(f5, "", "")
            Exit Do
        End If
        If (f1 = f7) Then
            'MsgBox "Es igual f1 a f7"
            'Llamo a proceso para eliminar fechas
            Call Elimina_Fechas(f6, "", "")
            Exit Do
        End If
        If (f1 = f8) Then
            'MsgBox "Es igual f1 a f8"
            'Llamo a proceso para eliminar fechas y otros procesos
            Call Elimina_Fechas(f5, f6, f7)
            Call Info_Tolerancia
            Call Excesos_Tolerancia_F1
            Call Excesos_Tolerancia_F3
            Call Borra_Info_Tolerancia
            Exit Do
        End If
        i = i + 1
    Loop While (i < 1)
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
End Sub
Sub Elimina_Fechas(ByVal f1 As String, ByVal f2 As String, ByVal f3 As String)
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 02.01.2018
    Dim fecha1, fecha2, fecha3 As String
    fecha1 = f1
    fecha2 = f2
    fecha3 = f3
    'Eliminación de otros fechas
    Sheets("Incidencias").Select
    Dim NroFila, NroColumna, Conteo As Integer
    Range("L11").End(xlDown).Select
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    'Selecciona todo el rango
    Conteo = Range(Cells(11, 12), Cells(NroFila, NroColumna)).Count
    Dim i As Integer
    'Elimino filas distintas de Ent. Atrasada y Ausencia
    Range("G11").Select
    ActiveCell.Select
    For i = 1 To Conteo
        If (ActiveCell.Value = fecha1 Or ActiveCell.Value = fecha2 Or ActiveCell.Value = fecha3) Then
            ActiveCell.Offset(1, 0).Select
        Else
            Selection.EntireRow.Delete
        End If
    Next i
End Sub

