Attribute VB_Name = "CValidacionesProceso"
Sub Validacion()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Call Validacion1
End Sub
Private Sub Validacion1()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    '---INCIDENCIAS---
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
        Call Validacion2
    Else
        Sheets("Incidencias").Select
        MsgBox "Hoja 'Incidencias' no contiene datos.", vbOKOnly + vbCritical + vbDefaultButton1, "Macro Tardanzas"
    End If
End Sub
Private Sub Validacion2()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    '---PAREO MARCAJES---
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
        Call Validacion3
    Else
        Sheets("PareoMarcajes").Select
        MsgBox "Hoja 'PareoMarcajes' no contiene datos.", vbOKOnly + vbCritical + vbDefaultButton1, "Macro Tardanzas"
    End If
End Sub
Private Sub Validacion3()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    '---CONTROL DISCIPLINARIO---
    Sheets("Control Disciplinario").Select
    'Validacion
    Range("A2").Select
    'Declaración de variables de validación
    Dim e, f As Integer
    e = 1
    f = 0
    'Validación de registros
    Do While e <= 5
        If Cells(2 + f, 1).Value = "" Then
            'Cells(2 + f, 1).Select
            e = e + 1
            f = f + 1
        Else
            Exit Do
        End If
    Loop
    If (f = 0) Then
        Call Validacion4
    Else
        Sheets("Control Disciplinario").Select
        MsgBox "Hoja 'Control Disciplinario' no contiene datos.", vbOKOnly + vbCritical + vbDefaultButton1, "Macro Tardanzas"
    End If
End Sub
Private Sub Validacion4()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    '---DOTACION OFISIS---
    Sheets("Dotacion Ofisis").Select
    'Validacion
    Range("A2").Select
    'Declaración de variables de validación
    Dim g, h As Integer
    g = 1
    h = 0
    'Validación de registros
    Do While g <= 5
        If Cells(2 + h, 1).Value = "" Then
            'Cells(2 + h, 1).Select
            g = g + 1
            h = h + 1
        Else
            Exit Do
        End If
    Loop
    If (h = 0) Then
        Call PROCESO
    Else
        Sheets("Dotacion Ofisis").Select
        MsgBox "Hoja 'Dotacion Ofisis' no contiene datos.", vbOKOnly + vbCritical + vbDefaultButton1, "Macro Tardanzas"
    End If
End Sub
Private Sub PROCESO()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 02.01.2018
    Call DNI_aTexto_PareoMarcajes
    Call Formato_Dotacion_Ofisis
    Call Formato_Control_Disciplinario
    Call DNI_aTexto_Incidencias
    'Call Excesos_Colacion 'Pertenece a FORMATOECA
    Call Dato_fechas
    Call Ordena_Incidencias
    Call Info_Incidencia
    UserForm1.Show
End Sub
