Attribute VB_Name = "kImpresion"
Sub Impresion()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("IMPRESION").Visible = xlSheetVisible
    Application.ScreenUpdating = False
    'Application.DisplayAlerts = False
    'Application.EnableEvents = False
    
    'Dato para REVISION
    Range("AL1").FormulaLocal = "=PareoMarcajes!L1"
    Range("AL1").Font.ColorIndex = 2
    Dim dato As String
    dato = Range("AL1").Value
    
    'Tralado de logo
    Dim logo As String
    Sheets("PareoMarcajes").Select
    logo = Range("A1").Value
    'Validacion de traslado de logo
    Sheets("IMPRESION").Select
    If (logo = "SODIMAC PERU S.A") Then
        'Traslado a la izquierda el icono Maestro
        ActiveSheet.Shapes.Range(Array("Group 5")).Select
        Selection.ShapeRange.IncrementLeft -650
    Else
        'Traslado a la izquierda el icono Sodimac
        ActiveSheet.Shapes.Range(Array("Picture 4")).Select
        Selection.ShapeRange.IncrementLeft -800
    End If
    
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
    'Recorrido de datos a imprimir
    Select Case b
        Case 0
        Case 1
            'Agrego las formulas
            Sheets("IMPRESION").Select
            Call Formulas_Impresion
            
            'Configuro la hoja a imprimir
            Application.PrintCommunication = False
            Call Configurar
            Application.PrintCommunication = True
            
            Sheets("Incidencias").Select
            Range("B11").Copy
            Sheets("IMPRESION").Select
            Range("BE7").PasteSpecial Paste:=xlPasteValues
            
            'Valido en caso de REVISION
            If (dato = "BendicemeDios") Then
                Call GenerarPDF
            Else
                ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, IgnorePrintAreas:=False
                Call GenerarPDF
            End If
            'Valido en caso de REVISION, caso contrario se borran las formulas
            If (dato = "BendicemeDios") Then
                'No borro formulas
            Else
                Call Borra_Formulas_Impresion
            End If
        Case Else
            Sheets("IMPRESION").Select
            Call Formulas_Impresion
            
            Application.PrintCommunication = False
            Call Configurar
            Application.PrintCommunication = True
            
            Sheets("Incidencias").Select
            Dim Conteo, NroFila, NroColumna, i As Integer
            Range("B11").End(xlDown).Select
            NroFila = ActiveCell.Row
            NroColumna = ActiveCell.Column
            Conteo = Range(Cells(11, 2), Cells(NroFila, NroColumna)).Count
            i = 0
            
            For i = 0 To Conteo - 1
                Sheets("Incidencias").Select
                Range("B" & 11 + i).Copy
                Sheets("IMPRESION").Select
                Range("BE7").PasteSpecial Paste:=xlPasteValues
                
                If (dato = "BendicemeDios") Then
                    Call GenerarPDF
                Else
                    ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, IgnorePrintAreas:=False
                    Call GenerarPDF
                End If
            Next i
            
            If (dato = "BendicemeDios") Then
                'No borro formulas
            Else
                Call Borra_Formulas_Impresion
            End If
    End Select
    
    'Validacion de retorno de logo
    Sheets("IMPRESION").Select
    If (logo = "SODIMAC PERU S.A") Then
        'Traslado a la derecha el icono Maestro
        ActiveSheet.Shapes.Range(Array("Group 5")).Select
        Selection.ShapeRange.IncrementLeft 650
    Else
        'Traslado a la derecha el icono Sodimac
        ActiveSheet.Shapes.Range(Array("Picture 4")).Select
        Selection.ShapeRange.IncrementLeft 800
    End If
    Range("A1").Select
    Application.ScreenUpdating = True
    'Application.DisplayAlerts = True
    'Application.EnableEvents = True
    If (dato = "BendicemeDios") Then
        'No oculta la hoja
    Else
        Sheets("IMPRESION").Visible = xlSheetVeryHidden
    End If
    'Posicion final al termino de las impresiones
    Sheets("Incidencias").Select
    Range("L11").Select
    '-------Liberar memoria
    Set Portapapeles = New MSForms.DataObject
    Portapapeles.Clear
End Sub
Private Sub Configurar()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("IMPRESION").Select
    'Establece área de impresión
    ActiveSheet.PageSetup.PrintArea = "$A$1:$AI$54"
    With ActiveSheet.PageSetup
        'Establece el margen izquierdo, derecho, superior e inferior en centimetros
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        'Establece el margen del encabezado y pie de pagina en centimetros
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
        'Establece calidad de impresión puede ser 300, 1200, 2400
        .PrintQuality = 600
        'centra horizontalmente y verticalmente
        .CenterHorizontally = True
        .CenterVertically = True
        'Horientación vertical de la hoja
        .Orientation = xlPortrait
        'Imprime en Banco y Negro
        .BlackAndWhite = True
        ' Ajusta a una página de ancho y 1 de alto
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        'Escala de impresión se puede reducir a un porcentaje X, en este caso es falso
        '.Zoom = False
        'Si se utiliza zoom y puede disminuir en un x porcentaje la configuración de la página
        '.Zoom = 100
    End With
End Sub
Private Sub GenerarPDF()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Dim nombre, motivo As String
    Sheets("IMPRESION").Select
    nombre = Cells(4, 9).Value
    motivo = Cells(11, 57).Value
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=ThisWorkbook.Path & "\" & "" & nombre & " - " & motivo & ".pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub
Private Sub Formulas_Impresion()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Application.Calculation = xlCalculationManual
    '---- 1RA PARTE de la columna BD a BK ----
    Sheets("IMPRESION").Select
    Range("BD6").FormulaLocal = "FECHA"
    Range("BE6").Formula = "=TODAY()"
    Range("BD7").FormulaLocal = "INGRESE DNI"
    Range("BD8").FormulaLocal = "Tiempo Tardanza"
    Range("BE8").Formula = "=IFERROR(IF(BE11<>""Ausencia"",VLOOKUP(BE7,Incidencias!B:K,10,0),""""),"""")"
    Range("BF8").Formula = "=IF(BE11<>""Ausencia"",BE8*24,"""")"
    Range("BD9").FormulaLocal = "MALLA"
    Range("BE9").Formula = "=IFERROR(VLOOKUP(BE7,Incidencias!B:I,8,0),"""")"
    Range("BD10").FormulaLocal = "MARCACION"
    Range("BE10").Formula = "=IFERROR(IF(BE11<>""Ausencia"",VLOOKUP(BE7,Incidencias!B:J,9,0),""""),"""")"
    Range("BD11").FormulaLocal = "TIPO DE FALTA"
    Range("BE11").Formula = "=VLOOKUP(BE7,Incidencias!B:L,11,0)"
    Range("BD12").FormulaLocal = "Fecha"
    Range("BE12").Formula = "=IFERROR(VLOOKUP(BE7,Incidencias!B:G,6,0),"""")"
    Range("BG12").Formula = "=IFERROR(MID(BE12,1,FIND(""("",BE12,1)-1),"""")"
    Range("BH12").Formula = "=IFERROR(MID(BE12,FIND(""("",BE12,1)+1,FIND("")"",BE12,1)-(FIND(""("",BE12,1)+1)),"""")"
    
    Range("BD13").FormulaLocal = "Tipo Incidencia"
    Range("BE13").Formula = "=VLOOKUP(BE11,BD18:BE22,2,0)"
    Range("BD14").FormulaLocal = "TIPO"
    Range("BE14").Formula = "=IF(BE13=2,""INAS"",""TARD"")"
    Range("BG14").Formula = "=IF(AND(BE14=""TARD"",BF8>1,AZ37="""",AZ39="""",BE15=0),""OBS"","""")"
    Range("BH14").Formula = "=IF(AND(BE14=""TARD"",AZ37="""",AZ39<>"""",BE15=0),""OBS"","""")"
    Range("BD15").FormulaLocal = "NRO. AVISOS"
    Range("BE15").Formula = "=IF(BE14=""TARD"",COUNTIF(AZ37:AZ43,""TARDANZAS""),COUNTIF(AZ37:AZ43,""INASISTENCIA""))+BF15"
    Range("BF15").Formula = "=IF(AND(AZ39<>"""",AZ37=""""),1,0)"
    
    Range("BE17").FormulaLocal = "Tipo"
    Range("BF17").Formula = "=IFERROR(VLOOKUP(BE13,BE18:BF22,2,0),"""")"
    Range("BG17").Formula = "=IFERROR(VLOOKUP(BE13,BE18:BG22,3,0),"""")"
    Range("BH17").Formula = "=IFERROR(VLOOKUP(BE13,BE18:BH22,4,0),"""")"
    Range("BI17").Formula = "=IFERROR(VLOOKUP(BE13,BE18:BI22,5,0),"""")"
    Range("BJ17").Formula = "=IFERROR(VLOOKUP(BE13,BE18:BJ22,6,0),"""")"
    Range("BK17").Formula = "=IFERROR(VLOOKUP(BE13,BE18:BK22,7,0),"""")"
    
    Range("BD18").FormulaLocal = "Ent. Atrasada"
    Range("BE18").FormulaLocal = "1"
    Range("BF18").FormulaLocal = "El(La) Asesor(a) llegó tarde :"
    Range("BG18").FormulaLocal = "el Día :"
    Range("BH18").FormulaLocal = "Debiendo iniciar sus labores a las :"
    Range("BI18").FormulaLocal = ",registrandose a las"
    Range("BJ18").FormulaLocal = "perjudicando el Servicio al Cliente del Dpto."
    Range("BK18").FormulaLocal = "Incumpliendo con el Capítulo IV, Art. 11, numeral 2, 3 - Capitulo VIII, Art. 24, 25, 28 - Capitulo IX, Art. 34, 39  Incisos a, e."
    
    Range("BD19").FormulaLocal = "Ausencia"
    Range("BE19").FormulaLocal = "2"
    Range("BF19").FormulaLocal = "Se sanciona al asesor(a) por haber faltado de manera injustificada."
    Range("BG19").FormulaLocal = "Fecha de Registro :"
    Range("BH19").FormulaLocal = "Debiendo iniciar sus labores a las :"
    Range("BI19").FormulaLocal = "y no registrandose,"
    Range("BJ19").FormulaLocal = "perjudicando el Servicio al Cliente del Dpto."
    Range("BK19").FormulaLocal = "Incumpliendo con el Capítulo IV, Art. 11, numeral 2, 3 - Capitulo VIII, Art. 24,28 - Capitulo IX, Art. 34, 39  Incisos a, e."
    
    Range("BD20").FormulaLocal = "Refrigerio Largo"
    Range("BE20").FormulaLocal = "3"
    Range("BF20").FormulaLocal = "El(La) Asesor(a) retornó tarde de su refrigerio :"
    Range("BG20").FormulaLocal = "el Día :"
    Range("BH20").FormulaLocal = "Debiendo gozar un tiempo máximo de:"
    Range("BI20").FormulaLocal = ",tomando un tiempo total:"
    Range("BJ20").FormulaLocal = "perjudicando el Servicio al Cliente del Dpto."
    Range("BK20").FormulaLocal = "Incumpliendo con el Capítulo IV, Art. 11, numeral 2, 3 - Capitulo VIII, Art. 24, 25, 28 - Capitulo IX, Art. 34, 39  Incisos a, e."
    
    Range("BD21").FormulaLocal = "Exc. Tol. Ingreso"
    Range("BE21").FormulaLocal = "4"
    Range("BF21").FormulaLocal = "El asesor(a) ha realizado Exc. Tol. Ingreso, en los días"
    Range("BG21").FormulaLocal = "En la semana :"
    Range("BH21").FormulaLocal = "Considerandose como tardanza"
    Range("BI21").FormulaLocal = "el Exc. Tolerancia,"
    Range("BJ21").FormulaLocal = "lo cual perjudica el Servicio al Cliente del Dpto."
    Range("BK21").FormulaLocal = "Incumpliendo con el Capítulo IV, Art. 11, numeral 2, 3 - Capitulo VIII, Art. 24, 25, 28 - Capitulo IX, Art. 34, 39  Incisos a, e."
    
    Range("BD22").FormulaLocal = "Exc. Tol. Refrigerio"
    Range("BE22").FormulaLocal = "5"
    Range("BF22").FormulaLocal = "El asesor(a) ha realizado Exc. Tol. Refrigerio, en los días"
    Range("BG22").FormulaLocal = "En la semana :"
    Range("BH22").FormulaLocal = "Considerandose como tardanza"
    Range("BI22").FormulaLocal = "el Exc. Tolerancia,"
    Range("BJ22").FormulaLocal = "lo cual perjudica el Servicio al Cliente del Dpto."
    Range("BK22").FormulaLocal = "Incumpliendo con el Capítulo IV, Art. 11, numeral 2, 3 - Capitulo VIII, Art. 24, 25, 28 - Capitulo IX, Art. 34, 39  Incisos a, e."
    
    Range("BF24").Formula = "=IFERROR(IF(AND(BE14=""TARD"",BE15<5),VLOOKUP(BE14,BE25:BF26,2,0),IF(AND(BE14=""INAS"",BE15<5),VLOOKUP(BE14,BE25:BF26,2,0),"""")),"""")"
    Range("BE25").FormulaLocal = "INAS"
    Range("BF25").FormulaLocal = "Se le recuerda que en la siguiente inasistencia registrada se procederá con:"
    Range("BE26").FormulaLocal = "TARD"
    Range("BF26").FormulaLocal = "Se le recuerda que en la siguiente tardanza registrada se procederá con:"
    
    Range("BE29").FormulaLocal = "0"
    Range("BE30").FormulaLocal = "1"
    Range("BE31").FormulaLocal = "2"
    Range("BE32").FormulaLocal = "3"
    Range("BE33").FormulaLocal = "4"
    Range("BF28").Formula = "=IFERROR(IF(AND(BE14=""TARD"",BE15<5),VLOOKUP(BE15,BE29:BF33,2,0),IF(AND(BE14=""INAS"",BE15<5),VLOOKUP(BE15,BG29:BH33,2,0),"""")),"""")"
    Range("BF29").Formula = "=IF(OR(BG14<>"""",BH14<>""""),""01 Aviso de Desempeño Grave."",""01 Aviso de Desempeño Escrito Simple."")"
    Range("BF30").Formula = "=IF(OR(BG14<>"""",BH14<>""""),""01 Dia de Suspensión."",""01 Aviso de Desempeño Grave."")"
    Range("BF31").Formula = "=IF(OR(BG14<>"""",BH14<>""""),""03 Dias de Suspensión."",""01 Dia de Suspensión."")"
    Range("BF32").Formula = "=IF(OR(BG14<>"""",BH14<>""""),""El proceso de despido por falta grave."",""03 Dias de Suspensión."")"
    Range("BF33").Formula = "=IF(OR(BG14<>"""",BH14<>""""),"""",""El proceso de despido por falta grave."")"
    
    Range("BG29").FormulaLocal = "0"
    Range("BG30").FormulaLocal = "2"
    Range("BG31").FormulaLocal = "2"
    Range("BG32").FormulaLocal = "3"
    Range("BG33").FormulaLocal = "4"
    Range("BH29").FormulaLocal = "01 Aviso de Desempeño Grave."
    Range("BH30").FormulaLocal = "01 Día de Suspensión."
    Range("BH31").FormulaLocal = "01 Día de Suspensión."
    Range("BH32").FormulaLocal = "03 Días de Suspensión."
    Range("BH33").FormulaLocal = "El proceso de despido por falta grave."
    
    '---- 2DA PARTE de la columna AP a BB ----
    Sheets("IMPRESION").Select
    Range("AP4").FormulaLocal = "NOMBRE Y APELLIDOS"
    Range("AP5").FormulaLocal = "PUNTO DE VENTA"
    Range("AP6").FormulaLocal = "AREA DE TRABAJO"
    Range("AS4").Formula = "=IFERROR(VLOOKUP(BE7,Incidencias!B:C,2,0),"""")"
    Range("AS5").Formula = "=IFERROR(IF(AS4<>"""",MID(Incidencias!A6,FIND(""-"",Incidencias!A6,1)+1,LEN(Incidencias!A6)-FIND(""-"",Incidencias!A6,1)),""""),"""")"
    Range("AS6").Formula = "=IFERROR(VLOOKUP(BE7,Incidencias!B:G,5,0),"""")"
    
    Range("AW6").FormulaLocal = "FT"
    Range("AX6").Formula = "=IFERROR(IF(VLOOKUP(BE7,Incidencias!B:D,3,0)=""FT"",""X"","" ""),"""")"
    Range("AY6").FormulaLocal = "PT"
    Range("AZ6").Formula = "=IFERROR(IF(VLOOKUP(BE7,Incidencias!B:D,3,0)=""P23"",""X"","" ""),"""")"
    Range("BA6").FormulaLocal = "PK"
    Range("BB6").Formula = "=IFERROR(IF(VLOOKUP(BE7,Incidencias!B:D,3,0)=""P19:45"",""X"","" ""),"""")"
    
    Range("AP10").Formula = "=IFERROR(IF(AND(BE14=""TARD"",BF8<=1,BE15=0),""X"",""""),"""")"
    Range("AQ10").FormulaLocal = "VERBAL"
    Range("AP14").Formula = "=IFERROR(IF(AND(OR(BE14=""TARD"",BE14=""INAS""),AP10<>""X""),""X"",""""),"""")"
    Range("AQ14").FormulaLocal = "ESCRITO"
    Range("AQ16").Formula = "=IFERROR(IF(BE14=""TARD"",IF(AND(AP14=""X"",BE15=1,BF8>=0),""X"",IF(AND(AP14=""X"",BE15=0,BF8>1),""X"","""")),IF(BE14=""INAS"",IF(AND(AP14=""X"",BE15=0),""X"",""""),"""")),"""")"
    Range("AR16").FormulaLocal = "SIMPLE"
    Range("AQ18").Formula = "=IFERROR(IF(BE14=""TARD"",IF(AND(AP14=""X"",BE15>=2),""X"",""""),IF(BE14=""INAS"",IF(AND(AP14=""X"",BE15>=1),""X"",""""),"""")),"""")"
    Range("AR18").FormulaLocal = "GRAVE"
    Range("AT14").Formula = "=IFERROR(IF(BE11=""Ausencia"",""X"","" ""),"""")"
    Range("AU14").FormulaLocal = "INASISTENCIA"
    Range("AT16").Formula = "=IFERROR(IF(BE11<>""Ausencia"",""X"","" ""),"""")"
    Range("AU16").FormulaLocal = "TARDANZA"
    
    Range("AP37").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR45<>""""),""X"",IF(AND($BE$14=""INAS"",AX45<>""""),""X"","""")),"""")"
    Range("AQ37").FormulaLocal = "VERBAL"
    Range("AT37").FormulaLocal = "FECHA"
    Range("AU37").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR45<>""""),AR45,IF(AND($BE$14=""INAS"",AX45<>""""),AX45,"""")),"""")"
    Range("AX37").FormulaLocal = "MOTIVO"
    Range("AZ37").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AT45<>""""),AT45,IF(AND($BE$14=""INAS"",AZ45<>""""),AZ45,"""")),"""")"

    Range("AP39").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR46<>""""),""X"",IF(AND($BE$14=""INAS"",AX46<>""""),""X"","""")),"""")"
    Range("AQ39").FormulaLocal = "ESCRITO SIMPLE"
    Range("AT39").FormulaLocal = "FECHA"
    Range("AU39").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR46<>""""),AR46,IF(AND($BE$14=""INAS"",AX46<>""""),AX46,"""")),"""")"
    Range("AX39").FormulaLocal = "MOTIVO"
    Range("AZ39").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AT46<>""""),AT46,IF(AND($BE$14=""INAS"",AZ46<>""""),AZ46,"""")),"""")"
    
    Range("AP41").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR47<>""""),""X"",IF(AND($BE$14=""INAS"",AX47<>""""),""X"","""")),"""")"
    Range("AQ41").FormulaLocal = "ESCRITO GRAVE"
    Range("AT41").FormulaLocal = "FECHA"
    Range("AU41").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR47<>""""),AR47,IF(AND($BE$14=""INAS"",AX47<>""""),AX47,"""")),"""")"
    Range("AX41").FormulaLocal = "MOTIVO"
    Range("AZ41").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AT47<>""""),AT47,IF(AND($BE$14=""INAS"",AZ47<>""""),AZ47,"""")),"""")"
    
    Range("AP43").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR48<>""""),""X"",IF(AND($BE$14=""INAS"",AX48<>""""),""X"","""")),"""")"
    Range("AQ43").FormulaLocal = "SUSPENSION"
    Range("AT43").FormulaLocal = "FECHA"
    Range("AU43").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AR48<>""""),AR48,IF(AND($BE$14=""INAS"",AX48<>""""),AX48,"""")),"""")"
    Range("AX43").FormulaLocal = "MOTIVO"
    Range("AZ43").Formula = "=IFERROR(IF(AND($BE$14=""TARD"",AT48<>""""),AT48,IF(AND($BE$14=""INAS"",AZ48<>""""),AZ48,"""")),"""")"
    
    Range("AP45").Formula = "=CONCATENATE(BE7,""TARD"",""A"")"
    Range("AR45").Formula = "=IFERROR(VLOOKUP(AP45,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AT45").Formula = "=IFERROR(VLOOKUP(AP45,'Control Disciplinario'!F:I,4,0),"""")"
    Range("AV45").Formula = "=CONCATENATE(BE7,""INAS"",""A"")"
    Range("AX45").Formula = "=IFERROR(VLOOKUP(AV45,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AZ45").Formula = "=IFERROR(VLOOKUP(AV45,'Control Disciplinario'!F:I,4,0),"""")"
    
    Range("AP46").Formula = "=CONCATENATE(BE7,""TARD"",""B"")"
    Range("AR46").Formula = "=IFERROR(VLOOKUP(AP46,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AT46").Formula = "=IFERROR(VLOOKUP(AP46,'Control Disciplinario'!F:I,4,0),"""")"
    Range("AV46").Formula = "=CONCATENATE(BE7,""INAS"",""B"")"
    Range("AX46").Formula = "=IFERROR(VLOOKUP(AV46,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AZ46").Formula = "=IFERROR(VLOOKUP(AV46,'Control Disciplinario'!F:I,4,0),"""")"
    
    Range("AP47").Formula = "=CONCATENATE(BE7,""TARD"",""C"")"
    Range("AR47").Formula = "=IFERROR(VLOOKUP(AP47,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AT47").Formula = "=IFERROR(VLOOKUP(AP47,'Control Disciplinario'!F:I,4,0),"""")"
    Range("AV47").Formula = "=CONCATENATE(BE7,""INAS"",""C"")"
    Range("AX47").Formula = "=IFERROR(VLOOKUP(AV47,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AZ47").Formula = "=IFERROR(VLOOKUP(AV47,'Control Disciplinario'!F:I,4,0),"""")"
    
    Range("AP48").Formula = "=CONCATENATE(BE7,""TARD"",""D"")"
    Range("AR48").Formula = "=IFERROR(VLOOKUP(AP48,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AT48").Formula = "=IFERROR(VLOOKUP(AP48,'Control Disciplinario'!F:I,4,0),"""")"
    Range("AV48").Formula = "=CONCATENATE(BE7,""INAS"",""D"")"
    Range("AX48").Formula = "=IFERROR(VLOOKUP(AV48,'Control Disciplinario'!F:J,5,0),"""")"
    Range("AZ48").Formula = "=IFERROR(VLOOKUP(AV48,'Control Disciplinario'!F:I,4,0),"""")"
    
    '---- 3RA PARTE de la columna B a AI ----
    Range("AB2").Formula = "=LEFT(IF(DAY(BE6)<10,CONCATENATE(""0"",DAY(BE6)),DAY(BE6)),1)"
    Range("AC2").Formula = "=RIGHT(IF(DAY(BE6)<10,CONCATENATE(""0"",DAY(BE6)),DAY(BE6)),1)"
    Range("AD2").Formula = "=LEFT(IF(MONTH(BE6)<10,CONCATENATE(""0"",MONTH(BE6)),MONTH(BE6)),1)"
    Range("AE2").Formula = "=RIGHT(IF(MONTH(BE6)<10,CONCATENATE(""0"",MONTH(BE6)),MONTH(BE6)),1)"
    Range("AF2").Formula = "2"
    Range("AG2").Formula = "0"
    Range("AH2").Formula = "=LEFT(MID(YEAR(BE6),3,2))"
    Range("AI2").Formula = "=RIGHT(MID(YEAR(BE6),3,2))"
    
    Range("I4").Formula = "=AS4"
    Range("I5").Formula = "=AS5"
    Range("I6").Formula = "=AS6"
    Range("AA6").Formula = "=AX6"
    Range("AE6").Formula = "=AZ6"
    Range("AI6").Formula = "=BB6"
    
    Range("C10").Formula = "=AP10"
    Range("C14").Formula = "=AP14"
    Range("E16").Formula = "=AQ16"
    Range("E18").Formula = "=AQ18"
    Range("J14").Formula = "=AT14"
    Range("J16").Formula = "=AT16"
    
    Range("B25").Formula = "=BF17"
    Range("R25").Formula = "=IFERROR(IF(BE13>3,BG12,BE8),"""")"
    Range("U25").Formula = "=BG17"
    Range("AB25").Formula = "=IFERROR(IF(BE13>3,BH12,BE12),"""")"
    Range("B26").Formula = "=BH17"
    Range("J26").Formula = "=IFERROR(IF(BE13>3,"""",BE9),"""")"
    Range("L26").Formula = "=BI17"
    Range("R26").Formula = "=IFERROR(IF(BE13>3,"""",BE10),"""")"
    Range("U26").Formula = "=BJ17"
    Range("B27").Formula = "=BF24"
    Range("B28").Formula = "=BF28"
    Range("B29").Formula = "=BK17"
    
    Range("B31").FormulaLocal = "Se ha brindado la retroalimentación respectiva acerca de la importancia del cumplimiento de la malla horaria con el fin de mantener un óptimo nivel de"
    Range("B32").FormulaLocal = "servicio en la tienda. Se continuará haciendo el seguimiento respectivo a las marcaciones del asesor."
    
    Range("C37").Formula = "=AP37"
    Range("C39").Formula = "=AP39"
    Range("C41").Formula = "=AP41"
    Range("C43").Formula = "=AP43"
    Range("N37").Formula = "=AU37"
    Range("N39").Formula = "=AU39"
    Range("N41").Formula = "=AU41"
    Range("N43").Formula = "=AU43"
    Range("Y37").Formula = "=AZ37"
    Range("Y39").Formula = "=AZ39"
    Range("Y41").Formula = "=AZ41"
    Range("Y43").Formula = "=AZ43"
    
    Range("J53").Formula = "=BE6"
    Range("AE53").Formula = "=BE6"
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
End Sub
Private Sub Borra_Formulas_Impresion()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 25.01.2018
    Sheets("IMPRESION").Select
    Range("I4:W4,I5:W5,I6:W6,C10:D10,C14:D14,J14:K14,J16:K16,J18:K18").UnMerge
    Range("B31:AI31,B32:AI32,C37:D37,C39:D39,C41:D41,C43:D43,N37:S37,N39:S39,N41:S41,N43:S43,Y37:AH37,Y39:AH39,Y41:AH41,Y43:AH43,J53:N53,AE53:AH53").UnMerge
    Range("AB2:AI2,I4:W6,AA6,AE6,AI6,C10:D10,C14:D14,E16:E16,E18:E18,J14:K14,J16:K16,B25:AI29,B31,B32").ClearContents
    Range("C37:D37,C39:D39,C41:D41,C43:D43,N37:S37,N39:S39,N41:S41,N43:S43,Y37:AH37,Y39:AH39,Y41:AH41,Y43:AH43,J53:N53,AE53:AH53").ClearContents
    Range("AP4:BK50").ClearContents
    Range("I4:W4,I5:W5,I6:W6,C10:D10,C14:D14,J14:K14,J16:K16,J18:K18").Merge
    Range("B31:AI31,B32:AI32,C37:D37,C39:D39,C41:D41,C43:D43,N37:S37,N39:S39,N41:S41,N43:S43,Y37:AH37,Y39:AH39,Y41:AH41,Y43:AH43,J53:N53,AE53:AH53").Merge
    Range("A1").Select
End Sub
