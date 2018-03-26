Attribute VB_Name = "FInformacionFormulas"
Sub Info_Tolerancia()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("PareoMarcajes").Select
    Range("AB12").End(xlDown).Select
    Dim NroFi, NroCo As Integer
    NroFi = ActiveCell.Row
    NroCo = ActiveCell.Column
    With Range("AJ11:AM11")
        .ClearContents
        .ColumnWidth = 8.43
        .NumberFormat = "General"
    End With
    Application.Calculation = xlCalculationManual
    Range("AJ11").FormulaLocal = "COD_1"
    Range("AJ12").Formula = "=CONCATENATE(B12,H12)"
    Range("AK11").FormulaLocal = "VALOR_1"
    Range("AK12").Formula = "=IF(AND(MID(K12,1,1)<>""-"",L12<>""Ausencia"",K12<>""00:00""),ROUND(K12*24,2),"""")"
    Range("AL11").FormulaLocal = "COD_2"
    Range("AL12").Formula = "=CONCATENATE(B12,H12)"
    Range("AM11").FormulaLocal = "VALOR_2"
    Range("AM12").Formula = "=IF((ROUND(Q12*24,2))-1<=0,"""",ROUND(Q12*24,2)-1)"
    Range("AJ12:AM12").Copy Range(Cells(13, 36), Cells(NroFi, NroCo + 11))
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
    Sheets("Dotacion Ofisis").Select
    Cells.EntireColumn.AutoFit
    Range("P1").End(xlDown).Select
    Dim NroFila, NroColumna As Integer
    NroFila = ActiveCell.Row
    NroColumna = ActiveCell.Column
    With Range("Q:AW")
        .ClearContents
        .ColumnWidth = 6
        .NumberFormat = "General"
    End With
    Call Formulas_Dotacion(NroFila, NroColumna)
    Call Formulas_A_Valores(NroFila, NroColumna)
End Sub
Private Sub Formulas_Dotacion(ByVal f As Integer, ByVal c As Integer)
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Application.Calculation = xlCalculationManual
    Range("Q1").FormulaLocal = "EAN"
    Range("Q2").Formula = "=IFERROR(VLOOKUP(R2,PareoMarcajes!B:AL,37,0),"""")"
    
    Range("R1").FormulaLocal = "DNI"
    Range("R2").Formula = "=IFERROR(MID(M2,7,8),""-"")"
    
    Range("S1").FormulaLocal = "TRABAJADOR"
    Range("S2").Formula = "=E2"
    
    Range("T1").FormulaLocal = "APELLIDOS_NOMBRES"
    Range("T2").Formula = "=F2"
    
    Range("U1").FormulaLocal = "PLANILLA"
    Range("U2").Formula = "=G2"
    
    Range("V1").FormulaLocal = "DESCRIPCION"
    Range("V2").Formula = "=J2"
    
    Range("W1").FormulaLocal = "TIPO"
    Range("W2").Formula = "=IFERROR(VLOOKUP(R2,PareoMarcajes!B:D,3,0),"""")"
    
    Range("X1").FormulaLocal = "=PareoMarcajes!$H$12"
    Range("X2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$X$1),PareoMarcajes!$AJ:$AK,2,0),"""")"
    
    Range("Y1").FormulaLocal = "=PareoMarcajes!$H$13"
    Range("Y2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$Y$1),PareoMarcajes!$AJ:$AK,2,0),"""")"
    
    Range("Z1").FormulaLocal = "=PareoMarcajes!$H$14"
    Range("Z2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$Z$1),PareoMarcajes!$AJ:$AK,2,0),"""")"
    
    Range("AA1").FormulaLocal = "=PareoMarcajes!$H$15"
    Range("AA2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AA$1),PareoMarcajes!$AJ:$AK,2,0),"""")"
    
    Range("AB1").FormulaLocal = "=PareoMarcajes!$H$16"
    Range("AB2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AB$1),PareoMarcajes!$AJ:$AK,2,0),"""")"
    
    Range("AC1").FormulaLocal = "=PareoMarcajes!$H$17"
    Range("AC2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AC$1),PareoMarcajes!$AJ:$AK,2,0),"""")"
    
    Range("AD1").FormulaLocal = "=PareoMarcajes!$H$18"
    Range("AD2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AD$1),PareoMarcajes!$AJ:$AK,2,0),"""")"
    
    Range("AE1").FormulaLocal = "CON_FT/PT"
    Range("AE2").Formula = "=IF(W2<>""P19:45"",COUNT(X2:AD2),"""")"
    
    Range("AF1").FormulaLocal = "DIAS"
    Range("AF2").Formula = "=IF(AND(AE2>=3,AE2<>""""),CONCATENATE(IF(X2<>"""",MID($X$1,1,2),"" ""),IF(X2<>"""","" ,"",""""),IF(Y2<>"""",MID($Y$1,1,2),""""),IF(Y2<>"""","" ,"",""""),IF(Z2<>"""",MID($Z$1,1,2),""""),IF(Z2<>"""","" ,"",""""),IF(AA2<>"""",MID($AA$1,1,2),""""),IF(AA2<>"""","" ,"",""""),IF(AB2<>"""",MID($AB$1,1,2),""""),IF(AB2<>"""","" ,"",""""),IF(AC2<>"""",MID($AC$1,1,2),""""),IF(AC2<>"""","" ,"",""""),IF(AD2<>"""",MID($AD$1,1,2),""""),IF(AD2<>"""","" ,"","""")),"""")"
    
    Range("AG1").FormulaLocal = "CON_PK"
    Range("AG2").Formula = "=IF(W2=""P19:45"",COUNT(AB2:AD2),"""")"
    
    Range("AH1").FormulaLocal = "DIAS_PK"
    Range("AH2").Formula = "=IF(AND(AG2>=2,AG2<>""""),CONCATENATE(IF(AB2<>"""",MID($AB$1,1,2),""""),IF(AB2<>"""","" ,"",""""),IF(AC2<>"""",MID($AC$1,1,2),""""),IF(AC2<>"""","" ,"",""""),IF(AD2<>"""",MID($AD$1,1,2),""""),IF(AD2<>"""","" ,"","""")),"""")"
    
    Range("AI1").FormulaLocal = "OBS"
    Range("AI2").Formula = "=IF(AND(AF2<>"""",AH2=""""),AF2,IF(AND(AH2<>"""",AF2=""""),AH2,""""))"
    
    Range("AJ1").FormulaLocal = "OBS_FINAL"
    Range("AJ2").Formula = "=IF(AI2<>"""",CONCATENATE(TRIM(MID(AI2,1,LEN(AI2)-1)),"" (Sem "",MID($X$1,1,2),"" al "",$AD$1,"")""),"""")"
    
    Range("AK1").FormulaLocal = "=PareoMarcajes!$H$12"
    Range("AK2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AK$1),PareoMarcajes!$AL:$AM,2,0),"""")"
    
    Range("AL1").FormulaLocal = "=PareoMarcajes!$H$13"
    Range("AL2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AL$1),PareoMarcajes!$AL:$AM,2,0),"""")"
    
    Range("AM1").FormulaLocal = "=PareoMarcajes!$H$14"
    Range("AM2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AM$1),PareoMarcajes!$AL:$AM,2,0),"""")"
    
    Range("AN1").FormulaLocal = "=PareoMarcajes!$H$15"
    Range("AN2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AN$1),PareoMarcajes!$AL:$AM,2,0),"""")"
    
    Range("AO1").FormulaLocal = "=PareoMarcajes!$H$16"
    Range("AO2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AO$1),PareoMarcajes!$AL:$AM,2,0),"""")"
    
    Range("AP1").FormulaLocal = "=PareoMarcajes!$H$17"
    Range("AP2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AP$1),PareoMarcajes!$AL:$AM,2,0),"""")"
    
    Range("AQ1").FormulaLocal = "=PareoMarcajes!$H$18"
    Range("AQ2").Formula = "=IFERROR(VLOOKUP(CONCATENATE(R2,$AQ$1),PareoMarcajes!$AL:$AM,2,0),"""")"
    
    Range("AR1").FormulaLocal = "CON_FT/PT"
    Range("AR2").Formula = "=IF(W2<>""P19:45"",COUNT(AK2:AQ2),"""")"
    
    Range("AS1").FormulaLocal = "DIAS"
    Range("AS2").Formula = "=IF(AND(AR2>=3,AR2<>""""),CONCATENATE(IF(AK2<>"""",MID($AK$1,1,2),"" ""),IF(AK2<>"""","" ,"",""""),IF(AL2<>"""",MID($AL$1,1,2),""""),IF(AL2<>"""","" ,"",""""),IF(AM2<>"""",MID($AM$1,1,2),""""),IF(AM2<>"""","" ,"",""""),IF(AN2<>"""",MID($AN$1,1,2),""""),IF(AN2<>"""","" ,"",""""),IF(AO2<>"""",MID($AO$1,1,2),""""),IF(AO2<>"""","" ,"",""""),IF(AP2<>"""",MID($AP$1,1,2),""""),IF(AP2<>"""","" ,"",""""),IF(AQ2<>"""",MID($AQ$1,1,2),""""),IF(AQ2<>"""","" ,"","""")),"""")"
    
    Range("AT1").FormulaLocal = "CON_PK"
    Range("AT2").Formula = "=IF(W2=""P19:45"",COUNT(AO2:AQ2),"""")"
    
    Range("AU1").FormulaLocal = "DIAS_PK"
    Range("AU2").Formula = "=IF(AND(AT2>=2,AT2<>""""),CONCATENATE(IF(AO2<>"""",MID($AO$1,1,2),""""),IF(AO2<>"""","" ,"",""""),IF(AP2<>"""",MID($AP$1,1,2),""""),IF(AP2<>"""","" ,"",""""),IF(AQ2<>"""",MID($AQ$1,1,2),""""),IF(AQ2<>"""","" ,"","""")),"""")"
    
    Range("AV1").FormulaLocal = "OBS"
    Range("AV2").Formula = "=IF(AND(AS2<>"""",AU2=""""),AS2,IF(AND(AU2<>"""",AS2=""""),AU2,""""))"
    
    Range("AW1").FormulaLocal = "OBS_FINAL"
    Range("AW2").Formula = "=IF(AV2<>"""",CONCATENATE(TRIM(MID(AV2,1,LEN(AV2)-1)),"" (Sem "",MID($AK$1,1,2),"" al "",$AQ$1,"")""),"""")"
    
    Range("Q2:AW2").Copy Range(Cells(3, 17), Cells(f, c + 33))
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
 End Sub
 Private Sub Formulas_A_Valores(ByVal f As Integer, ByVal c As Integer)
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    'Dato para REVISION
    Range("AY1").FormulaLocal = "=PareoMarcajes!L1"
    Dim dato As String
    dato = Range("AY1").Value
    
    If (dato = "BendicemeDios") Then
        'Se visualizan las fórmulas
    Else
        Range(Cells(1, 17), Cells(f, c + 33)).Copy
        Range(Cells(1, 17), Cells(f, c + 33)).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    Range("Q1").Select
 End Sub
Sub Info_Incidencia()
    'Dedicado a mi Señor Todopoderoso
    'Creado por: Hugo Garcia Silva V3.0
    'Ultima actualizacion 04.07.2017
    Sheets("Incidencias").Select
    Range("L11").End(xlDown).Select
    Dim NroFi, NroCo As Integer
    NroFi = ActiveCell.Row
    NroCo = ActiveCell.Column
    'Borro datos y agrego formato
    With Range("M:AD")
        .ClearContents
        .NumberFormat = "General"
    End With
    Range("M:AA").ColumnWidth = 6
    Range("AB:AB").ColumnWidth = 15
    Range("AC:AD").ColumnWidth = 17
    'Agrego formulas
    Range("N1").FormulaLocal = "Tipo"
    Range("N2").FormulaLocal = "Ent. Atrasada"
    Range("N3").FormulaLocal = "Ausencia"
    Range("N4").FormulaLocal = "Refrigerio Largo"
    Range("N5").FormulaLocal = "Exc. Tol. Ingreso"
    Range("N6").FormulaLocal = "Exc. Tol. Refrigerio"
    Range("O2").FormulaLocal = "1"
    Range("O3").FormulaLocal = "2"
    Range("O4").FormulaLocal = "3"
    Range("O5").FormulaLocal = "4"
    Range("O6").FormulaLocal = "5"
    
    Range("Q1").FormulaLocal = "Tardanzas"
    Range("Q2").FormulaLocal = "0"
    Range("Q3").FormulaLocal = "1"
    Range("Q4").FormulaLocal = "2"
    Range("Q5").FormulaLocal = "3"
    Range("Q6").FormulaLocal = "4"
    Range("Q7").FormulaLocal = "5"
    Range("R2").FormulaLocal = "Verbal"
    Range("R3").FormulaLocal = "Escrito Simple"
    Range("R4").FormulaLocal = "Escrito Grave"
    Range("R5").FormulaLocal = "Escrito Grave 01 Día Susp."
    Range("R6").FormulaLocal = "Escrito Grave 03 Días Susp."
    Range("R7").FormulaLocal = "Proceso de despido."
    
    Range("V1").FormulaLocal = "Inasistencias"
    Range("V2").FormulaLocal = "1"
    Range("V3").FormulaLocal = "2"
    Range("V4").FormulaLocal = "3"
    Range("V5").FormulaLocal = "4"
    Range("V6").FormulaLocal = "5"
    Range("W2").FormulaLocal = "Escrito Simple"
    Range("W3").FormulaLocal = "Escrito Grave"
    Range("W4").FormulaLocal = "Escrito Grave 01 Día Susp."
    Range("W5").FormulaLocal = "Escrito Grave 03 Días Susp."
    Range("W6").FormulaLocal = "Proceso de despido."
    
    Application.Calculation = xlCalculationManual
    Range("M10").FormulaLocal = "Tipo Incidencia"
    Range("M11").Formula = "=VLOOKUP($L11,$N$2:$O$6,2,0)"
    
    Range("N10").FormulaLocal = "TIPO"
    Range("N11").Formula = "=IF($M11=2,""INAS"",""TARD"")"
    
    Range("O10").FormulaLocal = "TARDA"
    Range("O11").Formula = "=IFERROR(IF($N11=""TARD"",VLOOKUP(CONCATENATE($B11,$O$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("P10").FormulaLocal = "TARDB"
    Range("P11").Formula = "=IFERROR(IF($N11=""TARD"",VLOOKUP(CONCATENATE($B11,$P$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("Q10").FormulaLocal = "TARDC"
    Range("Q11").Formula = "=IFERROR(IF($N11=""TARD"",VLOOKUP(CONCATENATE($B11,$Q$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("R10").FormulaLocal = "TARDD"
    Range("R11").Formula = "=IFERROR(IF($N11=""TARD"",VLOOKUP(CONCATENATE($B11,$R$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("S10").FormulaLocal = "TARDE"
    Range("S11").Formula = "=IFERROR(IF($N11=""TARD"",VLOOKUP(CONCATENATE($B11,$S$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("T10").FormulaLocal = "INASA"
    Range("T11").Formula = "=IFERROR(IF($N11=""INAS"",VLOOKUP(CONCATENATE($B11,$T$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("U10").FormulaLocal = "INASB"
    Range("U11").Formula = "=IFERROR(IF($N11=""INAS"",VLOOKUP(CONCATENATE($B11,$U$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("V10").FormulaLocal = "INASC"
    Range("V11").Formula = "=IFERROR(IF($N11=""INAS"",VLOOKUP(CONCATENATE($B11,$V$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("W10").FormulaLocal = "INASD"
    Range("W11").Formula = "=IFERROR(IF($N11=""INAS"",VLOOKUP(CONCATENATE($B11,$W$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("X10").FormulaLocal = "INASE"
    Range("X11").Formula = "=IFERROR(IF($N11=""INAS"",VLOOKUP(CONCATENATE($B11,$X$10),'Control Disciplinario'!F:I,4,0),""""),"""")"
    
    Range("Y10").FormulaLocal = "SUMA 1"
    Range("Y11").Formula = "=IF(OR(AND(P11<>"""",O11=""""),AND(U11<>"""",T11=""""),AND(N11=""TARD"",ROUND(K11*24,2)>1,Z11=0),AND(N11=""INAS"",Z11=0)),1,0)"
    
    Range("Z10").FormulaLocal = "TIENE"
    Range("Z11").Formula = "=IF(N11=""TARD"",COUNTIF(O11:S11,""TARDANZAS""),COUNTIF(T11:X11,""INASISTENCIA""))"
    
    Range("AA10").FormulaLocal = "AVISO"
    Range("AA11").Formula = "=Z11+Y11"
    
    Range("AB10").FormulaLocal = "Tipo de sanción"
    Range("AB11").Formula = "=IFERROR(IF(AND($N11=""TARD"",$AA11<=6),VLOOKUP($AA11,$Q$2:$R$7,2,0),IF(AND($N11=""INAS"",$AA11<=5),VLOOKUP($AA11,$V$2:$W$6,2,0),"""")),"""")"
    
    Range("M11:AB11").Copy Range(Cells(11, 13), Cells(NroFi, NroCo + 16))
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
    Range("AY1").FormulaLocal = "=PareoMarcajes!L1"
    Dim dato As String
    dato = Range("AY1").Value
    
    If (dato = "BendicemeDios") Then
        'Se visualizan las fórmulas
    Else
        Range(Cells(11, 13), Cells(NroFi, NroCo + 16)).Copy
        Range(Cells(11, 13), Cells(NroFi, NroCo + 16)).PasteSpecial Paste:=xlPasteValues
        
        Range(Cells(10, 6), Cells(NroFi, NroCo - 6)).Copy
        Range(Cells(10, 28), Cells(NroFi, NroCo + 16)).PasteSpecial Paste:=xlPasteFormats
        
        Application.CutCopyMode = False
        
        Range("M:AA").EntireColumn.Delete
        Columns("M:M").AutoFit
    End If
    Range("A11").Select
    Range("L11").Select
 End Sub
