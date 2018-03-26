VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Sheets("PareoMarcajes").Visible = xlSheetVisible
    Sheets("PareoMarcajes").Select
    If (Range("L1").Value = "BendicemeDios") Then
        'No borro formulas
    Else
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        
        Call Borra_Datos_Impresion
        'Call Borra_02Hojas
        
        'Posicion al abrir el archivo
        Sheets("Incidencias").Select
        'Guarda los cambios del archivo
        ActiveWorkbook.Save
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
    End If
End Sub
Private Sub Borra_Datos_Impresion()
    Sheets("IMPRESION").Visible = xlSheetVisible
    Sheets("IMPRESION").Select
    Range("I4:W4,I5:W5,I6:W6,C10:D10,C14:D14,J14:K14,J16:K16,J18:K18").UnMerge
    Range("C37:D37,C39:D39,C41:D41,C43:D43,N37:S37,N39:S39,N41:S41,N43:S43,Y37:AH37,Y39:AH39,Y41:AH41,Y43:AH43,J53:N53,AE53:AH53").UnMerge
    Range("AB2:AI2,I4:W6,AA6,AE6,AI6,C10:D10,C14:D14,E16:E16,E18:E18,J14:K14,J16:K16,B25:AI29").ClearContents
    Range("C37:D37,C39:D39,C41:D41,C43:D43,N37:S37,N39:S39,N41:S41,N43:S43,Y37:AH37,Y39:AH39,Y41:AH41,Y43:AH43,J53:N53,AE53:AH53").ClearContents
    Range("AP4:BK50").ClearContents
    Range("I4:W4,I5:W5,I6:W6,C10:D10,C14:D14,J14:K14,J16:K16,J18:K18").Merge
    Range("C37:D37,C39:D39,C41:D41,C43:D43,N37:S37,N39:S39,N41:S41,N43:S43,Y37:AH37,Y39:AH39,Y41:AH41,Y43:AH43,J53:N53,AE53:AH53").Merge
    Range("I4").Select
    Sheets("IMPRESION").Visible = xlSheetVeryHidden
End Sub
Private Sub Borra_02Hojas()
    On Error Resume Next
        
    For i = 1 To Worksheets.Count
        Sheets("Control Disciplinario").Visible = xlSheetVisible
        If (Worksheets(i).Name = "Control Disciplinario") Then
            Sheets("Control Disciplinario").Delete
            Exit For
        End If
    Next i
    For j = 1 To Worksheets.Count
        Sheets("Dotacion Ofisis").Visible = xlSheetVisible
        If (Worksheets(j).Name = "Dotacion Ofisis") Then
            Sheets("Dotacion Ofisis").Delete
            Exit For
        End If
    Next j
End Sub
