VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "0%"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    End
    Unload Me
End Sub
Private Sub UserForm_Activate()
    Dim Conteo, nFilas, nColumnas, f, c As Long
    Dim Porcentaje As Double
    'Cells.Clear
    Conteo = 1
    nFilas = 1000
    nColumnas = 500
    For f = 1 To nFilas
        For c = 1 To nColumnas
            'Cells(f, c) = Conteo
            Conteo = Conteo + 1
        Next c
            Porcentaje = Conteo / (nFilas * nColumnas)
            Me.Caption = Format(Porcentaje, "0%")
            Me.Label1.Width = Porcentaje * Me.Frame1.Width
            DoEvents
    Next f
    Unload Me
End Sub
