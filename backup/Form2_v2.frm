VERSION 5.00
Begin VB.Form Form2_v2 
   Caption         =   "Elaborar Tarjetas de Inventario"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Liso/Herreria"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Estriado"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1080
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Text            =   "1"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Text            =   "1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2_v2.frx":0000
      Left            =   2880
      List            =   "Form2_v2.frx":002E
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Boton_Cancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      Picture         =   "Form2_v2.frx":0073
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Boton_Ok 
      Caption         =   "Conforme"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Enumerar desde el Atado No?.:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Titulo 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad de Tarjetas?:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha?:"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Producto?:"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form2_v2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-------------------------------------------------------------
'*  MODULO: Form2 ( nombre logico )
'*  PROPOSITO: Captar datos claves para Impresion de Tarjetas
'*             de Inventario de Productos
'*  NOMBRE FISICO: Form2.frm
'*  Elaborado x Henry J. Pulgar B.
'*  Creado el 21 de Agosto de 2002.
'*  Actualizado el 21 de Agosto de 2002.
'*--------------------------------------------------------------

Private Sub Boton_Cancelar_Click()
    Unload Form2_v2
End Sub

Private Sub Boton_Ok_Click()
   iNVTARJ01v2_FRM.Show
 End Sub

Private Sub Combo1_LostFocus()
   Dim Mensaje
   '*-----------------------------------------
   '*If Combo1.Text = "AT" Then
   '*   Option1.Caption = "Estriado"
   '*   Option2.Caption = "Liso/Herreria"
   '*Else
   '*    Option1.Caption = ""
   '*    Option2.Caption = ""
   '*End If
   '*-----------------------------------------
   Combo1.Text = UCase(Combo1.Text)
   Select Case Combo1.Text
          Case "AT", "AH", "CE", "MP", "MR", "MR*", "MRA", "MRB", "MRC", "MRD", "MRE", "MRF", "MRG", "MRH"
              'OK!
              If (Combo1.Text = "MP" Or Combo1.Text = "MR*") Then
                  Text3.Enabled = True
                  Text3.BackColor = &H80000005
              Else
                  Text3.Enabled = False
                  Text3.BackColor = &HC0FFC0
              End If
          Case Else
               Mensaje = "Tipo de producto no definido."
               MsgBox Mensaje, vbCritical, "Atencion!"
               Combo1.SetFocus
   End Select
End Sub

Private Sub Text1_LostFocus()
  Dim Mensaje
  If Not IsDate(Text1.Text) Then
     Beep
     Text1.BackColor = &H8080FF
     Mensaje = "Fecha Invalida."
     MsgBox Mensaje, vbCritical, "Atencion!"
     Text1.SetFocus
  Else
     Text1.Text = Format(Text1.Text, "DD/MM/YYYY")
  End If
  Text1.BackColor = &HFFFFFF
End Sub

Private Sub Text2_LostFocus()
    Dim Mensaje
    If Not IsNumeric(Text2.Text) Then
       Beep
       Text2.BackColor = &H8080FF
       Mensaje = "Valor del dato numerico invalido."
       MsgBox Mensaje, vbCritical, "Atencion!"
       Text2.SetFocus
    Else
       If (Text2.Text <= 0) Then
          Beep
          Text2.BackColor = &H8080FF
          Mensaje = "Valor del dato numerico debe ser mayor o igual a 1."
          MsgBox Mensaje, vbCritical, "Atencion!"
          Text2.SetFocus
       End If
    End If
    Text2.BackColor = &HFFFFFF
End Sub

Private Sub Text3_LostFocus()
    Dim Mensaje
    If Not IsNumeric(Text3.Text) Then
       Beep
       Text3.BackColor = &H8080FF
       Mensaje = "Valor del dato numerico invalido."
       MsgBox Mensaje, vbCritical, "Atencion!"
       Text3.SetFocus
    Else
       If (Text3.Text <= 0) Then
          Beep
          Text3.BackColor = &H8080FF
          Mensaje = "Valor del dato numerico debe ser mayor o igual a 1."
          MsgBox Mensaje, vbCritical, "Atencion!"
          Text3.SetFocus
       End If
    End If
    Text3.BackColor = &HFFFFFF
End Sub
