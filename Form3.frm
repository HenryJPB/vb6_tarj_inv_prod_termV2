VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Elaborar Tarjetas de Productos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Boton_Conforme 
      Caption         =   "Conforme"
      Height          =   615
      Left            =   1680
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Boton_Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   3000
      Picture         =   "Form3.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Al"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Desde la No. "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "tarjeta(s) ?. "
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Fueron impresas correctamente"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "ACTUALIZAR REGISTRO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'*  -----------------------------------------------------
'*  IMPRIMIR TARJETAS DE INVENTARIO PRODUCTOS TERMINADOS
'*  -----------------------------------------------------
'*
'*  Ambiente: Visual Basic v 6.00
'*  Nombre del Proyecto: Inv_Traj_Prod
'*  Autor: Henry J. Pulgar B.
'*  Nombre logico del archivo: Form3
'*  Fecha de creacion: 29 de Agosto 2002.
'*  Ultima fecha actualizacion: 29 de Agosto de 2002.
'********************************************************


Private Sub Form_Load()
    Form3.Text1.Text = Form2_v2.Text2.Text
    Form3.Text2.Text = iNVTARJ01v2_FRM.adoPrimaryRS("C1_LOTE_ANT")
    Form3.Text3.Text = iNVTARJ01v2_FRM.adoPrimaryRS("C1_LOTE_PROX")
End Sub

Private Sub Boton_Cancelar_Click()
   Unload Form3
End Sub

Private Sub Boton_Conforme_Click()
  iNVTARJ01v2_FRM.cmdUpdate_Click
End Sub

'+----------------------------------------------------
Function Conforme()
   Dim Mensaje, Botones, Titulo, Respuesta
   Mensaje = "Conforme?"
   Botones = vbYesNo + vbDefaultButton2
   Titulo = "Conforme"
   Conforme = MsgBox(Mensaje, Botones, Titulo)
End Function

Private Sub Text1_Change()
   Text3.Text = Val(Text2.Text) + Val(Text1.Text) - 1
End Sub

Private Sub Text1_LostFocus()
    Dim Mensaje
    If Not IsNumeric(Text1.Text) Then
       Mensaje = "Valor del dato numerico invalido."
       MsgBox Mensaje, vbCritical, "Atencion!"
       Text1.SetFocus
    Else
       'MsgBox Text1.Text
       'MsgBox Form2.Text2.Text
       If (Val(Text1.Text) > Val(Form2.Text2.Text)) Then
          Mensaje = "Valor del dato numerico debe ser menor o igual a " & Form2.Text2.Text & "."
          MsgBox Mensaje, vbCritical, "Atencion!"
          Text1.SetFocus
       End If
    End If
End Sub


