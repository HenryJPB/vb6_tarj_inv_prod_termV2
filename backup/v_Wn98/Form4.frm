VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Actualizar Turno y Codigo de la maquina"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form4"
   ScaleHeight     =   3870
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   2880
      List            =   "Form4.frx":0013
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Boton_Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2760
      Picture         =   "Form4.frx":0026
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Boton_Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   1920
      Picture         =   "Form4.frx":0198
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Codigo(s) de la(s) maquina(s)?:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de Entrega? (dd-mm-yyyy):"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Turno?:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "ACTUALIZAR"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*------------------------------------------------------------------
'* Form4: ACTUALIZAR DATOS del GRID proveniente del formulario,
'*        TARJETAS_IMPRESAS
'* MODULO: Form4 ( nombre logico )
'* NOMBRE FISICO: Form4.frm
'* Autor: Henry J. Pulgar B.
'* Fecha de Creacion: 11 de Septiembre de 2002.
'* Modificado el :    11 de Septiembre de 2002.
'* Elaborado en Visual B. v6.0
'*------------------------------------------------------------------

Private Sub Boton_Aceptar_Click()
   TARJETAS_IMPRESAS.cmdUpdate
End Sub

Private Sub Form_Load()
   'Field-> Turno
   If Not IsNull(TARJETAS_IMPRESAS.adoPrimaryRS("TURNO")) Then
          Combo1.Text = TARJETAS_IMPRESAS.adoPrimaryRS("TURNO")
   End If
   'Field-> Fecha Entrega al Supervisor
   If Not IsNull(TARJETAS_IMPRESAS.adoPrimaryRS("FECHA_ENTREGA")) Then
          Text1.Text = TARJETAS_IMPRESAS.adoPrimaryRS("FECHA_ENTREGA")
   End If
   'Field->Cod de la Maq.
   If Not IsNull(TARJETAS_IMPRESAS.adoPrimaryRS("MAQUINA")) Then
          Text2.Text = TARJETAS_IMPRESAS.adoPrimaryRS("MAQUINA")
   End If
End Sub

Private Sub Boton_Cancelar_Click()
  Unload Me
End Sub

Private Sub Text2_LostFocus()
  Text2.Text = UCase(Text2.Text)
End Sub
