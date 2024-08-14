VERSION 5.00
Begin VB.Form Principal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "IMPRESION DE TARJETAS PRODUCTOS EN INVENTARIO"
   ClientHeight    =   4140
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4140
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Productos en Inventario"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "De"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Tarjetas para Control"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin VB.Menu Actualizar 
      Caption         =   "Actualizar"
      Begin VB.Menu Norma 
         Caption         =   "Logotipo de la Norma"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu Elaborar_Tarjetas 
      Caption         =   "Elaborar Tarjetas"
      Begin VB.Menu Inventario_Productos 
         Caption         =   "Productos Terminados"
      End
   End
   Begin VB.Menu Consultas_Reportes 
      Caption         =   "Consultas/Reportes"
      Begin VB.Menu Tarjetas_Elaboradas 
         Caption         =   "Tarjetas Elaboradas"
      End
      Begin VB.Menu RESUMEN_TARJ_vTP 
         Caption         =   "Tarjetas Impresas x Tipo Producto"
      End
      Begin VB.Menu RESUMEN_TARJ_vFECHA 
         Caption         =   "Tarjetas Impresas - Periodo fecha"
      End
   End
   Begin VB.Menu Mantenimiento 
      Caption         =   "Mantenimiento"
      Begin VB.Menu INVTARJ01_DAT 
         Caption         =   "INVTARJ01_DAT"
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*----------------------------------------------------
'* MODULO PRINCIPAL
'* Impresion de Tarjetas para Productos
'* Autor: Henry J Pulgar B.
'* Fecha de creacion
'* ( Como compendio General ): 19 de Agosto de 2002.
'* Ultima Fecha de Actualizacion : Septiembre 04, 2003.
'*----------------------------------------------------
Public CurrentUser As String
'

Private Sub Inventario_Productos_Click()
     Form2_v2.Show
     Form2_v2.Combo1.Text = "AT"
     Form2_v2.Text1.Text = Date
End Sub

Private Sub Mallas_Especiales_Click()
   Form1.Show
End Sub

Private Sub INVTARJ01_DAT_Click()
   ACTUALIZAR_INVTARJ01.Show
End Sub

Private Sub Norma_Click()
  'INVTARJ00_FRM.Show
  NORMAS_COVENIN.Show
End Sub

Private Sub RESUMEN_TARJ_vFECHA_Click()
  CurrentUser = "OPS$DESPRO03/OPS$DESPRO03@bd806"
  CurrentDir = "" ' <- Definir esta variable en tiempo de ejecucion (.EXE)
  'CurrentDir = "f:\vb6\proyectos\Tarjetas_Productos_Inv\"
  Comando = "rwrun60 report=" & CurrentDir & "RESUMEN_TARJ_vFECHA.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

Private Sub RESUMEN_TARJ_vTP_Click()
  CurrentUser = "OPS$DESpro03/OPS$DESPRO03@bd806"
  CurrentDir = ""  ' <- Definir esta variable en tiempo de ejecucion (.EXE)
  'CurrentDir = "f:\vb6\proyectos\Tarjetas_Productos_Inv\"
  Comando = "rwrun60 report=" & CurrentDir & "RESUMEN_TARJ_vTP.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

Private Sub Salir_Click()
   Unload Me  '?????
   Unload Principal
End Sub

Private Sub Tarjetas_Elaboradas_Click()
  TARJETAS_IMPRESAS.Show
End Sub
