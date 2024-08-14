VERSION 5.00
Begin VB.Form Principal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "IMPRESION DE TARJETAS PRODUCTOS EN INVENTARIO"
   ClientHeight    =   5595
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Perpetua Titling MT"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5595
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Productos en Inventario"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1215
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "De"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Tarjetas para Control"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
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
'--------------------------------------------------------------------------------------
' Bqto: 26 octubre 2017
' Ref doc: http://uenics.evansville.edu/~hwang/f06-courses/cs350/pizza3.Designer.vb
' NOTA:  funciona esto...???
'        El siguiente metodo funciona para entidades tipo Class.
'--------------------------------------------------------------------------------------
Private Sub InitializeComponent()
        '
        'MainMenuStrip
        '
        
        MsgBox "Aquica"   '  ?????
        'Me.MainMenuStrip.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Me.MainMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ActionsMenuItem, Me.HelpMenuItem, Me.AnotherFormMenuItem})
        'Me.MainMenuStrip.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        'Me.MainMenuStrip.Location = New System.Drawing.Point(0, 0)
        'Me.MainMenuStrip.Name = "MainMenuStrip"
        'Me.MainMenuStrip.Size = New System.Drawing.Size(437, 24)
        'Me.MainMenuStrip.TabIndex = 13
        'Me.MainMenuStrip.Text = "MainMenuStrip"
        '
        'ActionsMenuItem
        '
        'Me.ActionsMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SubmitMenuItem, Me.ResetMenuItem, Me.ToolStripSeparator1, Me.ExitMenuItem})
        'Me.ActionsMenuItem.Name = "ActionsMenuItem"
        'Me.ActionsMenuItem.Size = New System.Drawing.Size(64, 20)
        'Me.ActionsMenuItem.Text = "&Actions"
        '
        'SubmitMenuItem
        '
        'Me.SubmitMenuItem.Name = "SubmitMenuItem"
        'Me.SubmitMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        'Me.SubmitMenuItem.Size = New System.Drawing.Size(144, 22)
        'Me.SubmitMenuItem.Text = "&Submit"
        '
        'ResetMenuItem
        '
        'Me.ResetMenuItem.Name = "ResetMenuItem"
        'Me.ResetMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F2
        'Me.ResetMenuItem.Size = New System.Drawing.Size(144, 22)
        'Me.ResetMenuItem.Text = "&Reset"
        '
        'ExitMenuItem
        '
        'Me.ExitMenuItem.Name = "ExitMenuItem"
        'Me.ExitMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F3
        'Me.ExitMenuItem.Size = New System.Drawing.Size(144, 22)
        'Me.ExitMenuItem.Text = "E&xit"
        '
        'HelpMenuItem
        '
        'Me.HelpMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutMenuItem})
        'Me.HelpMenuItem.Name = "HelpMenuItem"
        'Me.HelpMenuItem.Size = New System.Drawing.Size(49, 20)
        'Me.HelpMenuItem.Text = "&Help"
        '
        'AboutMenuItem
        '
        'Me.AboutMenuItem.Name = "AboutMenuItem"
        'Me.AboutMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F4
        'Me.AboutMenuItem.Size = New System.Drawing.Size(138, 22)
        'Me.AboutMenuItem.Text = "About"
        '
        'AnotherFormMenuItem
        '
        'Me.AnotherFormMenuItem.Name = "AnotherFormMenuItem"
        'Me.AnotherFormMenuItem.Size = New System.Drawing.Size(97, 20)
        'Me.AnotherFormMenuItem.Text = "AnotherForm"
        '
    End Sub

Private Sub Inventario_Productos_Click()
     'Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
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
