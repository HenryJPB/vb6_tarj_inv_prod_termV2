VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form iNVTARJ01v2_FRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ELABORAR TARJETAS DE PRODUCTOS "
   ClientHeight    =   7035
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7755
   Begin VB.ComboBox Combo_TipoCercha 
      DataField       =   "C1_TIPO_CERCHA"
      Height          =   315
      ItemData        =   "INVTARJ01v2_FRM.frx":0000
      Left            =   5640
      List            =   "INVTARJ01v2_FRM.frx":000A
      TabIndex        =   13
      Text            =   "DISCONTINUAS"
      Top             =   2150
      Width           =   1575
   End
   Begin VB.CheckBox Check_Suprime_Norma 
      Caption         =   "Suprimir logo de la norma."
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFC0&
      DataField       =   "C1_NOMBRE_CLIENTE"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.ComboBox ComboLenguaje 
      DataField       =   "C1_LENGUAJE"
      Height          =   315
      ItemData        =   "INVTARJ01v2_FRM.frx":0027
      Left            =   5640
      List            =   "INVTARJ01v2_FRM.frx":0031
      TabIndex        =   17
      Text            =   "ESPAÑOL"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_PESO_ATADO_AUX"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   20
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_PESO_ATADO"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   2040
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_CANTIDAD_AUX"
      Height          =   285
      Index           =   15
      Left            =   5640
      TabIndex        =   19
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_PESO_AUX"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   18
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Boton_Catalogo 
      Height          =   375
      Left            =   3360
      Picture         =   "INVTARJ01v2_FRM.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Boton_Cancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4320
      Picture         =   "INVTARJ01v2_FRM.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Boton_Imprimir 
      Caption         =   "Conforme"
      Height          =   495
      Left            =   3120
      Picture         =   "INVTARJ01v2_FRM.frx":05FA
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5880
      Width           =   855
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7755
      TabIndex        =   45
      Top             =   6435
      Width           =   7755
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3360
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6720
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5640
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4560
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edición"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2280
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ag&regar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7755
      TabIndex        =   39
      Top             =   6735
      Width           =   7755
      Begin VB.CommandButton cmdLast 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4560
         Picture         =   "INVTARJ01v2_FRM.frx":0B2C
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4200
         Picture         =   "INVTARJ01v2_FRM.frx":0CBE
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Enabled         =   0   'False
         Height          =   300
         Left            =   360
         Picture         =   "INVTARJ01v2_FRM.frx":0E50
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Enabled         =   0   'False
         Height          =   300
         Left            =   0
         Picture         =   "INVTARJ01v2_FRM.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_TURNO"
      Height          =   285
      Index           =   13
      Left            =   5640
      TabIndex        =   22
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_MAQUINA"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   21
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_CANTIDAD"
      Height          =   285
      Index           =   11
      Left            =   5640
      TabIndex        =   16
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00E0E0E0&
      DataField       =   "C1_ESPACIAM"
      Height          =   285
      Index           =   10
      Left            =   5640
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_LONGITUD"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#####9.99"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   5640
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_PESO"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00E0E0E0&
      DataField       =   "C1_DIAMETRO"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_ANCHO"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_TIPO"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_ORDEN_FAB"
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      DataField       =   "C1_LOTE_PROX"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00FFFFFF&
      DataField       =   "C1_LOTE_ANT"
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      DataField       =   "C1_FECHA_TARJ"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      DataField       =   "C1_TIPO_PROD"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "La tarjeta sera impresa en:  "
      Height          =   255
      Left            =   3600
      TabIndex        =   61
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "Peso Atado? (999.99):"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   60
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Peso Atado?: (999.99)"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   59
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   DATOS ULTIMO ATADO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   58
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Line Line4 
      X1              =   7560
      X2              =   7560
      Y1              =   4080
      Y2              =   4920
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7560
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   4080
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cantidad ?:"
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   57
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Peso? (999.99):"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   56
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "<:-(Codigo / Tipo )"
      Height          =   255
      Left            =   3960
      TabIndex        =   55
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Nombre_Cliente 
      Caption         =   "Nombre Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "DETALLE DE LA TARJETA"
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
      Left            =   2040
      TabIndex        =   53
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Turno (0,1,2,3 o 4):"
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   38
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo(s) de la maquina(s):"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   37
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cantidad? (9999):"
      Height          =   255
      Index           =   11
      Left            =   3600
      TabIndex        =   36
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Espaciam/Separacion?:"
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   35
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Longitud? (999.99):"
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   34
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Peso? (999.99):"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   33
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Diametro?:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   32
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ancho/Altura? (999.99):"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   31
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Caracteristicas del Prod?:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   30
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Orden de Fabricacion?:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Al"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   26
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Imprimir Desde la Tarj No."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha de la Tarjeta:"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   23
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo de Prod:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "iNVTARJ01v2_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-----------------------------------------------------------------------
'*            *******************************************************
'*  PROYECTO: Imprimir Tarjetas para Identificar Productos Terminados
'*            *******************************************************
'*                                v. Windows 2000.
'*  MODULO: INVTARJ01v2_FRM ( nombre logico )
'*  NOMBRE FISICO: INVTARJ01v2_FRM.frm
'*  Actualizar datos para impresion de Tarjetas
'*  de Inventario de Productos
'*  Elaborado x Wizard de formas VB 6.0 basado en "Codigo ADO".
'*  Actualizado y personalizado x Henry J. Pulgar B.
'*  Creado el 28 de Agosto del año 2002.
'*  Actualizado el 03 de Julio del año 2006.
'*  * NOTA IMPORTANTE:(1) Esta aplicacion requiere que la tabla
'*                        INVTARJ01_DAT no este nula ( sin registros ).
'*                        Debe poseer al menos un registro (puede se nulo).
'*   (2). Las definiciones de Mallas en Rollos A..D fueron solicitadas x
'*        el Dpto. Produccion el Enero 10,2006. Para crear una nueva trazabilidad
'*        x  tipo de Malla ( 50 m, 100 m, etc. de long. ).
'*   (3). Este dia 03 de Julio del 2006 en el desarrollo del string/sql
'*        para conectarse a la base de Datos de Oracle v 8.0.6; plataforma Windows 2000,
'*        fue necesario emplear el recurso ..., TO_NUMBER( Nombre_Variable ) Nombre_Variable,
'*        ... para poder acceder a la lectura de campos numericos c/punto flotante.
'*-----------------------------------------------------------------------

Public WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Public N1, N2, N3 As Integer

Private Sub Boton_Cancelar_Click()
  Unload iNVTARJ01v2_FRM
End Sub

'* Activar el CATALOGO.
Private Sub Boton_Catalogo_Click()
   If (txtFields(0) <> "MP") Then
       CATALOGO_PRODUCTOS_STANDARD.Show
   Else
       If (IsNull(txtFields(4)) Or txtFields(4) = "") Then
           Mensaje = "No de Orden de fabricacion no definida."
           MsgBox Mensaje, vbCritical, "Atencion!"
       Else
            'MsgBox "La orden no es nula ????"
            CATALOGO_PRODUCTOS_ESPECIALES.Show
       End If
   End If
End Sub


Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  '
  ' Originalmente:
  'db.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
  '
  ' Usuario para efectos de prueba y ensayo:
  'db.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$contab;pwd=ops$contab;"
  '
  ' Usuario en tiempo real.
  db.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$despro03;pwd=ops$despro03;"
  '
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select C1_TIPO_PROD," & _
                    "C1_FECHA_TARJ," & _
                    "C1_FECHA_ENTREGA," & _
                    "C1_LOTE_ANT," & _
                    "C1_LOTE_PROX," & _
                    "C1_ORDEN_FAB," & _
                    "C1_NOMBRE_CLIENTE," & _
                    "C1_TIPO," & _
                    "C1_TIPO_CERCHA," & _
                    "TO_NUMBER( C1_ANCHO ) C1_ANCHO," & _
                    "C1_DIAMETRO," & _
                    "TO_NUMBER( C1_PESO ) C1_PESO," & _
                    "TO_NUMBER( C1_PESO_AUX ) C1_PESO_AUX," & _
                    "TO_NUMBER( C1_PESO_ATADO ) C1_PESO_ATADO," & _
                    "TO_NUMBER( C1_PESO_ATADO_AUX ) C1_PESO_ATADO_AUX," & _
                    "TO_NUMBER( C1_LONGITUD ) C1_LONGITUD," & _
                    "C1_ESPACIAM," & _
                    "C1_CANTIDAD," & _
                    "C1_CANTIDAD_AUX," & _
                    "C1_MAQUINA," & _
                    "C1_TURNO, " & _
                    "C1_LENGUAJE " & _
                    "from INVTARJ01_DAT " & _
                    "Order by C1_TIPO_PROD", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
  
  cmdAdd_Click
  INICIALIZAR_VALORES
End Sub
'*------------------------------------------------------------
Private Sub INICIALIZAR_VALORES()
   Dim Mensaje
   '* Ejecutar un segundo query; Ubicar ultima tarjeta impresa.
   Dim Coneccion2 As Connection
   Set Coneccion2 = New Connection
   Coneccion2.CursorLocation = adUseClient
   'Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA815;uid=ops$desinv02;pwd=ops$desinv02;"
   Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$despro03;pwd=ops$despro03;"
   Set Q1 = New Recordset
   '* Ejecuta una instruccion centinela para verificar si la
   '* tabla INVTARJ01_DAT esta vacia o no:
   CadenaSQL = "select C1_LOTE_ANT from INVTARJ01_DAT"
   Q1.Open CadenaSQL, Coneccion2, adOpenStatic, adLockOptimistic
   If (adoPrimaryRS.RecordCount > 0) Then
      Q1.Close
      Cadena1 = "select C1_LOTE_ANT, C1_LOTE_PROX " & _
                "from   INVTARJ01_DAT " & _
                "where  C1_TIPO_PROD = '" & Form2_v2.Combo1.Text & "' " & _
                "and    C1_FECHA_TARJ = ( select MAX( T3.C1_FECHA_TARJ ) " & _
                                         "from   INVTARJ01_DAT T3 " & _
                                         "where  T3.C1_TIPO_PROD = INVTARJ01_DAT.C1_TIPO_PROD ) "
      Cadena2 = "and     C1_LOTE_PROX = (  select MAX( T2.C1_LOTE_PROX ) " & _
                                           "from    INVTARJ01_DAT T2 " & _
                                           "Where   T2.C1_TIPO_PROD  = INVTARJ01_DAT.C1_TIPO_PROD " & _
                                           "and     T2.C1_FECHA_TARJ = INVTARJ01_DAT.C1_FECHA_TARJ ) "
      CadenaSQL = Cadena1 + Cadena2
      'MsgBox CadenaSQL
      Q1.Open CadenaSQL, Coneccion2, adOpenStatic, adLockOptimistic
      If (Q1.EOF) Then
          ProximaTarjeta = 0
      Else
          ProximaTarjeta = Q1("C1_LOTE_PROX")
      End If
   Else
      ProximaTarjeta = 0
   End If
   Q1.Close
   '* Iniciar datos
   txtFields(4).Enabled = True
   '
   With Combo_TipoCercha
         .Enabled = False
         .BackColor = &HC0FFC0
         .Text = ""
   End With
   '
   With ComboLenguaje
         .Enabled = False
         .BackColor = &HC0FFC0
   End With
   '
   adoPrimaryRS("C1_TIPO_PROD") = Form2_v2.Combo1.Text
   adoPrimaryRS("C1_FECHA_TARJ") = Form2_v2.Text1.Text
   adoPrimaryRS("C1_LOTE_ANT") = ProximaTarjeta + 1
   adoPrimaryRS("C1_LOTE_PROX") = ProximaTarjeta + Form2_v2.Text2.Text
   Select Case Form2_v2.Combo1.Text
          Case "AH"
               '
               txtFields(4).Enabled = False
               txtFields(4).BackColor = &HC0FFC0
               '
               'txtFields(7).Enabled = True
               'txtFields(7).BackColor = &H80000005
               '
               txtFields(14).Enabled = False
               txtFields(14).BackColor = &HC0FFC0
               '
               txtFields(15).Enabled = False
               txtFields(15).BackColor = &HC0FFC0
               '
               txtFields(16).Enabled = False
               txtFields(16).BackColor = &HC0FFC0
               '
               txtFields(17).Enabled = False
               txtFields(17).BackColor = &HC0FFC0
               '
               txtFields(18).Enabled = True
               txtFields(18).BackColor = &H80000005
               txtFields(18).Text = "STOCK"
               '
          Case "AT"
               '
               txtFields(4).Enabled = False
               txtFields(4).BackColor = &HC0FFC0
               '
               'txtFields(7).Enabled = True
               'txtFields(7).BackColor = &H80000005
               '
               txtFields(14).Enabled = False
               txtFields(14).BackColor = &HC0FFC0
               '
               txtFields(15).Enabled = False
               txtFields(15).BackColor = &HC0FFC0
               '
               txtFields(16).Enabled = False
               txtFields(16).BackColor = &HC0FFC0
               '
               txtFields(17).Enabled = False
               txtFields(17).BackColor = &HC0FFC0
               '
               txtFields(18).Enabled = True
               txtFields(18).BackColor = &H80000005
               txtFields(18).Text = "STOCK"
               '
               adoPrimaryRS("C1_TIPO") = "Con Resaltes"
          Case "CE"
               '
               Combo_TipoCercha.Text = "DISCONTINUAS"
               txtFields(4).Enabled = False
               txtFields(4).BackColor = &HC0FFC0
               '
               txtFields(14).Enabled = False
               txtFields(14).BackColor = &HC0FFC0
               '
               txtFields(15).Enabled = False
               txtFields(15).BackColor = &HC0FFC0
               '
               txtFields(16).Enabled = False
               txtFields(16).BackColor = &HC0FFC0
               '
               txtFields(17).Enabled = False
               txtFields(17).BackColor = &HC0FFC0
               '
               'adoPrimaryRS("C1_TIPO") = "C-7"
               '
               With Combo_TipoCercha
                    .Enabled = True
                    .BackColor = &H80000005
               End With
          Case "MR", "MRA", "MRB", "MRC", "MRD", "MRE"     'MR [ A..D ]: Mallas en Rollos ( MR? ) dif. calibres.
               '
               txtFields(4).Enabled = False
               txtFields(4).BackColor = &HC0FFC0
               '
               '
               txtFields(14).Enabled = False
               txtFields(14).BackColor = &HC0FFC0
               '
               txtFields(15).Enabled = False
               txtFields(15).BackColor = &HC0FFC0
               '
               txtFields(16).Enabled = False
               txtFields(16).BackColor = &HC0FFC0
               '
               txtFields(17).Enabled = False
               txtFields(17).BackColor = &HC0FFC0
               '
               'adoPrimaryRS("C1_TIPO") = "50"
          Case "MP", "MR*"    'MR*->Malla en Rollo especial, depende de una O.F.
               '
               txtFields(4).Enabled = True
               txtFields(4).BackColor = &H80000005
               '
               txtFields(14).Enabled = True
               txtFields(14).BackColor = &H80000005
               '
               txtFields(15).Enabled = True
               txtFields(15).BackColor = &H80000005
               '
               txtFields(16).Enabled = True
               txtFields(16).BackColor = &H80000005
               '
               txtFields(17).Enabled = True
               txtFields(17).BackColor = &H80000005
               '
               'adoPrimaryRS("C1_TIPO") = "50"
               '
               With ComboLenguaje
                    .Enabled = True
                    .BackColor = &H80000005
               End With
   End Select
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Public Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    'SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Public Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  '
  '************IMPORTANTE******************************************
  '* NOTA: Los siguientes campos fueron manipulados fuera
  '*       del ambito de "Genearcion Automatica de Formularios";
  '*       esto indica que fueron alterados o incluidos
  '*       posterior a la generacion automatica para
  '*       readaptarla a nuevas necesidades.
  '* Autor: Henry J. Pulgar B.
  '* Fecha Creacion: 26-09-2003.
  '* Actualizado   : 25-10-2005.
  '****************************************************************
  adoPrimaryRS("C1_FECHA_ENTREGA") = Form2_v2.Text1.Text
  adoPrimaryRS("C1_LOTE_ANT") = Form3.Text2.Text
  adoPrimaryRS("C1_LOTE_PROX") = Form3.Text3.Text
  adoPrimaryRS("C1_TIPO_CERCHA") = Mid(Combo_TipoCercha.Text, 1, 1)
  adoPrimaryRS("C1_LENGUAJE") = Mid(ComboLenguaje.Text, 1, 1)
  '
  adoPrimaryRS.UpdateBatch adAffectAll
  '
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  'SetButtons True
  mbDataChanged = False
  
  Unload Form3
  Unload Me
  
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub Boton_Imprimir_Click()
   '* Imprimir tarjetas  ( ver IMPRIMIR_CLICK  ).
   If Imprimir_Click = True Then
      Dialog02.Show
   Else
      MsgBox "Imprimir cancelado."
   End If
End Sub
'+----------------------------------------------------------
Function Conforme()
   Dim Mensaje, Botones, Titulo, Respuesta
   Mensaje = "Conforme?"
   Botones = vbYesNo + vbDefaultButton2
   Titulo = "Conforme"
   Conforme = MsgBox(Mensaje, Botones, Titulo)
End Function
'+----------------------------------------------------------
Private Sub VALIDAR_ORDEN_FAB()
  Dim Mensaje
  adoPrimaryRS("C1_ORDEN_FAB") = UCase(txtFields(4).Text)
  Dim Coneccion2 As Connection
  Set Coneccion2 = New Connection
  Coneccion2.CursorLocation = adUseClient
  'Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
  Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$despro03;pwd=ops$despro03;"
  Set Q1 = New Recordset
  '*---------------------------------------------------------
  '* Execute un query para validar la orden de fabricacion
  '* Get el nombre del cliente
  '*----------------------------------------------------------
  CadenaSQL = "select C3_ORDEN, NOMBRE_CLI_PROV " & _
              "from   INV03_DAT, CXCD_DAT " & _
              "where  C3_ORDEN = '" & adoPrimaryRS("C1_ORDEN_FAB") & "' " & _
              "and    C3_CODIGO_CLIENTE = CODIGO "
  '* MsgBox CadenaSQL
  Q1.Open CadenaSQL, Coneccion2, adOpenStatic, adLockOptimistic
  If (Q1.RecordCount > 0) Then
      'Text1.Text = Q1("NOMBRE_CLI_PROV")
      txtFields(18).Text = Q1("NOMBRE_CLI_PROV")
  Else
      Mensaje = "Orden de fabricacion no definida."
      MsgBox Mensaje, vbCritical, "Atencion!"
  End If
  Q1.Close
End Sub '* VALIDAR_ORDEN_FAB ...

'************************************************************
'************************************************************
Private Sub CHECK_ITEM_IMPRESO()
  Dim Coneccion3 As Connection
  Set Coneccion3 = New Connection
  Coneccion3.CursorLocation = adUseClient
  'Coneccion3.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
  Coneccion3.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$despro03;pwd=ops$despro03;"
  Set Q3 = New Recordset
  '*---------------------------------------------------------
  '* Execute un query para validar la orden de fabricacion
  '*----------------------------------------------------------
  No_Orden = txtFields(4).Text
  Cod_Prod = txtFields(5).Text
  Char_Pos = InStr(Cod_Prod, "'")
  If Char_Pos > 0 Then
      Mid(Cod_Prod, Char_Pos) = "_"
      ClausulaSQL = "select C1_TIPO_PROD, C1_ORDEN_FAB, C1_FECHA_TARJ " & _
                    "from   INVTARJ01_DAT " & _
                    "where  C1_ORDEN_FAB =  '" & No_Orden & "' " & _
                    "and    C1_TIPO like '" & Cod_Prod & "'"
  Else
     ClausulaSQL = "select C1_TIPO_PROD, C1_ORDEN_FAB, C1_FECHA_TARJ " & _
                   "from   INVTARJ01_DAT " & _
                   "where  C1_ORDEN_FAB =  '" & No_Orden & "' " & _
                   "and    C1_TIPO = '" & Cod_Prod & "'"
  End If 'El valor del campo posee el char "'"
  '* MsgBox ClausulaSQL
  Q3.Open ClausulaSQL, Coneccion3, adOpenStatic, adLockOptimistic
  If (Q3.RecordCount > 0) Then
      Mensaje = "Item seleccionado fue impreso el " & Q3("C1_FECHA_TARJ") & "."
      MsgBox Mensaje, vbCritical, "Atencion!"
      Q3.Close
      'ITEM_IMPRESO = True
  Else
      'OK! La tarjetas del mencionado producto no ha sido impresa.
      'MsgBox "Ojo-> Item de esta orden no ha sido impersa", vbCritical, "Atencion!"
      Q3.Close
      'ITEM_IMPRESO = False
  End If
End Sub  ' CHECK_ITEM_IMPRESO

'+-----------------VALIDAR ITEM de la ORDEN -----------------
Sub VALIDAR_ITEM_ORDEN()
  Dim Mensaje
  Const_Peso = 0.0061653
  adoPrimaryRS("C1_TIPO") = UCase(txtFields(5).Text)
  Dim Coneccion2 As Connection
  Set Coneccion2 = New Connection
  Coneccion2.CursorLocation = adUseClient
  'Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
  Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$despro03;pwd=ops$despro03;"
  Set Q2 = New Recordset
  '*---------------------------------------------------------
  '* Execute un query para validar la orden de fabricacion
  '* Get el nombre del cliente
  '*----------------------------------------------------------
  No_Orden = adoPrimaryRS("C1_ORDEN_FAB")
  Cod_Prod = adoPrimaryRS("C1_TIPO")
  Char_Pos = InStr(Cod_Prod, "'")
  If Char_Pos > 0 Then
     Mid(Cod_Prod, Char_Pos) = "_"
     ClausulaSQL = "select TO_NUMBER( C4_TAML ) Longitud, TO_NUMBER( C4_TAMT ) Ancho, " & _
                       "TO_NUMBER( C4_NUML ) Numl, TO_NUMBER( C4_NUMT ) Numt, " & _
                       "TO_NUMBER( C4_DIAML ) Diaml, TO_NUMBER( C4_DIAMT ) Diamt, " & _
                       "C4_SEPL||' x '||C4_SEPT Espaciam, C4_DIAML||' x '||C4_DIAMT Diam," & _
                       "TO_NUMBER( C4_LAM ) Cantidad " & _
                       "from   INV04_DAT " & _
                       "where  C4_ORDEN =  '" & No_Orden & "' " & _
                       "and    C4_CODIGO like '" & Cod_Prod & "' "
  Else
  ClausulaSQL = "select TO_NUMBER( C4_TAML ) Longitud, TO_NUMBER( C4_TAMT ) Ancho, " & _
                       "TO_NUMBER( C4_NUML ) Numl, TO_NUMBER( C4_NUMT ) Numt, " & _
                       "TO_NUMBER( C4_DIAML ) Diaml, TO_NUMBER( C4_DIAMT ) Diamt, " & _
                       "C4_SEPL||' x '||C4_SEPT Espaciam, C4_DIAML||' x '||C4_DIAMT Diam," & _
                       "TO_NUMBER( C4_LAM ) Cantidad " & _
                       "from   INV04_DAT " & _
                       "where  C4_ORDEN =  '" & No_Orden & "' " & _
                       "and    C4_CODIGO = '" & Cod_Prod & "' "
  End If  'If el valor del campo posee un char "'"
  '* MsgBox ClausulaSQL
  Q2.Open ClausulaSQL, Coneccion2, adOpenStatic, adLockOptimistic
  If (Q2.RecordCount > 0) Then
      'adoPrimaryRS("C1_ANCHO") = Q2("Ancho")
      'Ancho = Format(Q2("Ancho"), "###,##0.00")
      adoPrimaryRS("C1_ANCHO") = Q2("Ancho")
      Ancho = Q2("Ancho")
      Numt = Q2("Numt")
      Diamt = Q2("Diamt")
      adoPrimaryRS("C1_LONGITUD") = Q2("Longitud")
      Longitud = Q2("Longitud")
      Numl = Q2("Numl")
      Diaml = Q2("Diaml")
      adoPrimaryRS("C1_DIAMETRO") = Q2("Diam")
      adoPrimaryRS("C1_ESPACIAM") = Q2("Espaciam")
      adoPrimaryRS("C1_CANTIDAD") = Q2("Cantidad")
      adoPrimaryRS("C1_CANTIDAD_AUX") = Q2("Cantidad")
      '
      PesoT = Ancho * Numt * (Diamt * Diamt) * Const_Peso
      PesoL = Longitud * Numl * (Diaml * Diaml) * Const_Peso
      '
      adoPrimaryRS("C1_PESO") = Round(PesoT + PesoL, 2)
      adoPrimaryRS("C1_PESO_AUX") = Round(PesoT + PesoL, 2)
      adoPrimaryRS("C1_PESO_ATADO") = Round((PesoT + PesoL) * Q2("Cantidad"), 2)
      adoPrimaryRS("C1_PESO_ATADO_AUX") = Round((PesoT + PesoL) * Q2("Cantidad"), 2)
      CHECK_ITEM_IMPRESO
  Else
      Mensaje = "Codigo del Item no registrado."
      MsgBox Mensaje, vbCritical, "Atencion!"
  End If
  Q2.Close
End Sub '---VALIDAR_ITEM_ORDEN...
'*----------------------------------------------------------
'* Calcular peso del atado
'*----------------------------------------------------------
Private Sub CALCULAR_PESO_ATADO()
   Dim Peso As Double
   Dim Cantdad As Double
   'Peso del atado:
   If (Not IsNull(txtFields(8).Text)) And (Not IsNull(txtFields(11).Text)) Then
       'CantidadTxt = txtFields(11).Text
       'If InStr(CantidadTxt, ",") <> 0 Then
       '   Mid(CantidadTxt, InStr(CantidadTxt, ",")) = "."
       'End If
       'PesoTxt = txtFields(8).Text
       'If InStr(PesoTxt, ",") <> 0 Then
       '   Mid(PesoTxt, InStr(PesoTxt, ",")) = "."
       'End If
       'Peso = Val(PesoTxt)
       'Cantidad = Val(CantidadTxt)
       Peso = CDbl(txtFields(8).Text)
       Cantidad = CDbl(txtFields(11).Text)
       adoPrimaryRS("C1_PESO_ATADO") = Peso * Cantidad
   End If
   'Peso ultimo atado:
   If (Not IsNull(txtFields(14))) And (Not IsNull(txtFields(15))) Then
       Peso = CDbl(txtFields(14).Text)
       Cantidad = CDbl(txtFields(15).Text)
       adoPrimaryRS("C1_PESO_ATADO_AUX") = Peso * Cantidad
   End If
 End Sub
'*--------------------------------------------------------
'+------------ VALIDAR CAMPOS GENERALES-------------------
'*--------------------------------------------------------
Private Sub ComboLenguaje_LostFocus()
    ComboLenguaje.Text = UCase(ComboLenguaje.Text)
    If (ComboLenguaje <> "ESPAÑOL") And (ComboLenguaje.Text <> "INGLES") Then
        Beep
        ComboLenguaje.BackColor = &H8080FF
        MsgBox "Codigo del lenguaje no definido", vbCritical, "Atencion!"
        ComboLenguaje.SetFocus
        ComboLenguaje.BackColor = &HFFFFFF
    End If
End Sub

Private Sub Combo_TipoCercha_LostFocus()
    Combo_TipoCercha.Text = UCase(Combo_TipoCercha.Text)
    If (Combo_TipoCercha.Text <> "DISCONTINUAS") And (Combo_TipoCercha.Text <> "CONTINUAS") Then
        Beep
        Combo_TipoCercha.BackColor = &H8080FF
        MsgBox "Tipo de cercha no definido", vbCritical, "Atencion!"
        Combo_TipoCercha.SetFocus
        Combo_TipoCercha.BackColor = &HFFFFFF
    End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
   Dim Mensaje
   If (Index = 2) And Not IsNull(txtFields(2).Text) Then
           txtFields(3).Text = Val(txtFields(2).Text) + Val(Form2_v2.Text2.Text) - 1
           adoPrimaryRS("C1_LOTE_PROX") = txtFields(3).Text
   ElseIf (Index = 4) And (Len(txtFields(4))) Then
           VALIDAR_ORDEN_FAB
   ElseIf (Index = 5) And (Form2_v2.Combo1.Text = "MP") Then
           VALIDAR_ITEM_ORDEN
   ElseIf (Index = 6) Then
           If (Len(txtFields(6)) > 0) And (Not IsNumeric(txtFields(6))) Then
               Beep
               txtFields(6).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(6).SetFocus
               txtFields(6).BackColor = &HFFFFFF
           End If
   ElseIf (Index = 7) Then
           '  Campo diametro es un campo alfanumerico <-> no numerico.
   ElseIf (Index = 8) Then
           If (Len(txtFields(8)) > 0) And (Not IsNumeric(txtFields(8))) Then
               Beep
               txtFields(8).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(8).SetFocus
               txtFields(8).BackColor = &HFFFFFF
           ElseIf txtFields(0) = "MP" Then
               CALCULAR_PESO_ATADO
           End If
   ElseIf (Index = 9) Then
           If (Len(txtFields(9)) > 0) And (Not IsNumeric(txtFields(9))) Then
               Beep
               txtFields(9).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(9).SetFocus
               txtFields(9).BackColor = &HFFFFFF
           End If
   ElseIf (Index = 11) Then
           If (Len(txtFields(11)) > 0) And (Not IsNumeric(txtFields(11))) Then
               Beep
               txtFields(11).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(11).SetFocus
               txtFields(11).BackColor = &HFFFFFF
           ElseIf txtFields(0) = "MP" Then
               CALCULAR_PESO_ATADO
           End If
   ElseIf (Index = 14) Then
           If (Len(txtFields(14)) > 0) And (Not IsNumeric(txtFields(14))) Then
               Beep
               txtFields(14).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(14).SetFocus
               txtFields(14).BackColor = &HFFFFFF
           ElseIf txtFields(0) = "MP" Then
               CALCULAR_PESO_ATADO
           End If
   ElseIf (Index = 15) Then
          If (Len(txtFields(15)) > 0) And (Not IsNumeric(txtFields(15))) Then
               Beep
               txtFields(15).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(15).SetFocus
               txtFields(15).BackColor = &HFFFFFF
           ElseIf txtFields(0) = "MP" Then
               CALCULAR_PESO_ATADO
           End If
    ElseIf (Index = 16) Then
          If (Len(txtFields(16)) > 0) And (Not IsNumeric(txtFields(16))) Then
               Beep
               txtFields(16).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(16).SetFocus
               txtFields(16).BackColor = &HFFFFFF
          End If
    ElseIf (Index = 17) Then
          If (Len(txtFields(17)) > 0) And (Not IsNumeric(txtFields(17))) Then
               Beep
               txtFields(17).BackColor = &H8080FF
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(17).SetFocus
               txtFields(17).BackColor = &HFFFFFF
           End If
   End If
End Sub '*Lost_Focus()
'* OLD *
Private Sub txtFields_LostFocus_OLD(Index As Integer)
   Dim Mensaje
   If (Index = 4) And (Not IsNull(txtFields(4))) Then
       adoPrimaryRS("C1_ORDEN_FAB") = UCase(txtFields(4).Text)
       If (txtFields(0).Text = "MP") Then
           Dim Coneccion2 As Connection
           Set Coneccion2 = New Connection
           Coneccion2.CursorLocation = adUseClient
           'Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
           Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$desinv02;pwd=ops$desinv02;"
           Set Q1 = New Recordset
           '*---------------------------------------------------------
           '* Execute un query para validar la orden de fabricacion
           '* Get el nombre del cliente
           '*----------------------------------------------------------
           CadenaSQL = "select C3_ORDEN, NOMBRE_CLI_PROV " & _
                       "from   INV03_DAT, CXCD_DAT " & _
                       "where  C3_ORDEN = '" & adoPrimaryRS("C1_ORDEN_FAB") & "' " & _
                       "and    C3_CODIGO_CLIENTE = CODIGO "
           'MsgBox CadenaSQL
           Q1.Open CadenaSQL, Coneccion2, adOpenStatic, adLockOptimistic
           If (Q1.RecordCount > 0) Then
                Text1.Text = Q1("NOMBRE_CLI_PROV")
           Else
                Mensaje = "No de Orden de fabricacion no definida."
                MsgBox Mensaje, vbCritical, "Atencion!"
           End If
           Q1.Close
       End If
   ElseIf (Index = 6) Then
           If (Not IsNumeric(txtFields(6))) Then
                MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
                txtFields(6).SetFocus
           End If
   ElseIf (Index = 8) Then
           If (Not IsNumeric(txtFields(8))) Then
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(8).SetFocus
           End If
   ElseIf (Index = 9) Then
           If (Not IsNumeric(txtFields(9))) Then
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(9).SetFocus
           End If
   ElseIf (Index = 11) Then
           If (Not IsNumeric(txtFields(11))) Then
               MsgBox "Dato numerico invalido.", vbCritical, "Atencion"
               txtFields(11).SetFocus
           End If
   End If
End Sub

'*****Creacion e Implementacion: RUTINAS DE IMPRRESION *****
' ** IMPRIMIR_CLICK.**
Function Imprimir_Click()
  Dim J As Integer
  Dim Desde As Integer
  ' Valores de impresión
  Dim PrimeraPag, ÚltimaPag, NumCopias, ImpArchivo, i, T
  ' Si ocurre un error ejecutar ManipularErrorImprimir
  On Error GoTo ManipularErrorImprimir
  ' Generar un error cuando se pulse Cancelar
  CommonDialog1.CancelError = True
  ' Visualizar la caja de diálogo
  ' Iniciar Copias a imprimir = "No atados"
  'CommonDialog1.Copies = Val(Form1.Text5.Text)
  CommonDialog1.Copies = 1
  CommonDialog1.ShowPrinter
  ' Obtener las propiedades de impresión
  PrimeraPag = CommonDialog1.FromPage
  ÚltimaPag = CommonDialog1.ToPage
  NumCopias = CommonDialog1.Copies   '<- Esta instruccion no esta funcionando ????
  'Desde = Val(Form1.Text3.Text)
  Desde = 1
  'NumCopias = Val(Form1.Text4.Text)  ' Depende del numero de atados
  ImpArchivo = CommonDialog1.Flags And cdlPDPrintToFile
  ' Imprimir
  If ImpArchivo Then
    For i = Desde To NumCopias
      ' Escriba el código para enviar datos a un archivo
      'GENERAR_ARCHIVO
    Next i
  Else  'Dirigir salida a la impresora
    For i = Desde To NumCopias
         For J = 1 To 1   '1 Impresion=Cliente + 1 Impresion=Despacho
            ' Escriba el código para enviar datos a la impresora
            'IMPRIMIR_TARJETAS (i), (NumCopias) '-->Parametros pasados por valor "()" -> la var no puede ser modificada
            IMPRIMIR_TARJETAS
         Next J
    Next i
  End If
  
Imprimir_Click = True
SalirImprimir:
   Exit Function

ManipularErrorImprimir:
  ' Manipular el error
  If Err.Number = cdlCancel Then Exit Function
  MsgBox Err.Description
  Imprimir_Click = False
  Resume SalirImprimir
End Function 'Imprimir_Click_

'*----------------------------------------------------*
'* NOTA: en este modulo ajustar variables / S.O. en
'*       cuestion ( Windows NT/95/98 ????.
'*       ver SET_VARS_WNT o SET_VARS_W98.
'*----------------------------------------------------*
Private Sub IMPRIMIR_TARJETAS()
    Dim i As Integer
    Dim Norma As String
    Dim Coneccion2 As Connection
    '------------------------------------------------
    Set Coneccion2 = New Connection
    Coneccion2.CursorLocation = adUseClient
    'Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA815;uid=ops$desinv02;pwd=ops$desinv02;"
    Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$despro03;pwd=ops$despro03;"
    Set Q = New Recordset
    '*-----------------------------------------------
    '* Execute un query : get el nombre de la norma
    '*-----------------------------------------------
    CadenaSQL = "select C0_NORMA " & _
                "from   INVTARJ00_DAT " & _
                "where  C0_TIPO_PROD ='" & Mid(txtFields(0), 1, 2) & "' "
    '* Chequear query:
    'MsgBox CadenaSQL
    Q.Open CadenaSQL, Coneccion2, adOpenStatic, adLockOptimistic
    Norma = Q("C0_NORMA")
    Q.Close
    If (Check_Suprime_Norma = 1) Then
        'MsgBox "El logo de la norma debe no ser impreso"
         Norma = ""
    End If
    '----------------------
    'Printer.Font.Name = "Draft 17cpi" 'Modo Comprimido ??.
    'Printer.Font.Name = "Draft 10cpi" 'Modo Normal ??.
    'Printer.Font.Name = "Draft"
    '****************
    ' Configurar impresora:
    '****************
    Printer.Font.Name = "Device Font 17cpi"
    'Printer.Font.Bold = False
    'Printer.FontSize = 8
    'SET_VARS_W98
    SET_VARS_WN2K
    Select Case txtFields(0)
           Case "AT"
                PRINT_AT (Norma)
           Case "AH"
                PRINT_AH (Norma)
           Case "CE"
                PRINT_CE (Norma)
           Case "MP"
                PRINT_MP (Norma)
                'PRINT_MP_prueba
           Case "MR", "MRA", "MRB", "MRC", "MRD", "MRE", "MRF", "MRG", "MRH"
                PRINT_MR (Norma)
           Case "MR*"
                PRINT_MR_ESPECIAL (Norma)
    End Select
    Printer.EndDoc
End Sub

'*-----------------------------------------------------------
'* Set las variables de tabulacion global
'* para:
'* Windows NT Server 4.0
'*-----------------------------------------------------------
Private Sub SET_VARS_WNT()
   N1 = 85
   N2 = 115
   N3 = 152
End Sub 'SET_VARS_WNT

'*-----------------------------------------------------------
'* Set las variables de tabulacion global
'* para:
'* WiniK
'*-----------------------------------------------------------
Private Sub SET_VARS_WN2K()
   Aux_Tab = 5
   '
   N1 = 81 - Aux_Tab
   N2 = 111 - Aux_Tab
   N3 = 157 - Aux_Tab
End Sub 'SET_VARS_WNT

'*-----------------------------------------------------------
'* Set las variables de tabulacion global
'* para Windows 95/98.
'*-----------------------------------------------------------
Private Sub SET_VARS_W98()
   N1 = 86
   'N2 = 117
   N2 = 119
   N3 = 172
End Sub 'SET_VARS_W98

'*-----------------------------------------------------------
'* ALAMBRE TREFILADO de HERRERIA.
'*-----------------------------------------------------------
Private Sub PRINT_AH(Norma As String)
   Dim Lote As String
   '
   N1 = 7
   N2 = 41
   N3 = 84
   Punto_Ajuste = 8
   Contador = 1
   TxtCliente = "CLIENTE: "
   Printer.Print
   Designacion = "ALAMBRES PARA HERRERIA"
   For i = txtFields(2) To txtFields(3)
        Lineas_Entre_Tarjetas = 5
        Lote = Format(i, "00000") + "-" + txtFields(0) + "-" + Format(txtFields(1), "YY")
        Printer.Print Tab(N1); Lote; Tab(N2); "DESIGNACION: " + Designacion
        Printer.Print Tab(N1); "D: " + Format(txtFields(7), "###,##0.00"); Tab(N2); TxtCliente + Mid(txtFields(18).Text, 1, 18)
        Printer.Print Tab(N1); "L: " + Format(txtFields(9), "###,##0.00"); Tab(N2); "DIAMETRO: " + Format(txtFields(7), "###,##0.00") + " (mm)"; Tab(N3); "LONGITUD: " + Format(txtFields(9), "###,##0.00") + " (m)"
        Printer.Print Tab(N1); "C: " + txtFields(11); Tab(N2); "CANTIDAD: " + txtFields(11) + " (und)"; Tab(N3); "LOTE: " + Lote
        Printer.Print Tab(N1); "Ref.:________."
        Printer.Print
        Printer.Print Tab(N1); "Peso:__________."; Tab(N2); "No. Referencia:______________."
        Printer.Print
        Printer.Print Tab(N1); "Fecha:______."
        '*  Ajustar salto de pagina:
        Contador = Contador + 1
        If (Contador = Punto_Ajuste) Then
            Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas + (-2)
            Contador = 1
        End If
        '*
        For R = 1 To Lineas_Entre_Tarjetas
            Printer.Print
        Next R
   Next i
'
End Sub   'PRINT_AH  Imprimir ALAMBRE trefilado de Herreria.


'*-----------------------------------------------------------
'* ALAMBRE TREFILADO de HERRERIA. ( Old ).
'*-----------------------------------------------------------
Private Sub PRINT_AH_OLD(Norma As String)
   Dim Lote As String
   '
   N1 = N1 - 80
   N2 = N2 - 80
   N3 = N3 - 80
   Punto_Ajuste = 8
   Contador = 1
   TxtCliente = "CLIENTE: "
   Printer.Print
   Designacion = "ALAMBRES PARA HERRERIA"
   For i = txtFields(2) To txtFields(3)
        Lineas_Entre_Tarjetas = 5
        Lote = Format(i, "00000") + "-" + txtFields(0) + "-" + Format(txtFields(1), "YY")
        Printer.Print Tab(N1); "DESIGNACION: " + Designacion; Tab(N3); Tab(N3); Lote
        Printer.Print Tab(N1); TxtCliente + Mid(txtFields(18).Text, 1, 18); Tab(N3); "D: " + Format(txtFields(7), "###,##0.00")
        Printer.Print Tab(N1); "DIAMETRO: " + Format(txtFields(7), "###,##0.00"); Tab(N2); "LONGITUD: " + Format(txtFields(9), "###,##0.00") + " (mm)"; Tab(N3); "L: " + Format(txtFields(9), "###,##0.00")
        Printer.Print Tab(N1); "CANTIDAD: " + txtFields(11) + " (und)"; Tab(N2); "LOTE: " + Lote; Tab(N3); "C: " + txtFields(11)
        Printer.Print Tab(N3); "Ref.:________."
        Printer.Print
        Printer.Print Tab(N1); "No. Referencia:______________."; Tab(N3); "Peso:__________."
        Printer.Print
        Printer.Print Tab(N3); "Fecha:______."
        '*  Ajustar salto de pagina:
        Contador = Contador + 1
        If (Contador = Punto_Ajuste) Then
            Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas + (-2)
            Contador = 1
        End If
        '*
        For R = 1 To Lineas_Entre_Tarjetas
            Printer.Print
        Next R
   Next i
'
End Sub   'PRINT_AH OLD  Imprimir ALAMBRE trefilado de Herreria.

'*-----------------------------------------------------------
'* ALAMBRE TREFILADO.
'*-----------------------------------------------------------
Private Sub PRINT_AT(Norma As String)
   Dim Lote As String
   'Punto_Ajuste = 6 '<: Valor for WNT.
   Punto_Ajuste = 4  '=4 <: Valor for W98.
   Contador = 0
   TxtCliente = "CLIENTE: "
   Printer.Print
   Designacion = "ALAMBRES DE ACERO"
   For i = txtFields(2) To txtFields(3)
        Lineas_Entre_Tarjetas = 2  '<: Valor for WNT.
        'Lineas_Entre_Tarjetas = 3 '<: Valor for W98.
        Lote = Format(i, "00000") + "-" + txtFields(0) + "-" + Format(txtFields(1), "YY")
        Printer.Print Tab(N1); "DESIGNACION: " + Designacion; Tab(N3); Tab(N3); Lote
        Printer.Print Tab(N1); "GRADO: 50"; Tab(N2); TxtCliente + Mid(txtFields(18).Text, 1, 18); Tab(N3); "D: " + Format(txtFields(7), "###,##0.00")
        Printer.Print Tab(N1); "TIPO: " + txtFields(5); Tab(N2); "DIAMETRO: " + Format(txtFields(7), "###,##0.00") + " (mm)"; Tab(N3); "L: " + Format(txtFields(9), "###,##0.00")
        Printer.Print Tab(N1); "LONGITUD: " + Format(txtFields(9), "###,##0.00") + " (m)"; Tab(N2); "CANTIDAD: " + txtFields(11) + " (und)"; Tab(N3); "C: " + txtFields(11)
        Printer.Print Tab(N1); "NORMA: " + Norma; Tab(N2); "LOTE: " + Lote; Tab(N3); "Ref.:________."
        Printer.Print
        Printer.Print Tab(N1); "No. Referencia:______________."; Tab(N3); "Peso:__________."
        Printer.Print
        Printer.Print Tab(N3); "Fecha:______."
        '*  Ajustar salto de pagina:
        Contador = Contador + 1
        If (Contador = Punto_Ajuste) Then
            'Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas -1 '<: Valor for WNT
            Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas + 1 '-3<: Valor for W98.
            Contador = 0
        End If
        '*
        For R = 1 To Lineas_Entre_Tarjetas
            Printer.Print
        Next R
   Next i
End Sub
'*----------------------------------------------------------
'* CERCHAS ELECTROSOLDADAS.
'*----------------------------------------------------------
Private Sub PRINT_CE(Norma As String)
   Dim Lote As String
   Aux_Tab = 5
   Punto_Ajuste = 4 '<: Valor for WNT
   'Punto_Ajuste = 4  '<: Valor for W98
   Contador = 0
   '
   MAX_LONG_FIELD_TIPO = 18
   txtTipoCercha = Mid(RTrim(txtFields(5)) + " " + Combo_TipoCercha.Text, 1, MAX_LONG_FIELD_TIPO)
   '
   Printer.Print
   Designacion = "CERCHAS ELECTROSOLDADAS"
   For i = txtFields(2) To txtFields(3)
        Lineas_Entre_Tarjetas = 4 '<: Valor for WNT
        'Lineas_Entre_Tarjetas = 6 '<: Valor for W98
        Lote = Format(i, "00000") + "-" + txtFields(0) + "-" + Format(txtFields(1), "YY")
        Printer.Print Tab(N1); "DESIGNACION: " + Designacion; Tab(N3); Tab(N3); Lote
        Printer.Print Tab(N3); "L: " + Format(txtFields(9), "###,##0.00")
        Printer.Print Tab(N1); "TIPO: " + txtTipoCercha; Tab(N2 + Aux_Tab); "LONGITUD: " + Format(txtFields(9), "###,##0.00") + " (m)"; Tab(N3); "A: " + Format(txtFields(6), "###,##0.00")
        Printer.Print Tab(N3); "C: " + txtFields(11)
        Printer.Print Tab(N1); "ALTURA: " + Format(txtFields(6), "###,##0.00") + " (cm)"; Tab(N2 + Aux_Tab); "CANTIDAD: " + txtFields(11) + " (und)"; Tab(N3); Combo_TipoCercha.Text
        Printer.Print
        Printer.Print Tab(N1); "NORMA: " + Norma; Tab(N2 + Aux_Tab); "LOTE: " + Lote; Tab(N3); "Fecha: _________."
        '*
        Contador = Contador + 1
        If (Contador = Punto_Ajuste) Then
            'Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas - 1
            Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas + 1
            Contador = 0
        End If
        '*
        For R = 1 To Lineas_Entre_Tarjetas
            Printer.Print
        Next R
   Next i
End Sub
Private Sub PRINT_MP_prueba()
    For i = 1 To 140
        Printer.Print "***********"
    Next i
End Sub


'*----------------------------------------------------------
'* MALLAS ESPECIALES (MP)
'*----------------------------------------------------------
Private Sub PRINT_MP(Norma As String)
   Dim Lote As String
   Punto_Ajuste = 4 '<: Valor for WNT
   'Punto_Ajuste = 6 '<: Valor for W98
   Contador = 0
   Contador_Atds = Val(Form2_v2.Text3.Text) - 1
   Cont_Atds_Hasta = Str(Val(Form2_v2.Text2.Text) + Val(Form2_v2.Text3.Text) - 1)
   Printer.Print
   Designacion = "MALLAS PLANAS"
   If ComboLenguaje.Text = "ESPAÑOL" Then
      TxtDesig = "DESIG: "
      TxtOrden = "ORDEN: "
      TxtCliente = "CLIENTE: "
      TxtTipo = "TIPO: "
      TxtAncho = "ANCHO: "
      TxtLong = "LONGITUD: "
      TxtDiam = "DIAM: "
      TxtSep = "ESPACIAM: "
      TxtPeso = "PESO: "
      TxtCant = "CANTIDAD: "
      TxtNorma = "NORMA: "
      TxtLote = "LOTE: "
   ElseIf ComboLenguaje = "INGLES" Then
      TxtDesig = "DESIG: "
      TxtOrden = "ORDER: "
      TxtCliente = "CLIENT: "
      TxtTipo = "TYPE: "
      TxtAncho = "WIDTH: "
      TxtLong = "LENGTH: "
      TxtDiam = "DIAM: "
      TxtSep = "SPACING: "
      TxtPeso = "WEIGHT: "
      TxtCant = "QUANTITY: "
      TxtNorma = "STANDARD: "
      TxtLote = "LOT: "
   Else
      MsgBox "Error lenguaje no definido"
   End If
   '
   For i = txtFields(2) To txtFields(3)
        Lineas_Entre_Tarjetas = 5    'Old Value = 6 '<: Valor for WNT
        'Lineas_Entre_Tarjetas = 7   '<: Valor for W98
        Contador_Atds = Contador_Atds + 1
        Lote = Format(i, "00000") + "-" + txtFields(0) + "-" + Format(txtFields(1), "YY")
        'Printer.Print Tab(N1); TxtDesig + Designacion; Tab(N2); TxtCliente + Mid(txtFields(18).Text, 1, 22); Tab(N3); Lote
        Printer.Print Tab(N1); Designacion; Tab(N2); Mid(txtFields(18).Text, 1, 30); Tab(N3); Lote
        Printer.Print Tab(N1); TxtOrden + txtFields(4); Tab(N2); TxtTipo + txtFields(5); Tab(N3); "Nº " + txtFields(4)
        Printer.Print Tab(N1); TxtAncho + Format(txtFields(6), "###,##0.00") + " (m)"; Tab(N2); TxtLong + Format(txtFields(9), "###,##0.00") + " (m)"; Tab(N3); "T: " + txtFields(5)
        If (i = txtFields(3)) Then  ' Ultimo atado
            Printer.Print Tab(N1); TxtDiam + Format(txtFields(7), "###,##0.00") + " (mm)"; Tab(N2); TxtSep + txtFields(10) + " mm"; Tab(N3); "C: " + txtFields(15)
            Printer.Print Tab(N1); TxtPeso + Format(txtFields(14), "###,##0.00") + " (kg)"; Tab(N2); TxtCant + txtFields(15) + " (und)"; Tab(N3); "ATD: " + Str(Contador_Atds) + "/" + Cont_Atds_Hasta
            Printer.Print Tab(N1); TxtNorma + Norma; Tab(N2); TxtLote + Lote; Tab(N3); "P: " + Format(txtFields(17), "###,##0.00") + " (kg)"
        Else
            Printer.Print Tab(N1); TxtDiam + Format(txtFields(7), "###,##0.00") + " (mm)"; Tab(N2); TxtSep + txtFields(10) + " mm"; Tab(N3); "C: " + txtFields(11)
            Printer.Print Tab(N1); TxtPeso + Format(txtFields(8), "###,##0.00") + " (kg)"; Tab(N2); TxtCant + txtFields(11) + " (und)"; Tab(N3); "ATD: " + Str(Contador_Atds) + "/" + Cont_Atds_Hasta
            Printer.Print Tab(N1); TxtNorma + Norma; Tab(N2); TxtLote + Lote; Tab(N3); "P: " + Format(txtFields(16), "###,##0.00") + " (kg)"
        End If
        Contador = Contador + 1
        If (Contador = Punto_Ajuste) Then
            'Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas - 1
            Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas + 1   'Old Value = - 2
            Contador = 0
        End If
        For R = 1 To Lineas_Entre_Tarjetas
            Printer.Print
        Next R
   Next i
End Sub
'*----------------------------------------------------------
'* MALLAS EN ROLLO ESPECIAL  (MR*)
'*----------------------------------------------------------
Private Sub PRINT_MR_ESPECIAL(Norma As String)
   Dim Lote As String
   'Punto_Ajuste = 8 '<: Valor for WNT
   Punto_Ajuste = 4  '<: Valor for W98
   Contador = 0
   Contador_Atds = Val(Form2_v2.Text3.Text) - 1
   Cont_Atds_Hasta = Str(Val(Form2_v2.Text2.Text) + Val(Form2_v2.Text3.Text) - 1)
   Printer.Print
   Designacion = "MALLA EN ROLLO"
   If ComboLenguaje.Text = "ESPAÑOL" Then
      TxtDesig = "DESIG: "
      TxtOrden = "ORDEN: "
      TxtCliente = "CLIENTE: "
      TxtTipo = "TIPO: "
      TxtAncho = "ANCHO: "
      TxtLong = "LONGITUD: "
      TxtDiam = "DIAM: "
      TxtSep = "ESPACIAM: "
      TxtPeso = "PESO: "
      TxtCant = "CANTIDAD: "
      TxtNorma = "NORMA: "
      TxtLote = "LOTE: "
   ElseIf ComboLenguaje = "INGLES" Then
      TxtDesig = "DESIG: "
      TxtOrden = "ORDER: "
      TxtCliente = "CLIENT: "
      TxtTipo = "TYPE: "
      TxtAncho = "WIDTH: "
      TxtLong = "LENGTH: "
      TxtDiam = "DIAM: "
      TxtSep = "SEPARATION: "
      TxtPeso = "WEIGHT: "
      TxtCant = "QUANTITY: "
      TxtNorma = "STANDARD: "
      TxtLote = "LOT: "
   Else
      MsgBox "Error lenguaje no definido"
   End If
   '
   For i = txtFields(2) To txtFields(3)
        Lineas_Entre_Tarjetas = 5    '<: Valor for WNT
        'Lineas_Entre_Tarjetas = 6   '<: Valor for W98
        Contador_Atds = Contador_Atds + 1
        Lote = Format(i, "00000") + "-" + txtFields(0) + "-" + Format(txtFields(1), "YY")
        'Printer.Print Tab(N1); TxtDesig + Designacion; Tab(N2); TxtCliente + Mid(txtFields(18).Text, 1, 22); Tab(N3); Lote
        Printer.Print Tab(N1); Designacion; Tab(N2); Mid(txtFields(18).Text, 1, 30); Tab(N3); Lote
        Printer.Print Tab(N1); TxtOrden + txtFields(4); Tab(N2); TxtTipo + txtFields(5); Tab(N3); "Nº " + txtFields(4)
        Printer.Print Tab(N1); TxtAncho + Format(txtFields(6), "###,##0.00") + " (m)"; Tab(N2); TxtLong + Format(txtFields(9), "###,##0.00") + " (m)"; Tab(N3); "T: " + txtFields(5)
        If (i = txtFields(3)) Then  ' Ultimo atado
            Printer.Print Tab(N1); TxtDiam + Format(txtFields(7), "###,##0.00") + " (mm)"; Tab(N2); TxtSep + txtFields(10) + " mm"; Tab(N3); "C: " + txtFields(15)
            Printer.Print Tab(N1); TxtPeso + Format(txtFields(14), "###,##0.00") + " (kg)"; Tab(N2); TxtCant + txtFields(15) + " (und)"; Tab(N3); "ATD: " + Str(Contador_Atds) + "/" + Cont_Atds_Hasta
            'Printer.Print Tab(N1); TxtNorma + Norma; Tab(N2); TxtLote + Lote; Tab(N3); "P: " + Format(txtFields(17), "###,##0.00")
            Printer.Print Tab(N1); Norma; Tab(N2); TxtLote + Lote; Tab(N3); "P: " + Format(txtFields(17), "###,##0.00") + " (kg)"
        Else
            Printer.Print Tab(N1); TxtDiam + Format(txtFields(7), "###,##0.00") + " (mm)"; Tab(N2); TxtSep + txtFields(10) + " mm"; Tab(N3); "C: " + txtFields(11)
            Printer.Print Tab(N1); TxtPeso + Format(txtFields(8), "###,##0.00") + " (kg)"; Tab(N2); TxtCant + txtFields(11) + " (und)"; Tab(N3); "ATD: " + Str(Contador_Atds) + "/" + Cont_Atds_Hasta
            'Printer.Print Tab(N1); TxtNorma + Norma; Tab(N2); TxtLote + Lote; Tab(N3); "P: " + Format(txtFields(16), "###,##0.00") + " (kg)"
            Printer.Print Tab(N1); Norma; Tab(N2); TxtLote + Lote; Tab(N3); "P: " + Format(txtFields(16), "###,##0.00") + " (kg)"
        End If
        Contador = Contador + 1
        If (Contador = Punto_Ajuste) Then
            'Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas - 1
            Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas + 1
            Contador = 0
        End If
        For R = 1 To Lineas_Entre_Tarjetas
            Printer.Print
        Next R
   Next i
End Sub 'PRINT_MR_ESPECIAL (MR*)

'*------------------------------------------------------------
'* MALLAS ELECTROSOLDADAS EN ROLLOS (MR)
'*------------------------------------------------------------
Private Sub PRINT_MR(Norma As String)
   Dim Lote As String
   Dim Espaciametro As String
   'Punto_Ajuste = 4  '<: Valor for WNT
   Punto_Ajuste = 4   '<: Valor for W98
   Contador = 0
   Printer.Print
   Espaciametro = txtFields(10)
   If Len(Espaciametro) > 0 Then  ' es decir Not IsNull( Espaciametro ) <- esto no funciona.
        Espaciametro = Espaciametro + " (mm)"
   End If
   Designacion = "MALLA ELECTROSOLDADA EN ROLLO"
   For i = txtFields(2) To txtFields(3)
        Lineas_Entre_Tarjetas = 6 '<: Valor for WNT
        'Lineas_Entre_Tarjetas = 8 '<: Valor for W98
        Lote = Format(i, "00000") + "-" + txtFields(0) + "-" + Format(txtFields(1), "YY")
        Printer.Print Tab(N1); "DESIGNACION: " + Designacion; Tab(N3); Tab(N3); Lote
        Printer.Print Tab(N1); "TIPO: " + txtFields(5) + " (m2)"; Tab(N2); "PESO: " + Format(txtFields(8), "###,##0.00") + " (kg)"; Tab(N3); txtFields(5) + " m2"
        Printer.Print Tab(N1); "LONGITUD: " + Format(txtFields(9), "###,##0.00") + " (m)"; Tab(N2); "ANCHO: " + Format(txtFields(6), "###,##0.00") + " (m)"; Tab(N3); Espaciametro
        Printer.Print Tab(N1); "DIAM: " + Format(txtFields(7), "###,##0.00") + " (mm)"; Tab(N2); "ESPACIAM: " + Espaciametro
        Printer.Print Tab(N1); "NORMA: " + Norma; Tab(N2); "LOTE: " + Lote
        '*
        Contador = Contador + 1
        If (Contador = Punto_Ajuste) Then
            Lineas_Entre_Tarjetas = Lineas_Entre_Tarjetas + 1
            Contador = 0
        End If
        '*
        For R = 1 To Lineas_Entre_Tarjetas
            Printer.Print
        Next R
   Next i
End Sub
'****-----------------EOF(InvTarj01v2_FRM)------------------------------------***
