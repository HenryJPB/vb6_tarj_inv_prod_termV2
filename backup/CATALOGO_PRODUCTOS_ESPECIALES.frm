VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CATALOGO_PRODUCTOS_ESPECIALES 
   Caption         =   "Catalogo de Productos Especiales"
   ClientHeight    =   2520
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   2520
   ScaleWidth      =   6615
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6615
      TabIndex        =   7
      Top             =   1920
      Width           =   6615
      Begin VB.CommandButton Boton_Selec 
         Caption         =   "&Seleccionar"
         Height          =   300
         Left            =   3000
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton BOTON_BUSCAR 
         Caption         =   "&Buscar?"
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4080
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   300
         Left            =   720
         TabIndex        =   8
         Top             =   0
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
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   2220
      Width           =   6615
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4560
         Picture         =   "CATALOGO_PRODUCTOS_ESPECIALES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "CATALOGO_PRODUCTOS_ESPECIALES.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "CATALOGO_PRODUCTOS_ESPECIALES.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "CATALOGO_PRODUCTOS_ESPECIALES.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   3360
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CATALOGO_PRODUCTOS_ESPECIALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'* CATALOGO DE PRODUCTOS ESPECIALES
'* Nombre Logico: CATALOGO_PRODUCTOS_ESPECIALES
'* Nombre fisico: CATALOGO_PRODUCTOS_ESPECIALES
'* Autor: Henry J Pulgar
'* Creado el Febrero 12, 2003.
'* Modificado el :
'* Creado x Wizard de Formularios Visual Basic 6.00
'* ( Acceso a datos remoto-ODBC )
'****************************************************************

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  'db.Open "PROVIDER=MSDASQL;dsn=DESICA733;uid=ops$desinv02;pwd=ops$desinv02;"
  
  db.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$despro03;pwd=ops$despro03;"
  
  Set adoPrimaryRS = New Recordset
  No_Orden = iNVTARJ01v2_FRM.txtFields(4)
  adoPrimaryRS.Open "select C4_ORDEN NO_ORDEN," & _
                    "C4_CODIGO COD_PROD," & _
                    "C4_TIPO_PROD TIPO_PROD," & _
                    "TO_NUMBER(C4_TAML) TAML," & _
                    "TO_NUMBER(C4_TAMT) TAMT," & _
                    "TO_NUMBER(C4_SEPL) SEPL," & _
                    "TO_NUMBER(C4_SEPT) SEPT," & _
                    "TO_NUMBER(C4_DIAML) DIAML," & _
                    "TO_NUMBER(C4_DIAMT) DIAMT," & _
                    "TO_NUMBER(C4_NUML) NUML," & _
                    "TO_NUMBER(C4_NUMT) NUMT," & _
                    "TO_NUMBER(C4_SOBRL1) SOBRL1," & _
                    "TO_NUMBER(C4_SOBRL2) SOBRL2," & _
                    "TO_NUMBER(C4_SOBRT1) SOBRT1," & _
                    "TO_NUMBER(C4_SOBRT2) SOBRT2," & _
                    "TO_NUMBER(C4_LAM) CANTIDAD," & _
                    "TO_NUMBER(C4_BSXTON) BSxTON " & _
                    "from INV04_DAT " & _
                    "where C4_ORDEN = '" & No_Orden & "' " & _
                    "Order by C4_CODIGO", db, adOpenStatic, adLockOptimistic

  Set grdDataGrid.DataSource = adoPrimaryRS

  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario
  grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
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

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub
'* Boton seleccionar la fila marcada
Private Sub Boton_Selec_Click()
        grdDataGrid_DblClick
End Sub


'* Boton Buscar
Private Sub Boton_Buscar_Click()
   Dim SQLCriterio As String
   cmdRefresh_Click
   SQLCriterio = InputBox("Introdusca criterio de busqueda (Codigo Producto)?:", "BUSQUEDA", "", 30, 30)
   If SQLCriterio <> "" Then 'Input Box No Cancelado
      SQLCriterio = UCase(SQLCriterio)
      SQLCriterio = " COD_PROD like " + "'*" + SQLCriterio + "*'"
      'MsgBox SQLCriterio
      '* Buscar:
      adoPrimaryRS.Find (SQLCriterio)
   End If
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
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

'* El usuario presiona click en una de las filas.
Private Sub grdDataGrid_DblClick()
   'Aqui instrucciones para seleccionar los datos de la fila ...
   Codigo = adoPrimaryRS("COD_PROD")
   iNVTARJ01v2_FRM.txtFields(5) = Codigo
   iNVTARJ01v2_FRM.VALIDAR_ITEM_ORDEN
   '**Fin:
   Unload Me
End Sub

'**--------EOF(CATALOGO_PRODUCTOS_ESPECIALES)--------------
