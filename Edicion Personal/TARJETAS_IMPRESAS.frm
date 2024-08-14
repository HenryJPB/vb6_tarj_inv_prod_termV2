VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TARJETAS_IMPRESAS 
   Caption         =   "TARJETAS IMPRESAS"
   ClientHeight    =   4245
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   4245
   ScaleWidth      =   7785
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7785
      TabIndex        =   7
      Top             =   3645
      Width           =   7785
      Begin VB.CommandButton Boton_Buscar 
         Caption         =   "&Buscar"
         Height          =   300
         Left            =   3120
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Boton_Seleccionar 
         Caption         =   "&Seleccionar"
         Height          =   300
         Left            =   4200
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   5280
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Reno&var"
         Height          =   300
         Left            =   2040
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
      ScaleWidth      =   7785
      TabIndex        =   1
      Top             =   3945
      Width           =   7785
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "TARJETAS_IMPRESAS.frx":0000
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
         Picture         =   "TARJETAS_IMPRESAS.frx":0342
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
         Picture         =   "TARJETAS_IMPRESAS.frx":0684
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
         Picture         =   "TARJETAS_IMPRESAS.frx":09C6
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
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
Attribute VB_Name = "TARJETAS_IMPRESAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-----------------------------------------------------------
'* NOMBRE LOGICO: TARJETAS_IMPRESAS.
'* NOMBRE FISICO: TARJETAS_IMPRESAS.
'* Microsoft Visual Studio ( Visual Basic v 6.0 )
'* Autor: Henry J. Pulgar B.
'* Creado en Diciembre 02, 2002.
'* Fue actualizado: Enero 28, 2003.
'*-----------------------------------------------------------
'Dim WithEvents adoPrimaryRS As Recordset
Public WithEvents adoPrimaryRS As Recordset
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
  'db.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;database=bd816;"
  db.Open "PROVIDER=MSDASQL;dsn=DESICA815;uid=ops$desinv02;pwd=ops$desinv02;"
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select C1_TIPO_PROD Tipo_Prod," & _
                           "C1_FECHA_TARJ Fecha_Tarjeta," & _
                           "C1_FECHA_ENTREGA Fecha_Entrega," & _
                           "C1_LOTE_ANT Desde_No," & _
                           "C1_LOTE_PROX Hasta_No," & _
                           "C1_ORDEN_FAB Orden_Fab," & _
                           "C1_TIPO Tipo," & _
                           "C1_ANCHO Ancho," & _
                           "C1_DIAMETRO Diametro," & _
                           "C1_PESO Peso," & _
                           "C1_PESO_AUX Peso_Ult_Atd," & _
                           "C1_PESO_ATADO Peso_Atado," & _
                           "C1_PESO_ATADO_AUX Peso_Ultimo_Atd," & _
                           "C1_LONGITUD Longitud," & _
                           "C1_ESPACIAM Espaciam," & _
                           "C1_CANTIDAD Cantidad," & _
                           "C1_CANTIDAD_AUX Cant_Ult_Atd," & _
                           "C1_MAQUINA Maquina," & _
                           "C1_TURNO Turno " & _
                    "from  INVTARJ01_DAT " & _
                    "where C1_TIPO_PROD != '00' " & _
                    "Order by C1_TIPO_PROD, C1_FECHA_TARJ", db, adOpenStatic, adLockOptimistic

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

'* Autor Henry J. Pulgar B.
Public Sub cmdUpdate()
  On Error GoTo UpdateErr

  adoPrimaryRS("Turno") = Val(Form4.Combo1.Text)
  If Not IsNull(Form4.Text1.Text) Or (Len(Form4.Text1.Text) > 0) Then
         adoPrimaryRS("Fecha_Entrega") = CDate(Form4.Text1.Text)
  Else
         '???adoPrimaryRS("Fecha_Entrega") = CDate(Null)
  End If
  adoPrimaryRS("Maquina") = Form4.Text2.Text
  Unload Form4
  adoPrimaryRS.UpdateBatch adAffectAll

  Exit Sub
UpdateErr:
  MsgBox Err.Description
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

Private Sub Boton_Seleccionar_Click()
  Form4.Show
  'Form4.Text1.Text = Date <- fuese necesario solo si este campo no estuviese
  '                           inicializado.
End Sub

Private Sub Boton_Buscar_Click()
   Dim SQLCriterio As String
   cmdRefresh_Click
   SQLCriterio = InputBox("Introdusca criterio de busqueda (Tipo Producto)?:", "BUSQUEDA", "MP", 30, 30)
   If SQLCriterio <> "" Then 'Input Box No Cancelado
      SQLCriterio = UCase(SQLCriterio)
      SQLCriterio = " TIPO_PROD like " + "'*" + SQLCriterio + "*'"
      'MsgBox SQLCriterio
      '* Buscar:
      adoPrimaryRS.Find (SQLCriterio)
   End If
End Sub

Private Sub grdDataGrid_DblClick()
   Form4.Show
   'Form4.Text1.Text = Date
End Sub
