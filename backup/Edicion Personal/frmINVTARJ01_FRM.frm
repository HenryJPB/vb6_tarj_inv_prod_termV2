VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form INVTARJ01_FRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ELABORAR TARJETAS de INVENTARIO"
   ClientHeight    =   6285
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6465
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Boton_Imprimir 
      Caption         =   "Conforme"
      Height          =   495
      Left            =   2160
      Picture         =   "frmINVTARJ01_FRM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_TIPO"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   27
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton Boton_Cancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3120
      Picture         =   "frmINVTARJ01_FRM.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5280
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
      ScaleWidth      =   6465
      TabIndex        =   24
      Top             =   5685
      Width           =   6465
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   6480
         TabIndex        =   25
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
      ScaleWidth      =   6465
      TabIndex        =   23
      Top             =   5985
      Width           =   6465
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_CANTIDAD"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   22
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_ESPACIAM"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   20
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_LONGITUD"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   18
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_PESO"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   16
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_DIAMETRO"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_ANCHO"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C1_ORDEN_FAB"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      DataField       =   "C1_LOTE_PROX"
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      DataField       =   "C1_LOTE_ANT"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      DataField       =   "C1_FECHA_TARJ"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      DataField       =   "C1_TIPO_PROD"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "DATOS DE LA TARJETA"
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
      Left            =   1920
      TabIndex        =   28
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cantidad de Items?:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Espaciam?:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Longitud?:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Peso?:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Diametro?:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ancho?:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Caracteristicas del prod?:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Orden de Fab. No?:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Al "
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Imprimir Tarjetas, Desde:  "
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo de Tarjeta:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "INVTARJ01_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-------------------------------------------------------------
'*  Actualizar datos para impresion de Tarjetas
'*  de Inventario de Productos
'*  Elaborado x Wizard de formas VB 6.0 basado en "Codigo ADO".
'*  Actualizado y personalizado x Henry J. Pulgar B.
'*  Creado el 20 de Agosto de 2002.
'*  Actualizado el 29 de Agosto de 2002.
'*--------------------------------------------------------------

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
  db.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
   
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select C1_TIPO_PROD,C1_FECHA_TARJ,C1_LOTE_ANT,C1_LOTE_PROX,C1_ORDEN_FAB,C1_TIPO,C1_ANCHO,C1_DIAMETRO,C1_PESO,C1_LONGITUD,C1_ESPACIAM,C1_CANTIDAD from INVTARJ01_DAT Order by C1_TIPO_PROD, C1_LOTE_PROX", db, adOpenStatic, adLockOptimistic
     
  'Enlaza los cuadros de texto con el proveedor de datos
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
  
  '* Coloque aqui sus instrucciones
  INICIALIZAR_VALORES
End Sub

Private Sub INICIALIZAR_VALORES()
   Dim Mensaje
   '* Ejecutar un segundo query; Ubicar ultima tarjeta impresa.
   Dim Coneccion2 As Connection
   Set Coneccion2 = New Connection
   Coneccion2.CursorLocation = adUseClient
   Coneccion2.Open "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
   Set Q1 = New Recordset
   CadenaSQL = "select C1_LOTE_ANT, C1_LOTE_PROX " & _
               "from   INVTARJ01_DAT " & _
               "where  C1_TIPO_PROD = '" & adoPrimaryRS("C1_TIPO_PROD") & "'" & _
               "and    C1_LOTE_PROX = ( select MAX( C1_LOTE_PROX ) " & _
                                        "from  INVTARJ01_DAT " & _
                                        "where  C1_TIPO_PROD = '" & adoPrimaryRS("C1_TIPO_PROD") & "' )"
   'MsgBox CadenaSQL
   Q1.Open CadenaSQL, Coneccion2, adOpenStatic, adLockOptimistic
   If (Q1.EOF) Then
      ProximaTarjeta = 0
      'MsgBox Mensaje, vbCritical, "Atencion!"
   Else
      ProximaTarjeta = Q1("C1_LOTE_PROX")
   End If
   '---
   adoPrimaryRS("C1_TIPO_PROD") = Form2.Combo1.Text
   adoPrimaryRS("C1_FECHA_TARJ") = Form2.Text1.Text
   adoPrimaryRS("C1_LOTE_ANT") = ProximaTarjeta + 1
   adoPrimaryRS("C1_LOTE_PROX") = ProximaTarjeta + Form2.Text2.Text
   Select Case Form2.Combo1.Text
          Case "AT"
               adoPrimaryRS("C1_TIPO") = "Con Resaltes"
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

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
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

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

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
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub Boton_Cancelar_Click()
   Unload Me
End Sub

