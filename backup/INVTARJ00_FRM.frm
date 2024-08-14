VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form INVTARJ00_FRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVTARJ00_DAT"
   ClientHeight    =   3285
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5955
   Begin VB.CommandButton Bottom 
      Caption         =   ">|"
      Height          =   290
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Next_Reg 
      Caption         =   ">"
      Height          =   290
      Left            =   1440
      TabIndex        =   9
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Back_Reg 
      Caption         =   "<"
      Height          =   290
      Left            =   960
      TabIndex        =   8
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Tope 
      Caption         =   "|<"
      Height          =   290
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   2985
      Width           =   5955
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar/Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "A&ctualizar/Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "C0_NORMA"
      DataSource      =   "datPrimaryRS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00FFFFC0&
      DataField       =   "C0_TIPO_PROD"
      DataSource      =   "datPrimaryRS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   2655
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=DESICA816;uid=ops$desinv02;pwd=ops$desinv02;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select C0_TIPO_PROD,C0_NORMA from INVTARJ00_DAT order by C0_TIPO_PROD"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Registro No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "NORMA ASOCIADA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TIPO_PROD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "INVTARJ00_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*----------------------------------------------------
'*  Actualizar tabla de Normas/Tipo de Prod.
'*  Elaborado x Wizard de formas VB 6.0 basado en "Control de Datos".
'*  Actualizado y personalizado x Henry J. Pulgar B.
'*  Creado el 19 de Agosto del año 2002.
'*  Actualizado el 22 de Mayo del año 2003.
'*-----------------------------------------------------

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Aquí es donde puede colocar el código de control de errores
  'Si desea pasar por alto los errores, marque como comentario la siguiente línea
  'Si desea detectarlos, agregue código aquí para controlarlos
  MsgBox "Data error event hit err:" & Description
End Sub
'Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  adoPrimaryRS.Caption = "Registro No. " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
  Label2.Caption = CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'NOTA I: originalmente asi. Programa basado en componete ADO ( nor codigo ADO ):
'Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'            ***
'* Sin embargo y por razones desconocidas dejo de funcionar ???-> ERROR de compilacion.
'* Se solventó cambiando adoPrimaryRS (...) x datPrimaryRS( ... ); pero se sacrifican los
'  objetos: "adoPrimaryRS.Caption = "Registro No. " & CStr(datPrimaryRS.AbsolutePosition)
' Label2.Caption = CStr( ... )" del procedimiento anterior...
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  If Conforme = vbYes Then
     With datPrimaryRS.Recordset
         .Delete
         .MoveNext
     If .EOF Then .MoveLast
     End With
  Else
     cmdRefresh_Click
  End If
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  '...............................................
  'Dummy centinels tips:
  'datPrimaryRS.Recordset("C0_NORMA") = "Henry"
  'TempStr =datPrimaryRS.Recordset("C0_NORMA")
  '...............................................
  If Conforme() = vbYes Then
     datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Else
     cmdRefresh_Click
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Function Conforme()
   Dim Mensaje, Botones, Titulo, Respuesta
   Mensaje = "Conforme?"
   Botones = vbYesNo + vbDefaultButton2
   Titulo = "Conforme"
   Conforme = MsgBox(Mensaje, Botones, Titulo)
End Function


Private Sub Tope_Click()
    With datPrimaryRS.Recordset
         .MoveFirst
    End With
End Sub

'Derechos de autor
Private Sub Back_Reg_Click()
 With datPrimaryRS.Recordset
      If .BOF Then
         .MoveLast
      Else
         .MovePrevious
      End If
 End With
End Sub

'Copirigth de Visual B.
Private Sub Back_Reg_Click_old()
     On Error GoTo GoPrevError

  With datPrimaryRS.Recordset
     If Not .BOF Then .MovePrevious
     If .BOF And .RecordCount > 0 Then
       Beep
       'ha sobrepasado el final; vuelva atrás
       .MoveFirst
      End If
     'muestra el registro actual
      mbDataChanged = False
  End With
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub Next_Reg_Click()
 With datPrimaryRS.Recordset
      If .EOF Then
         .MoveFirst
      Else
         .MoveNext
      End If
 End With
End Sub


Private Sub Bottom_Click()
 With datPrimaryRS.Recordset
        .MoveLast
 End With
End Sub
