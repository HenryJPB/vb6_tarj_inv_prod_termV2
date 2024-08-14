VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Impresion Tarjetas de Inventario"
   ClientHeight    =   6465
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "MS Dialog"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tarjetas01_FRM.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Tarjetas01_FRM.frx":0442
   ScaleHeight     =   6465
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5160
      TabIndex        =   39
      Text            =   "03"
      Top             =   720
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   960
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=DESICA733"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DESICA733"
      OtherAttributes =   ""
      UserName        =   "OPS$DESINV02"
      Password        =   "OPS$DESINV02"
      RecordSource    =   "select C1_TIPO, C1_DESCRIPCION from INV01_DAT order by C1_TIPO, C1_DESCRIPCION"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Dialog"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=DESICA733"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DESICA733"
      OtherAttributes =   ""
      UserName        =   "ops$desica"
      Password        =   "hp150"
      RecordSource    =   "select NOMBRE_CLI_PROV ""Nombre Cliente"" from CXCD_DAT where   order by NOMBRE_CLI_PROV"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Dialog"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   38
      Text            =   "Combo2"
      Top             =   1680
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   37
      Text            =   "Combo1"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   5160
      TabIndex        =   36
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   1800
      TabIndex        =   35
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   5160
      TabIndex        =   28
      Text            =   "58,23"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1680
      TabIndex        =   27
      Text            =   "6,23"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   5160
      TabIndex        =   26
      Text            =   "20"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      Text            =   "10 x 10"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   5160
      TabIndex        =   24
      Text            =   "3,25"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1680
      TabIndex        =   23
      Text            =   "2,50"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5160
      TabIndex        =   22
      Text            =   "2,15"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Text            =   "6,50"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Text            =   "1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "1"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Text            =   "0002"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos Ultimo Atado"
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   6375
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   4920
         TabIndex        =   32
         Text            =   "52,67"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Text            =   "10"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "16.Peso/atado:"
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "15. Cant (uni"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5280
      Picture         =   "Tarjetas01_FRM.frx":F6508
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   3840
      Picture         =   "Tarjetas01_FRM.frx":F6A3A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   3240
      X2              =   3240
      Y1              =   600
      Y2              =   1090
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   3240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   3240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   2280
      X2              =   2040
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Label Label18 
      Caption         =   "18. Turno:"
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "17. Fecha:"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "14.Peso/atado"
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "13.Peso/malla:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "12. Cant.(unid):"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "11. Sep (mm):"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "10. Diam (mm)"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "9. Ancho (mm):"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "8. Diam (mm):"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "7. Largo (mm):"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "6. Tipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "5. Cliente:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "4. Maq. No.:"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "3. Atados:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "2. Orden No.:"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1. Ref No.:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form1"
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
'*  Nombre del Proyecto: Tarjetas.vbw
'*  Autor: Henry J. Pulgar B.
'*  Fecha de creacion: 24 de Abril de 2001
'*  Ultima fecha actualizacion: 01-08-2001
'*********************************************************

Private Sub Form_Load()
   Text17.Text = Date
   'MsgBox "Procedere a cargar los datos; este procedimiento tomara algun tiempo ...", vbOKOnly
   Cargar_Datos_Oracle
End Sub  ' Sub-Form Load ...

'Cargar datos Oracle v733 a traves de ODBC-OLE conneccion.
Private Sub Cargar_Datos_Oracle()
Dim i As Integer
   Dim StringSQL As String
   Dim TempVar As String
   Dim Conn As ADODB.Connection
   Dim ConnRec As ADODB.Recordset
   Dialog.Show
   'Set variables:
   Set Conn = New ADODB.Connection
   Set ConnRec = New ADODB.Recordset
   '------------------------------------------------------------------
   'Implantar coneccion: Para carga de datos CLIENTES: OPS$DESICA.
   '------------------------------------------------------------------
    'Conn.ConnectionString = "DSN=DESICA733;UID=OPS$DESICA;PWD=hp150"
    Conn.ConnectionString = "DSN=DESICA806;UID=OPS$DESICA;PWD=OPS$DESICA"
    Conn.Open
    StringSQL = "select NOMBRE_CLI_PROV from CXCD_DAT where TIPO_DE_CLIENTE = 'A' order by NOMBRE_CLI_PROV"
    ConnRec.Open StringSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdUnknown
    TempVar = ConnRec("NOMBRE_CLI_PROV")
    While (Not ConnRec.EOF)
          Combo1.AddItem ConnRec("NOMBRE_CLI_PROV")
          ConnRec.MoveNext
    Wend
    ConnRec.Close
    Conn.Close
    Combo1.Text = TempVar
   '-------------------------------------------------------------------------------
   'Implantar coneccion: Para carga de datos INVENTARIO DE PRODUCTOS: OPS$DESINV02.
   '-------------------------------------------------------------------------------
   Conn.ConnectionString = "DSN=DESICA806;UID=OPS$DESINV02;PWD=OPS$DESINV02"
   Conn.Open
   StringSQL = "select C1_TIPO,C1_DESCRIPCION from INV01_DAT order by C1_TIPO, C1_DESCRIPCION"
   ConnRec.Open StringSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdUnknown
   TempVar = ConnRec("C1_TIPO") & " " & ConnRec("C1_DESCRIPCION")
   'Display state coneccion:
   'MsgBox "Conn State=" & Conn.State
   'MsgBox "Conn Rec State =" & ConnRec.State
   i = 0
   While (Not ConnRec.EOF)
         Combo2.AddItem ConnRec("C1_TIPO") & " " & ConnRec("C1_DESCRIPCION")
         ConnRec.MoveNext
   Wend
   ConnRec.Close
   Conn.Close
   Combo2.Text = TempVar
Unload Dialog
End Sub  'Cargar datos ORACLE

Private Sub Command1_Click()
    Imprimir_Click
End Sub

Private Sub Command2_Click()
   Salir_Click
End Sub

Private Sub IMPRIMIR_TARJETAS(i As Integer, NumCopias As Integer)
   Dim J
   'NOTA: El uso de Tab(n) implica el valor absoluto desde el origen; coord (0,0).
   '      mientras que el uso de la funcion Spc(n) implica el valor relativo desde la ultima impresion
   Tab1 = 25
   Printer.FontSize = 9
   Printer.FontBold = False
   Printer.Print Tab(Tab1 + 20); Text2.Text; Tab(Tab1 + 39); i; "/"; Text4.Text
   Printer.Print 'Print Blank line
   Printer.Print Tab(Tab1); Mid(Combo1.Text, 1, 23); Tab(Tab1 + 38); Text5.Text
   Printer.Print 'Print Blank line
   Printer.Print Tab(Tab1); Combo2.Text   '<- old, "Text6.Text"
   Printer.Print 'Print Blank line
   Printer.Print Tab(Tab1 + 4); Text7.Text; Tab(Tab1 + 34); Text8.Text
   Printer.Print 'Print Blank line
   Printer.Print Tab(Tab1 + 4); Text9.Text; Tab(Tab1 + 34); Text10.Text
   Printer.Print 'Print Blank line
   If (i = NumCopias) Then 'Imprimir ultima tarjeta
      Printer.Print Tab(Tab1 + 4); Text11.Text; Tab(Tab1 + 34); Text15.Text
      Printer.Print 'Print Blank line
      Printer.Print Tab(Tab1 + 7); Text13.Text; Tab(Tab1 + 37); Text16.Text
   Else
      Printer.Print Tab(Tab1 + 4); Text11.Text; Tab(Tab1 + 34); Text12.Text
      Printer.Print 'Print Blank line
      Printer.Print Tab(Tab1 + 7); Text13.Text; Tab(Tab1 + 37); Text14.Text
   End If
   Printer.Print 'Print Blank line
   Printer.Print Tab(Tab1 + 25); Text17.Text
   For J = 1 To 6
       Printer.Print
   Next J
 End Sub
 

Private Sub Imprimir_Click()
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
  CommonDialog1.Copies = Val(Form1.Text5.Text)
  CommonDialog1.ShowPrinter
  
  ' Obtener las propiedades de impresión
  PrimeraPag = CommonDialog1.FromPage
  ÚltimaPag = CommonDialog1.ToPage
  NumCopias = CommonDialog1.Copies   '<- Esta instruccion no esta funcionando ????
  Desde = Val(Form1.Text3.Text)
  NumCopias = Val(Form1.Text4.Text)  ' Depende del numero de atados
  ImpArchivo = CommonDialog1.Flags And cdlPDPrintToFile
  ' Imprimir
  If ImpArchivo Then
    For i = Desde To NumCopias
      ' Escriba el código para enviar datos a un archivo
      GENERAR_ARCHIVO
    Next i
  Else  'Dirigir salida a la impresora
    T = NumCopias
    For i = Desde To NumCopias
         For J = 1 To 2   '1 Impresion=Cliente + 1 Impresion=Despacho
            ' Escriba el código para enviar datos a la impresora
            IMPRIMIR_TARJETAS (i), (NumCopias) '-->Parametros pasados por valor "()" -> la var no puede ser modificada
         Next J
    Next i
    Printer.EndDoc
  End If

SalirImprimir:
  Exit Sub

ManipularErrorImprimir:
  ' Manipular el error
  If Err.Number = cdlCancel Then Exit Sub
  MsgBox Err.Description
  Resume SalirImprimir
End Sub 'Imprimir_Click_

Private Sub Salir_Click()
   Dim Mensaje, Botones, Titulo, Respuesta
   Mensaje = "¿Deseas Salir?"
   Botones = vbYesNo + vbQuestion + vbDefault
   Titulo = "¿Salir?"
   Respuesta = MsgBox(Mensaje, Botones, Titulo)
   If (Respuesta = vbYes) Then
       Unload Form1
   End If
End Sub 'Sub Salir_Click

Private Sub GENERAR_ARCHIVO()
   Open "e:\Vb5\Proyectos\Tarjetas\reporte.lst" For Output As #1
   Print #1, Text1.Text, Text2.Text
   Print #1, Text3.Text, Text4.Text
   Print #1, Text5.Text
   Print #1, Text6.Text
   'Datos del frame:
   Print #1, Text15.Text, Text16.Text
   Close #1
End Sub 'Sub GENERAR_ARCHIVO a disco en formato ASCII

'------------------------
' Atado Numero.
'------------------------
Private Sub Text3_LostFocus()
   Dim Mensaje
   If Len(Text3.Text) <> 0 Then
   If Not IsNumeric(Text3.Text) Then
     Mensaje = Text3.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text3.Text = Format(Text3.Text, "###,###,##0")
   End If
   End If
End Sub

'-------------
'Total Atados:
'-------------
Private Sub Text4_LostFocus()
 Dim Mensaje
   If Len(Text4.Text) <> 0 Then
   If Not IsNumeric(Text4.Text) Then
     Mensaje = Tex4.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text4.Text = Format(Text4.Text, "###,###,##0")
     If Text4.Text < Text3.Text Then
         Mensaje = "Total atados no puede ser menor al numero de atado."
         MsgBox Mensaje, vbCritical, "Atencion!"
     End If
   End If
   End If
End Sub

'----------------------------
'Largo (mm):
'----------------------------
Private Sub Text7_LostFocus()
   Dim Mensaje
   If Len(Text7.Text) <> 0 Then
   If Not IsNumeric(Text7.Text) Then
     Mensaje = Text7.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text7.Text = Format(Text7.Text, "###,###,##0.00")
   End If
   End If
End Sub

'---------------------------
' Diam longitudinal:
'---------------------------
Private Sub Text8_LostFocus()
   Dim Mensaje
   If Len(Text8.Text) <> 0 Then
   If Not IsNumeric(Text8.Text) Then
     Mensaje = Text8.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text8.Text = Format(Text8.Text, "###,###,##0.00")
   End If
   End If
End Sub

'----------------------------
' Ancho (mm):
'----------------------------
Private Sub Text9_LostFocus()
   Dim Mensaje
   If Len(Text9.Text) <> 0 Then
   If Not IsNumeric(Text9.Text) Then
     Mensaje = Text9.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text9.Text = Format(Text9.Text, "###,###,##0.00")
   End If
   End If
End Sub

'-----------------------------
' Diam (mm)
'-----------------------------
Private Sub Text10_LostFocus()
   Dim Mensaje
   If Len(Text10.Text) <> 0 Then
   If Not IsNumeric(Text10.Text) Then
     Mensaje = Text10.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text10.Text = Format(Text10.Text, "###,###,##0.00")
   End If
   End If
End Sub



'-----------------------------
' Cantidad de items:
'-----------------------------
Private Sub Text12_LostFocus()
   Dim Mensaje
   If Len(Text12.Text) <> 0 Then
   If Not IsNumeric(Text12.Text) Then
     Mensaje = Text12.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text12.Text = Format(Text12.Text, "###,###,##0")
   End If
   Text14.Text = Text12.Text * Text13.Text
   End If
End Sub

'-----------------------------
' Peso/item
'-----------------------------
Private Sub Text13_LostFocus()
   Dim Mensaje
   If Len(Text13.Text) <> 0 Then
   If Not IsNumeric(Text13.Text) Then
     Mensaje = Text13.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text13.Text = Format(Text13.Text, "###,###,##0.00")
   End If
   Text14.Text = Text12.Text * Text13.Text
   End If
End Sub

'-----------------------------
' Peso/atado:
'-----------------------------
Private Sub Text14_LostFocus()
   Dim Mensaje
   If Len(Text14.Text) <> 0 Then
   If Not IsNumeric(Text14.Text) Then
     Mensaje = Text14.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text14.Text = Format(Text14.Text, "###,###,##0.00")
   End If
   'Text14.Text = Text12.Text * Text13.TextText14.Text = Text12.Text * Text13.Text
   End If
End Sub

'------------------------------
' Cantidad items ultimo atado
'------------------------------
Private Sub Text15_LostFocus()
   Dim Mensaje
   If Len(Text15.Text) <> 0 Then
   If Not IsNumeric(Text15.Text) Then
     Mensaje = Text15.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text15.Text = Format(Text15.Text, "###,###,##0")
   End If
     Text16.Text = Text15.Text * Text13.Text
   End If
End Sub

'-----------------------------
' Peso ultimo atado:
'-----------------------------
Private Sub Text16_LostFocus()
   Dim Mensaje
   If Len(Text16.Text) <> 0 Then
   If Not IsNumeric(Text16.Text) Then
     Mensaje = Text16.Text & ": Valor numerico incorrecto"
     MsgBox Mensaje, vbCritical, "Atencion!"
   Else
     Text16.Text = Format(Text16.Text, "###,###,##0.00")
   End If
   End If
End Sub

