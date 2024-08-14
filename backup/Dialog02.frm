VERSION 5.00
Begin VB.Form Dialog02 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir tarjetas para identificar Productos Terminados.   "
   ClientHeight    =   2175
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H008080FF&
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "En proceso impresion de tarjetas... Presione ACEPTAR para continuar ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "Dialog02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Caja de Dialogo: Dialog02; nombre fisico: Dialog02.
'* Autor : Henry J. Pulgar B.
'* Creado: el 22 de Enero de 2003.
'* Ultima fecha de actualizacion:  Enero 22, 2003.
'**********************************************************
Option Explicit

Private Sub CancelButton_Click()
   '* Cancelar impresion ???
   Unload Me
End Sub

Private Sub OKButton_Click()
   Unload Me
   Form3.Show
End Sub
