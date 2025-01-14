VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarCode128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'////////////////////////////////////////////////////////////////////
'//
'// Gestionar el BarCode 128
'//
'// Dise�ado por Sergio Fustero Carreras (2006)
'//
'// Crea e imprime c�digo de barras 128 (tipos A y B)
'//     El tipo C est� pendiente de implementar
'// Este c�digo de barras se usa, por ejemplo, en las
'//     cartas certificadas, listas de correos, acuses de
'//     recibo, etc. de Correos (Espa�a)
'//
'// Puede imprimir directamente a impresora o a un
'//     contexto de dispositivo (p.e. PictureBox)
'//
'////////////////////////////////////////////////////////////////////

'variables locales para almacenar los valores de las propiedades
Private mvarTabla As colC128 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarFuente As String 'copia local

Public Enum eTipoDeC�digo128
    cC128_A = 1
    cC128_B = 2
    cC128_C = 3
End Enum

Public Function GenerarBarras(Texto As String, DC As Object, Optional Codificaci�n As eTipoDeC�digo128 = cC128_B, Optional X As Single = -1, Optional Y As Single = -1, Optional Densidad As Single = 15, Optional Alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tama�o As Integer = 4) As Boolean
'------------------------------
' Genera las barras (secuencias de l�neas negras y espacios en blanco)
'       en c�digo 128 del tipo A, B � C para un contexto dado.
' El contexto puede ser un PictureBox, p.e.
'
' Par�metros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- DC:
'       Contexto donde se imprimir�. Puede ser util para los que us�is vista previa.
'       Para imprimir directamente a la impresora, usa la funci�n GenerarBarrasImpresora
'           (puede generar barras mucho m�s peque�as)
'   -- Codificaci�n:
'       Indicar cC128_A, cC128_B � cC128_C
'   -- X:
'       Coordenada horizontal
'   -- Y:
'       Coordenada vertical
'   -- Densidad:
'       Ancho de las barras (no todos los anchos funcionan bien)
'   -- Alto:
'       Alto de las barras
'   -- ImprimirTexto:
'       Indicamos si queremos que se imprima el texto debajo de las barras o no
'   -- Fuente:
'       Fuente para el texto (si se quiere imprimir)
'   -- Tama�o:
'       Tama�o de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aqu� As Single, CBarras As String, C As Integer, D As Integer, Paso, Cx, Ancho
'DC.CurrentX = Densidad
' -- Establece el ancho de las l�neas
On Error Resume Next
DC.DrawWidth = 1    'Algunos contextos dan error
On Error GoTo 0
Ancho = Densidad
' -- Establecer posici�n vertical (fija durante todo el proceso)
If Y = -1 Then
    Y = DC.CurrentY
End If
' -- Establecer posici�n horizontal inicial
If X = -1 Then
    Aqu� = DC.CurrentX
Else
    Aqu� = X
End If
' -- Espacio en blanco antes del c�digo de barras
Aqu� = Aqu� + 10 * Densidad
' -- Conseguir el c�digo de barras
CBarras = Barras(Texto, Codificaci�n)
' -- Imprimir el c�digo de barras (se intercalan l�neas y espacios)
For C = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    For D = 1 To Val(Mid(CBarras, C, 1))
        If (C Mod 2) = 0 Then ' 0 = Falso --> L�nea
                                     ' 1 = Verdadero --> Espacio
            DC.Line (Aqu�, Y)-(Aqu�, Y + Alto), vbWhite
        Else
            DC.Line (Aqu�, Y)-(Aqu�, Y + Alto), vbBlack
        End If
        Aqu� = Aqu� + Ancho
    Next
Next
For C = 1 To 10 'Espacio en blanco despu�s
    DC.Line (Aqu�, Y)-(Aqu�, Y + Alto), vbWhite
    Aqu� = Aqu� + Ancho
Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    DC.Font.Name = Fuente
    DC.Font.Size = Tama�o
    DC.CurrentY = Y + Alto
    DC.CurrentX = Cx + (10 * Densidad) + (Aqu� - Cx - DC.TextWidth(Texto)) / 2
    DC.Print Texto
End If
End Function

Public Function GenerarBarrasImpresora(Texto As String, Optional Codificaci�n As eTipoDeC�digo128 = cC128_B, Optional X As Single = -1, Optional Y As Single = -1, Optional Densidad As Single = 10, Optional Alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tama�o As Integer = 4) As Boolean
'------------------------------
' Genera las barras (secuencias de l�neas negras y espacios en blanco)
'       en c�digo 128 del tipo A, B � C directamente en la impresora.
'
' Par�metros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- Codificaci�n:
'       Indicar cC128_A, cC128_B � cC128_C
'   -- X:
'       Coordenada horizontal
'   -- Y:
'       Coordenada vertical
'   -- Densidad:
'       Ancho de las barras (no todos los anchos funcionan bien)
'   -- Alto:
'       Alto de las barras
'   -- ImprimirTexto:
'       Indicamos si queremos que se imprima el texto debajo de las barras o no
'   -- Fuente:
'       Fuente para el texto (si se quiere imprimir)
'   -- Tama�o:
'       Tama�o de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aqu� As Single, CBarras As String, C As Integer, D As Integer, Paso, Cx, Ancho
' -- Establece el ancho de las l�neas
Printer.DrawWidth = 5
Ancho = Densidad
' -- Establecer posici�n vertical (fija durante todo el proceso)
If Y = -1 Then
    Y = Printer.CurrentY
End If
' -- Establecer posici�n horizontal inicial
If X = -1 Then
    Aqu� = Printer.CurrentX
Else
    Aqu� = X
End If
' -- Espacio en blanco antes del c�digo de barras
Aqu� = Aqu� + 10 * Densidad
' -- Conseguir el c�digo de barras
CBarras = Barras(Texto, Codificaci�n)
' -- Imprimir el c�digo de barras (se intercalan l�neas y espacios)
For C = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    For D = 1 To Val(Mid(CBarras, C, 1))
        If (C Mod 2) = 0 Then ' 0 = Falso --> L�nea
                                     ' 1 = Verdadero --> Espacio
            Printer.Line (Aqu�, Y)-(Aqu�, Y + Alto), vbWhite
        Else
            Printer.Line (Aqu�, Y)-(Aqu�, Y + Alto), vbBlack
        End If
        Aqu� = Aqu� + Ancho
    Next
Next
' -- Si es necesario, pondremos espacio en blanco despu�s
'For C = 1 To 10 'Espacio en blanco despu�s
'    Printer.Line (Aqu�, Y)-(Aqu�, Y + Alto), vbWhite
'    Aqu� = Aqu� + Ancho
'Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    Printer.Font.Name = Fuente
    Printer.Font.Size = Tama�o
    Printer.CurrentY = Y + Alto
    Printer.CurrentX = Cx + (10 * Densidad) + (Aqu� - Cx - Printer.TextWidth(Texto)) / 2
    Printer.Print Texto
End If
End Function


Public Function Barras(Texto As String, Optional Codificaci�n As eTipoDeC�digo128 = cC128_B) As String
'------------------------------
' Crea la secuencia de barras y espacios para el texto dado
'
' Falta de implementar el c�digo para el tipo C
'       (solamente admite n�meros y es mucho m�s reducido)
'------------------------------
Dim C As Integer, Letra As Integer, Suma As Long
' -- Valor inicial
Select Case Codificaci�n
    Case eTipoDeC�digo128.cC128_A
        Barras = mvarTabla(103 + 1).Patr�n
        Suma = 103
    Case eTipoDeC�digo128.cC128_B
        Barras = mvarTabla(104 + 1).Patr�n
        Suma = 104
    Case eTipoDeC�digo128.cC128_C
        Barras = mvarTabla(105 + 1).Patr�n
        Suma = 105
End Select
' -- Secuencia del texto
If Codificaci�n = cC128_C Then
Else
    For C = 1 To Len(Texto)
        For Letra = 1 To 63
            If mvarTabla(Letra).C�digo = Mid(Texto, C, 1) Then
                Barras = Barras & mvarTabla(Letra).Patr�n
                Suma = Suma + mvarTabla(Letra).Valor * C
                Exit For
            End If
        Next
        If Codificaci�n = cC128_A Then
            For Letra = 0 To 31
                If Letra = Asc(Mid(Texto, C, 1)) Then
                    Barras = Barras & mvarTabla(Letra).Patr�n
                    Suma = Suma + mvarTabla(Letra).Valor * C
                    Exit For
                End If
            Next
        Else
            For Letra = 64 To 94
                If mvarTabla(Letra).C�digo = Mid(Texto, C, 1) Then
                    Barras = Barras & mvarTabla(Letra).Patr�n
                    Suma = Suma + mvarTabla(Letra).Valor * C
                    Exit For
                End If
            Next
        End If
        For Letra = 95 To 102
            If Asc(Mid(Texto, C, 1)) = Letra + 32 Then
                Barras = Barras & mvarTabla(Letra).Patr�n
                Suma = Suma + mvarTabla(Letra).Valor * C
                Exit For
            End If
        Next
    Next
End If
' -- C�digo de control
Barras = Barras & mvarTabla((Suma Mod 103) + 1).Patr�n
' -- Final
Barras = Barras & mvarTabla(106 + 1).Patr�n
End Function


Public Function PonerFuente(Objeto As Object) As Boolean
'---------------------------
' Asigna la fuente (que se debe haber indicado con anterioridad
'       al objeto pasado como par�metro, si la tiene
'---------------------------
Dim C As Integer
PonerFuente = False
On Error Resume Next
For C = 0 To Objeto.FontCount - 1
    If UCase(Trim(mvarFuente)) = UCase(Trim(Objeto.Fonts(C))) Then
        Objeto.Font.Name = mvarFuente
        PonerFuente = True
        Exit For
    End If
Next
If Err.Number <> 0 Then
    PonerFuente = False
End If
On Error GoTo 0
End Function

Public Function AsignarFuente(Fuente As String) As Boolean
Dim C As Integer
'------------------------------
' Asigna la fuente para escribir el dato debajo de las barras
'       (si se quiere)
'------------------------------
AsignarFuente = False
'Recorremos todas las fuentes de pantalla
'       para saber si existe la fuente indicada
For C = 0 To Screen.FontCount - 1
    If UCase(Trim(Screen.Fonts(C))) = UCase(Trim(Fuente)) Then
        AsignarFuente = True
        mvarFuente = Fuente
        Exit For
    End If
Next
End Function

Public Property Get Fuente() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Fuente
    Fuente = mvarFuente
End Property

Public Function Texto(Frase As String, Optional Clase As eTipoDeC�digo128 = cC128_A) As String
'*******************************************************
'Devuelve el texto correspondiente para usarlo
'   con la fuente de c�digo de barras
'
'*******************************************************
Dim Suma As Long, C As Integer, V As Integer, Resto As Integer
' -- Valor de inicio, seg�n tipo
Select Case Clase
    Case cC128_A
        Suma = 203
    Case cC128_B
        Suma = 204
    Case cC128_C
        Suma = 205
End Select
Texto = Chr(Suma)
' -- C�lculo de la suma a codificar
If Clase = cC128_C Then
    For C = 1 To Len(Frase) Step 2
        V = Val(Mid(Frase, C, 2))
        Texto = Texto & mvarTabla(V).C�digo
        Suma = Suma + V * C
    Next
Else
    For C = 1 To Len(Frase)
        Select Case Clase
            Case eTipoDeC�digo128.cC128_A
                V = Valor(Mid(Frase, C, 1))
                If V < 64 Then
                    Texto = Texto & Mid(Frase, C, 1)
                End If
            Case eTipoDeC�digo128.cC128_B
                V = Valor(Mid(Frase, C, 1))
                If V < 95 Then
                    Texto = Texto & Mid(Frase, C, 1)
                End If
            Case eTipoDeC�digo128.cC128_C
        End Select
        Suma = Suma + V * C
    Next
End If
' -- D�gito de control
Resto = Suma Mod 103
Texto = Texto + Chr(Resto)
' -- C�digo final
Texto = Texto & Chr(106)
End Function

Private Function Patr�n(Texto As String) As String
'---------------------------
' Devuelve el patr�n para un car�cter dado
'---------------------------
Dim V As C128
Patr�n = ""
For Each V In mvarTabla
    If V.C�digo = Texto Then
        Patr�n = V.Patr�n
        Exit For
    End If
Next
End Function

Private Function Valor(Texto As String) As Integer
'---------------------------
' Devuelve el valor num�rico correspondiente
'       a un car�cter dado
'---------------------------
Dim V As C128
Valor = -1
For Each V In mvarTabla
    If V.C�digo = Texto Then
        Valor = V.Valor
        Exit For
    End If
Next
End Function

Public Sub InicializarTabla()
'---------------------------
' Crea la tabla con los patrones para cada s�mbolo
'
' El patr�n indica, alternativamente, l�neas y espacios
'---------------------------
Dim C As C128
Set C = mvarTabla.Add("A1")
C.Valor = 0
C.C�digo = " "
C.Patr�n = "212222"
Set C = mvarTabla.Add("A2")
C.Valor = 1
C.C�digo = "!"
C.Patr�n = "222122"
Set C = mvarTabla.Add("A3")
C.Valor = 2
C.C�digo = Chr(34)  ' Comillas dobles
C.Patr�n = "222221"
Set C = mvarTabla.Add("A4")
C.Valor = 3
C.C�digo = "#"
C.Patr�n = "121223"
Set C = mvarTabla.Add("A5")
C.Valor = 4
C.C�digo = "$"
C.Patr�n = "121322"
Set C = mvarTabla.Add("A6")
C.Valor = 5
C.C�digo = "%"
C.Patr�n = "131222"
Set C = mvarTabla.Add("A7")
C.Valor = 6
C.C�digo = "&"
C.Patr�n = "122213"
Set C = mvarTabla.Add("A8")
C.Valor = 7
C.C�digo = "'"
C.Patr�n = "122312"
Set C = mvarTabla.Add("A9")
C.Valor = 8
C.C�digo = "("
C.Patr�n = "132212"
Set C = mvarTabla.Add("A10")
C.Valor = 9
C.C�digo = ")"
C.Patr�n = "221213"
Set C = mvarTabla.Add("A11")
'Set C = New C128
C.Valor = 10
C.C�digo = "*"
C.Patr�n = "221312"
Set C = mvarTabla.Add("A12")
C.Valor = 11
C.C�digo = "+"
C.Patr�n = "231212"
Set C = mvarTabla.Add("A13")
C.Valor = 12
C.C�digo = ","
C.Patr�n = "112232"
Set C = mvarTabla.Add("A14")
C.Valor = 13
C.C�digo = "-"
C.Patr�n = "122132"
Set C = mvarTabla.Add("A15")
C.Valor = 14
C.C�digo = "."
C.Patr�n = "122231"
Set C = mvarTabla.Add("A16")
C.Valor = 15
C.C�digo = "/"
C.Patr�n = "113222"
Set C = mvarTabla.Add("A17")
C.Valor = 16
C.C�digo = "0"
C.Patr�n = "123122"
Set C = mvarTabla.Add("A18")
C.Valor = 17
C.C�digo = "1"
C.Patr�n = "123221"
Set C = mvarTabla.Add("A19")
C.Valor = 18
C.C�digo = "2"
C.Patr�n = "223211"
Set C = mvarTabla.Add("A20")
C.Valor = 19
C.C�digo = "3"
C.Patr�n = "221132"
Set C = mvarTabla.Add("A21")
'Set C = New C128
C.Valor = 20
C.C�digo = "4"
C.Patr�n = "221231"
Set C = mvarTabla.Add("A22")
C.Valor = 21
C.C�digo = "5"
C.Patr�n = "213212"
Set C = mvarTabla.Add("A23")
C.Valor = 22
C.C�digo = "6"
C.Patr�n = "223112"
Set C = mvarTabla.Add("A24")
C.Valor = 23
C.C�digo = "7"
C.Patr�n = "312131"
Set C = mvarTabla.Add("A25")
C.Valor = 24
C.C�digo = "8"
C.Patr�n = "311222"
Set C = mvarTabla.Add("A26")
C.Valor = 25
C.C�digo = "9"
C.Patr�n = "321122"
Set C = mvarTabla.Add("A27")
C.Valor = 26
C.C�digo = ":"
C.Patr�n = "321221"
Set C = mvarTabla.Add("A28")
C.Valor = 27
C.C�digo = ";"
C.Patr�n = "312212"
Set C = mvarTabla.Add("A29")
C.Valor = 28
C.C�digo = "<"
C.Patr�n = "322112"
Set C = mvarTabla.Add("A30")
C.Valor = 29
C.C�digo = "="
C.Patr�n = "322211"
Set C = mvarTabla.Add("A31")
'Set C = New C128
C.Valor = 30
C.C�digo = ">"
C.Patr�n = "212123"
Set C = mvarTabla.Add("A32")
C.Valor = 31
C.C�digo = "?"
C.Patr�n = "212321"
Set C = mvarTabla.Add("A33")
C.Valor = 32
C.C�digo = "@"
C.Patr�n = "232121"
Set C = mvarTabla.Add("A34")
C.Valor = 33
C.C�digo = "A"
C.Patr�n = "111323"
Set C = mvarTabla.Add("A35")
C.Valor = 34
C.C�digo = "B"
C.Patr�n = "131123"
Set C = mvarTabla.Add("A36")
C.Valor = 35
C.C�digo = "C"
C.Patr�n = "131321"
Set C = mvarTabla.Add("A37")
C.Valor = 36
C.C�digo = "D"
C.Patr�n = "112313"
Set C = mvarTabla.Add("A38")
C.Valor = 37
C.C�digo = "E"
C.Patr�n = "132113"
Set C = mvarTabla.Add("A39")
C.Valor = 38
C.C�digo = "F"
C.Patr�n = "132311"
Set C = mvarTabla.Add("A40")
C.Valor = 39
C.C�digo = "G"
C.Patr�n = "211313"
Set C = mvarTabla.Add("A41")
C.Valor = 40
C.C�digo = "H"
C.Patr�n = "231113"
Set C = mvarTabla.Add("A42")
C.Valor = 41
C.C�digo = "I"
C.Patr�n = "231311"
Set C = mvarTabla.Add("A43")
C.Valor = 42
C.C�digo = "J"
C.Patr�n = "112133"
Set C = mvarTabla.Add("A44")
C.Valor = 43
C.C�digo = "K"
C.Patr�n = "112331"
Set C = mvarTabla.Add("A45")
C.Valor = 44
C.C�digo = "L"
C.Patr�n = "132131"
Set C = mvarTabla.Add("A46")
C.Valor = 45
C.C�digo = "M"
C.Patr�n = "113123"
Set C = mvarTabla.Add("A47")
C.Valor = 46
C.C�digo = "N"
C.Patr�n = "113321"
Set C = mvarTabla.Add("A48")
C.Valor = 47
C.C�digo = "O"
C.Patr�n = "133121"
Set C = mvarTabla.Add("A49")
C.Valor = 48
C.C�digo = "P"
C.Patr�n = "313121"
Set C = mvarTabla.Add("A50")
C.Valor = 49
C.C�digo = "Q"
C.Patr�n = "211331"
Set C = mvarTabla.Add("A51")
C.Valor = 50
C.C�digo = "R"
C.Patr�n = "231131"
Set C = mvarTabla.Add("A52")
C.Valor = 51
C.C�digo = "S"
C.Patr�n = "213113"
Set C = mvarTabla.Add("A53")
C.Valor = 52
C.C�digo = "T"
C.Patr�n = "213311"
Set C = mvarTabla.Add("A54")
C.Valor = 53
C.C�digo = "U"
C.Patr�n = "213131"
Set C = mvarTabla.Add("A55")
C.Valor = 54
C.C�digo = "V"
C.Patr�n = "311123"
Set C = mvarTabla.Add("A56")
C.Valor = 55
C.C�digo = "W"
C.Patr�n = "311321"
Set C = mvarTabla.Add("A57")
C.Valor = 56
C.C�digo = "X"
C.Patr�n = "331121"
Set C = mvarTabla.Add("A58")
C.Valor = 57
C.C�digo = "Y"
C.Patr�n = "312113"
Set C = mvarTabla.Add("A59")
C.Valor = 58
C.C�digo = "Z"
C.Patr�n = "312311"
Set C = mvarTabla.Add("A60")
C.Valor = 59
C.C�digo = "["
C.Patr�n = "332111"
Set C = mvarTabla.Add("A61")
C.Valor = 60
C.C�digo = "\"
C.Patr�n = "314111"
Set C = mvarTabla.Add("A62")
C.Valor = 61
C.C�digo = "]"
C.Patr�n = "221411"
Set C = mvarTabla.Add("A63")
C.Valor = 62
C.C�digo = "^"
C.Patr�n = "431111"
Set C = mvarTabla.Add("A64")
C.Valor = 63
C.C�digo = "_"
C.Patr�n = "111224"
Set C = mvarTabla.Add("A65")
C.Valor = 64
C.C�digo = "'"
C.Patr�n = "111422"
Set C = mvarTabla.Add("A66")
C.Valor = 65
C.C�digo = "a"
C.Patr�n = "121124"
Set C = mvarTabla.Add("A67")
C.Valor = 66
C.C�digo = "b"
C.Patr�n = "121421"
Set C = mvarTabla.Add("A68")
C.Valor = 67
C.C�digo = "c"
C.Patr�n = "141122"
Set C = mvarTabla.Add("A69")
C.Valor = 68
C.C�digo = "d"
C.Patr�n = "141221"
Set C = mvarTabla.Add("A70")
C.Valor = 69
C.C�digo = "e"
C.Patr�n = "112214"
Set C = mvarTabla.Add("A71")
C.Valor = 70
C.C�digo = "f"
C.Patr�n = "112412"
Set C = mvarTabla.Add("A72")
C.Valor = 71
C.C�digo = "g"
C.Patr�n = "122114"
Set C = mvarTabla.Add("A73")
C.Valor = 72
C.C�digo = "h"
C.Patr�n = "122411"
Set C = mvarTabla.Add("A74")
C.Valor = 73
C.C�digo = "i"
C.Patr�n = "142112"
Set C = mvarTabla.Add("A75")
C.Valor = 74
C.C�digo = "j"
C.Patr�n = "142211"
Set C = mvarTabla.Add("A76")
C.Valor = 75
C.C�digo = "k"
C.Patr�n = "241211"
Set C = mvarTabla.Add("A77")
C.Valor = 76
C.C�digo = "l"
C.Patr�n = "221114"
Set C = mvarTabla.Add("A78")
C.Valor = 77
C.C�digo = "m"
C.Patr�n = "413111"
Set C = mvarTabla.Add("A79")
C.Valor = 78
C.C�digo = "n"
C.Patr�n = "241112"
Set C = mvarTabla.Add("A80")
C.Valor = 79
C.C�digo = "o"
C.Patr�n = "134111"
Set C = mvarTabla.Add("A81")
C.Valor = 80
C.C�digo = "p"
C.Patr�n = "111242"
Set C = mvarTabla.Add("A82")
C.Valor = 81
C.C�digo = "q"
C.Patr�n = "121142"
Set C = mvarTabla.Add("A83")
C.Valor = 82
C.C�digo = "r"
C.Patr�n = "121241"
Set C = mvarTabla.Add("A84")
C.Valor = 83
C.C�digo = "s"
C.Patr�n = "114212"
Set C = mvarTabla.Add("A85")
C.Valor = 84
C.C�digo = "t"
C.Patr�n = "124112"
Set C = mvarTabla.Add("A86")
C.Valor = 85
C.C�digo = "u"
C.Patr�n = "124211"
Set C = mvarTabla.Add("A87")
C.Valor = 86
C.C�digo = "v"
C.Patr�n = "411212"
Set C = mvarTabla.Add("A88")
C.Valor = 87
C.C�digo = "w"
C.Patr�n = "421112"
Set C = mvarTabla.Add("A89")
C.Valor = 88
C.C�digo = "x"
C.Patr�n = "421211"
Set C = mvarTabla.Add("A90")
C.Valor = 89
C.C�digo = "y"
C.Patr�n = "212141"
Set C = mvarTabla.Add("A91")
C.Valor = 90
C.C�digo = "z"
C.Patr�n = "214121"
Set C = mvarTabla.Add("A92")
C.Valor = 91
C.C�digo = "{"
C.Patr�n = "412121"
Set C = mvarTabla.Add("A93")
C.Valor = 92
C.C�digo = "|"
C.Patr�n = "111143"
Set C = mvarTabla.Add("A94")
C.Valor = 93
C.C�digo = "}"
C.Patr�n = "111341"
Set C = mvarTabla.Add("A95")
C.Valor = 94
C.C�digo = "~"
C.Patr�n = "131141"
Set C = mvarTabla.Add("A96")
C.Valor = 95
C.C�digo = ""
C.Patr�n = "114113"
Set C = mvarTabla.Add("A97")
C.Valor = 96
C.C�digo = ""
C.Patr�n = "114311"
Set C = mvarTabla.Add("A98")
C.Valor = 97
C.C�digo = ""
C.Patr�n = "411113"
Set C = mvarTabla.Add("A99")
C.Valor = 98
C.C�digo = ""
C.Patr�n = "411311"
Set C = mvarTabla.Add("A100")
C.Valor = 99
C.C�digo = ""
C.Patr�n = "113141"
Set C = mvarTabla.Add("A101")
C.Valor = 100
C.C�digo = ""
C.Patr�n = "114131"
Set C = mvarTabla.Add("A102")
C.Valor = 101
C.C�digo = ""
C.Patr�n = "311141"
Set C = mvarTabla.Add("A103")
C.Valor = 102
C.C�digo = ""
C.Patr�n = "411131"
Set C = mvarTabla.Add("A104")
C.Valor = 103
C.C�digo = ""
C.Patr�n = "211412"
Set C = mvarTabla.Add("A105")
C.Valor = 104
C.C�digo = ""
C.Patr�n = "211214"
Set C = mvarTabla.Add("A106")
C.Valor = 105
C.C�digo = ""
C.Patr�n = "211232"
Set C = mvarTabla.Add("A107")
C.Valor = 106
C.C�digo = ""
C.Patr�n = "2331112"
'Set C = mvarTabla.Add("A108")
End Sub

Public Property Set Tabla(ByVal vData As colC128)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci�n Set.
'Syntax: Set x.Tabla = Form1
    Set mvarTabla = vData
End Property

Public Property Get Tabla() As colC128
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci�n.
'Syntax: Debug.Print X.Tabla
    Set Tabla = mvarTabla
End Property

Private Sub Class_Initialize()
mvarFuente = "PF_C128.ttf"
Set mvarTabla = New colC128
InicializarTabla
End Sub


