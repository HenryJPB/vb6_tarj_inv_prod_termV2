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
'// Diseñado por Sergio Fustero Carreras (2006)
'//
'// Crea e imprime código de barras 128 (tipos A y B)
'//     El tipo C está pendiente de implementar
'// Este código de barras se usa, por ejemplo, en las
'//     cartas certificadas, listas de correos, acuses de
'//     recibo, etc. de Correos (España)
'//
'// Puede imprimir directamente a impresora o a un
'//     contexto de dispositivo (p.e. PictureBox)
'//
'////////////////////////////////////////////////////////////////////

'variables locales para almacenar los valores de las propiedades
Private mvarTabla As colC128 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarFuente As String 'copia local

Public Enum eTipoDeCódigo128
    cC128_A = 1
    cC128_B = 2
    cC128_C = 3
End Enum

Public Function GenerarBarras(Texto As String, DC As Object, Optional Codificación As eTipoDeCódigo128 = cC128_B, Optional X As Single = -1, Optional Y As Single = -1, Optional Densidad As Single = 15, Optional Alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tamaño As Integer = 4) As Boolean
'------------------------------
' Genera las barras (secuencias de líneas negras y espacios en blanco)
'       en código 128 del tipo A, B ó C para un contexto dado.
' El contexto puede ser un PictureBox, p.e.
'
' Parámetros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- DC:
'       Contexto donde se imprimirá. Puede ser util para los que uséis vista previa.
'       Para imprimir directamente a la impresora, usa la función GenerarBarrasImpresora
'           (puede generar barras mucho más pequeñas)
'   -- Codificación:
'       Indicar cC128_A, cC128_B ó cC128_C
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
'   -- Tamaño:
'       Tamaño de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aquí As Single, CBarras As String, C As Integer, D As Integer, Paso, Cx, Ancho
'DC.CurrentX = Densidad
' -- Establece el ancho de las líneas
On Error Resume Next
DC.DrawWidth = 1    'Algunos contextos dan error
On Error GoTo 0
Ancho = Densidad
' -- Establecer posición vertical (fija durante todo el proceso)
If Y = -1 Then
    Y = DC.CurrentY
End If
' -- Establecer posición horizontal inicial
If X = -1 Then
    Aquí = DC.CurrentX
Else
    Aquí = X
End If
' -- Espacio en blanco antes del código de barras
Aquí = Aquí + 10 * Densidad
' -- Conseguir el código de barras
CBarras = Barras(Texto, Codificación)
' -- Imprimir el código de barras (se intercalan líneas y espacios)
For C = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    For D = 1 To Val(Mid(CBarras, C, 1))
        If (C Mod 2) = 0 Then ' 0 = Falso --> Línea
                                     ' 1 = Verdadero --> Espacio
            DC.Line (Aquí, Y)-(Aquí, Y + Alto), vbWhite
        Else
            DC.Line (Aquí, Y)-(Aquí, Y + Alto), vbBlack
        End If
        Aquí = Aquí + Ancho
    Next
Next
For C = 1 To 10 'Espacio en blanco después
    DC.Line (Aquí, Y)-(Aquí, Y + Alto), vbWhite
    Aquí = Aquí + Ancho
Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    DC.Font.Name = Fuente
    DC.Font.Size = Tamaño
    DC.CurrentY = Y + Alto
    DC.CurrentX = Cx + (10 * Densidad) + (Aquí - Cx - DC.TextWidth(Texto)) / 2
    DC.Print Texto
End If
End Function

Public Function GenerarBarrasImpresora(Texto As String, Optional Codificación As eTipoDeCódigo128 = cC128_B, Optional X As Single = -1, Optional Y As Single = -1, Optional Densidad As Single = 10, Optional Alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tamaño As Integer = 4) As Boolean
'------------------------------
' Genera las barras (secuencias de líneas negras y espacios en blanco)
'       en código 128 del tipo A, B ó C directamente en la impresora.
'
' Parámetros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- Codificación:
'       Indicar cC128_A, cC128_B ó cC128_C
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
'   -- Tamaño:
'       Tamaño de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aquí As Single, CBarras As String, C As Integer, D As Integer, Paso, Cx, Ancho
' -- Establece el ancho de las líneas
Printer.DrawWidth = 5
Ancho = Densidad
' -- Establecer posición vertical (fija durante todo el proceso)
If Y = -1 Then
    Y = Printer.CurrentY
End If
' -- Establecer posición horizontal inicial
If X = -1 Then
    Aquí = Printer.CurrentX
Else
    Aquí = X
End If
' -- Espacio en blanco antes del código de barras
Aquí = Aquí + 10 * Densidad
' -- Conseguir el código de barras
CBarras = Barras(Texto, Codificación)
' -- Imprimir el código de barras (se intercalan líneas y espacios)
For C = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    For D = 1 To Val(Mid(CBarras, C, 1))
        If (C Mod 2) = 0 Then ' 0 = Falso --> Línea
                                     ' 1 = Verdadero --> Espacio
            Printer.Line (Aquí, Y)-(Aquí, Y + Alto), vbWhite
        Else
            Printer.Line (Aquí, Y)-(Aquí, Y + Alto), vbBlack
        End If
        Aquí = Aquí + Ancho
    Next
Next
' -- Si es necesario, pondremos espacio en blanco después
'For C = 1 To 10 'Espacio en blanco después
'    Printer.Line (Aquí, Y)-(Aquí, Y + Alto), vbWhite
'    Aquí = Aquí + Ancho
'Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    Printer.Font.Name = Fuente
    Printer.Font.Size = Tamaño
    Printer.CurrentY = Y + Alto
    Printer.CurrentX = Cx + (10 * Densidad) + (Aquí - Cx - Printer.TextWidth(Texto)) / 2
    Printer.Print Texto
End If
End Function


Public Function Barras(Texto As String, Optional Codificación As eTipoDeCódigo128 = cC128_B) As String
'------------------------------
' Crea la secuencia de barras y espacios para el texto dado
'
' Falta de implementar el código para el tipo C
'       (solamente admite números y es mucho más reducido)
'------------------------------
Dim C As Integer, Letra As Integer, Suma As Long
' -- Valor inicial
Select Case Codificación
    Case eTipoDeCódigo128.cC128_A
        Barras = mvarTabla(103 + 1).Patrón
        Suma = 103
    Case eTipoDeCódigo128.cC128_B
        Barras = mvarTabla(104 + 1).Patrón
        Suma = 104
    Case eTipoDeCódigo128.cC128_C
        Barras = mvarTabla(105 + 1).Patrón
        Suma = 105
End Select
' -- Secuencia del texto
If Codificación = cC128_C Then
Else
    For C = 1 To Len(Texto)
        For Letra = 1 To 63
            If mvarTabla(Letra).Código = Mid(Texto, C, 1) Then
                Barras = Barras & mvarTabla(Letra).Patrón
                Suma = Suma + mvarTabla(Letra).Valor * C
                Exit For
            End If
        Next
        If Codificación = cC128_A Then
            For Letra = 0 To 31
                If Letra = Asc(Mid(Texto, C, 1)) Then
                    Barras = Barras & mvarTabla(Letra).Patrón
                    Suma = Suma + mvarTabla(Letra).Valor * C
                    Exit For
                End If
            Next
        Else
            For Letra = 64 To 94
                If mvarTabla(Letra).Código = Mid(Texto, C, 1) Then
                    Barras = Barras & mvarTabla(Letra).Patrón
                    Suma = Suma + mvarTabla(Letra).Valor * C
                    Exit For
                End If
            Next
        End If
        For Letra = 95 To 102
            If Asc(Mid(Texto, C, 1)) = Letra + 32 Then
                Barras = Barras & mvarTabla(Letra).Patrón
                Suma = Suma + mvarTabla(Letra).Valor * C
                Exit For
            End If
        Next
    Next
End If
' -- Código de control
Barras = Barras & mvarTabla((Suma Mod 103) + 1).Patrón
' -- Final
Barras = Barras & mvarTabla(106 + 1).Patrón
End Function


Public Function PonerFuente(Objeto As Object) As Boolean
'---------------------------
' Asigna la fuente (que se debe haber indicado con anterioridad
'       al objeto pasado como parámetro, si la tiene
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
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Fuente
    Fuente = mvarFuente
End Property

Public Function Texto(Frase As String, Optional Clase As eTipoDeCódigo128 = cC128_A) As String
'*******************************************************
'Devuelve el texto correspondiente para usarlo
'   con la fuente de código de barras
'
'*******************************************************
Dim Suma As Long, C As Integer, V As Integer, Resto As Integer
' -- Valor de inicio, según tipo
Select Case Clase
    Case cC128_A
        Suma = 203
    Case cC128_B
        Suma = 204
    Case cC128_C
        Suma = 205
End Select
Texto = Chr(Suma)
' -- Cálculo de la suma a codificar
If Clase = cC128_C Then
    For C = 1 To Len(Frase) Step 2
        V = Val(Mid(Frase, C, 2))
        Texto = Texto & mvarTabla(V).Código
        Suma = Suma + V * C
    Next
Else
    For C = 1 To Len(Frase)
        Select Case Clase
            Case eTipoDeCódigo128.cC128_A
                V = Valor(Mid(Frase, C, 1))
                If V < 64 Then
                    Texto = Texto & Mid(Frase, C, 1)
                End If
            Case eTipoDeCódigo128.cC128_B
                V = Valor(Mid(Frase, C, 1))
                If V < 95 Then
                    Texto = Texto & Mid(Frase, C, 1)
                End If
            Case eTipoDeCódigo128.cC128_C
        End Select
        Suma = Suma + V * C
    Next
End If
' -- Dígito de control
Resto = Suma Mod 103
Texto = Texto + Chr(Resto)
' -- Código final
Texto = Texto & Chr(106)
End Function

Private Function Patrón(Texto As String) As String
'---------------------------
' Devuelve el patrón para un carácter dado
'---------------------------
Dim V As C128
Patrón = ""
For Each V In mvarTabla
    If V.Código = Texto Then
        Patrón = V.Patrón
        Exit For
    End If
Next
End Function

Private Function Valor(Texto As String) As Integer
'---------------------------
' Devuelve el valor numérico correspondiente
'       a un carácter dado
'---------------------------
Dim V As C128
Valor = -1
For Each V In mvarTabla
    If V.Código = Texto Then
        Valor = V.Valor
        Exit For
    End If
Next
End Function

Public Sub InicializarTabla()
'---------------------------
' Crea la tabla con los patrones para cada símbolo
'
' El patrón indica, alternativamente, líneas y espacios
'---------------------------
Dim C As C128
Set C = mvarTabla.Add("A1")
C.Valor = 0
C.Código = " "
C.Patrón = "212222"
Set C = mvarTabla.Add("A2")
C.Valor = 1
C.Código = "!"
C.Patrón = "222122"
Set C = mvarTabla.Add("A3")
C.Valor = 2
C.Código = Chr(34)  ' Comillas dobles
C.Patrón = "222221"
Set C = mvarTabla.Add("A4")
C.Valor = 3
C.Código = "#"
C.Patrón = "121223"
Set C = mvarTabla.Add("A5")
C.Valor = 4
C.Código = "$"
C.Patrón = "121322"
Set C = mvarTabla.Add("A6")
C.Valor = 5
C.Código = "%"
C.Patrón = "131222"
Set C = mvarTabla.Add("A7")
C.Valor = 6
C.Código = "&"
C.Patrón = "122213"
Set C = mvarTabla.Add("A8")
C.Valor = 7
C.Código = "'"
C.Patrón = "122312"
Set C = mvarTabla.Add("A9")
C.Valor = 8
C.Código = "("
C.Patrón = "132212"
Set C = mvarTabla.Add("A10")
C.Valor = 9
C.Código = ")"
C.Patrón = "221213"
Set C = mvarTabla.Add("A11")
'Set C = New C128
C.Valor = 10
C.Código = "*"
C.Patrón = "221312"
Set C = mvarTabla.Add("A12")
C.Valor = 11
C.Código = "+"
C.Patrón = "231212"
Set C = mvarTabla.Add("A13")
C.Valor = 12
C.Código = ","
C.Patrón = "112232"
Set C = mvarTabla.Add("A14")
C.Valor = 13
C.Código = "-"
C.Patrón = "122132"
Set C = mvarTabla.Add("A15")
C.Valor = 14
C.Código = "."
C.Patrón = "122231"
Set C = mvarTabla.Add("A16")
C.Valor = 15
C.Código = "/"
C.Patrón = "113222"
Set C = mvarTabla.Add("A17")
C.Valor = 16
C.Código = "0"
C.Patrón = "123122"
Set C = mvarTabla.Add("A18")
C.Valor = 17
C.Código = "1"
C.Patrón = "123221"
Set C = mvarTabla.Add("A19")
C.Valor = 18
C.Código = "2"
C.Patrón = "223211"
Set C = mvarTabla.Add("A20")
C.Valor = 19
C.Código = "3"
C.Patrón = "221132"
Set C = mvarTabla.Add("A21")
'Set C = New C128
C.Valor = 20
C.Código = "4"
C.Patrón = "221231"
Set C = mvarTabla.Add("A22")
C.Valor = 21
C.Código = "5"
C.Patrón = "213212"
Set C = mvarTabla.Add("A23")
C.Valor = 22
C.Código = "6"
C.Patrón = "223112"
Set C = mvarTabla.Add("A24")
C.Valor = 23
C.Código = "7"
C.Patrón = "312131"
Set C = mvarTabla.Add("A25")
C.Valor = 24
C.Código = "8"
C.Patrón = "311222"
Set C = mvarTabla.Add("A26")
C.Valor = 25
C.Código = "9"
C.Patrón = "321122"
Set C = mvarTabla.Add("A27")
C.Valor = 26
C.Código = ":"
C.Patrón = "321221"
Set C = mvarTabla.Add("A28")
C.Valor = 27
C.Código = ";"
C.Patrón = "312212"
Set C = mvarTabla.Add("A29")
C.Valor = 28
C.Código = "<"
C.Patrón = "322112"
Set C = mvarTabla.Add("A30")
C.Valor = 29
C.Código = "="
C.Patrón = "322211"
Set C = mvarTabla.Add("A31")
'Set C = New C128
C.Valor = 30
C.Código = ">"
C.Patrón = "212123"
Set C = mvarTabla.Add("A32")
C.Valor = 31
C.Código = "?"
C.Patrón = "212321"
Set C = mvarTabla.Add("A33")
C.Valor = 32
C.Código = "@"
C.Patrón = "232121"
Set C = mvarTabla.Add("A34")
C.Valor = 33
C.Código = "A"
C.Patrón = "111323"
Set C = mvarTabla.Add("A35")
C.Valor = 34
C.Código = "B"
C.Patrón = "131123"
Set C = mvarTabla.Add("A36")
C.Valor = 35
C.Código = "C"
C.Patrón = "131321"
Set C = mvarTabla.Add("A37")
C.Valor = 36
C.Código = "D"
C.Patrón = "112313"
Set C = mvarTabla.Add("A38")
C.Valor = 37
C.Código = "E"
C.Patrón = "132113"
Set C = mvarTabla.Add("A39")
C.Valor = 38
C.Código = "F"
C.Patrón = "132311"
Set C = mvarTabla.Add("A40")
C.Valor = 39
C.Código = "G"
C.Patrón = "211313"
Set C = mvarTabla.Add("A41")
C.Valor = 40
C.Código = "H"
C.Patrón = "231113"
Set C = mvarTabla.Add("A42")
C.Valor = 41
C.Código = "I"
C.Patrón = "231311"
Set C = mvarTabla.Add("A43")
C.Valor = 42
C.Código = "J"
C.Patrón = "112133"
Set C = mvarTabla.Add("A44")
C.Valor = 43
C.Código = "K"
C.Patrón = "112331"
Set C = mvarTabla.Add("A45")
C.Valor = 44
C.Código = "L"
C.Patrón = "132131"
Set C = mvarTabla.Add("A46")
C.Valor = 45
C.Código = "M"
C.Patrón = "113123"
Set C = mvarTabla.Add("A47")
C.Valor = 46
C.Código = "N"
C.Patrón = "113321"
Set C = mvarTabla.Add("A48")
C.Valor = 47
C.Código = "O"
C.Patrón = "133121"
Set C = mvarTabla.Add("A49")
C.Valor = 48
C.Código = "P"
C.Patrón = "313121"
Set C = mvarTabla.Add("A50")
C.Valor = 49
C.Código = "Q"
C.Patrón = "211331"
Set C = mvarTabla.Add("A51")
C.Valor = 50
C.Código = "R"
C.Patrón = "231131"
Set C = mvarTabla.Add("A52")
C.Valor = 51
C.Código = "S"
C.Patrón = "213113"
Set C = mvarTabla.Add("A53")
C.Valor = 52
C.Código = "T"
C.Patrón = "213311"
Set C = mvarTabla.Add("A54")
C.Valor = 53
C.Código = "U"
C.Patrón = "213131"
Set C = mvarTabla.Add("A55")
C.Valor = 54
C.Código = "V"
C.Patrón = "311123"
Set C = mvarTabla.Add("A56")
C.Valor = 55
C.Código = "W"
C.Patrón = "311321"
Set C = mvarTabla.Add("A57")
C.Valor = 56
C.Código = "X"
C.Patrón = "331121"
Set C = mvarTabla.Add("A58")
C.Valor = 57
C.Código = "Y"
C.Patrón = "312113"
Set C = mvarTabla.Add("A59")
C.Valor = 58
C.Código = "Z"
C.Patrón = "312311"
Set C = mvarTabla.Add("A60")
C.Valor = 59
C.Código = "["
C.Patrón = "332111"
Set C = mvarTabla.Add("A61")
C.Valor = 60
C.Código = "\"
C.Patrón = "314111"
Set C = mvarTabla.Add("A62")
C.Valor = 61
C.Código = "]"
C.Patrón = "221411"
Set C = mvarTabla.Add("A63")
C.Valor = 62
C.Código = "^"
C.Patrón = "431111"
Set C = mvarTabla.Add("A64")
C.Valor = 63
C.Código = "_"
C.Patrón = "111224"
Set C = mvarTabla.Add("A65")
C.Valor = 64
C.Código = "'"
C.Patrón = "111422"
Set C = mvarTabla.Add("A66")
C.Valor = 65
C.Código = "a"
C.Patrón = "121124"
Set C = mvarTabla.Add("A67")
C.Valor = 66
C.Código = "b"
C.Patrón = "121421"
Set C = mvarTabla.Add("A68")
C.Valor = 67
C.Código = "c"
C.Patrón = "141122"
Set C = mvarTabla.Add("A69")
C.Valor = 68
C.Código = "d"
C.Patrón = "141221"
Set C = mvarTabla.Add("A70")
C.Valor = 69
C.Código = "e"
C.Patrón = "112214"
Set C = mvarTabla.Add("A71")
C.Valor = 70
C.Código = "f"
C.Patrón = "112412"
Set C = mvarTabla.Add("A72")
C.Valor = 71
C.Código = "g"
C.Patrón = "122114"
Set C = mvarTabla.Add("A73")
C.Valor = 72
C.Código = "h"
C.Patrón = "122411"
Set C = mvarTabla.Add("A74")
C.Valor = 73
C.Código = "i"
C.Patrón = "142112"
Set C = mvarTabla.Add("A75")
C.Valor = 74
C.Código = "j"
C.Patrón = "142211"
Set C = mvarTabla.Add("A76")
C.Valor = 75
C.Código = "k"
C.Patrón = "241211"
Set C = mvarTabla.Add("A77")
C.Valor = 76
C.Código = "l"
C.Patrón = "221114"
Set C = mvarTabla.Add("A78")
C.Valor = 77
C.Código = "m"
C.Patrón = "413111"
Set C = mvarTabla.Add("A79")
C.Valor = 78
C.Código = "n"
C.Patrón = "241112"
Set C = mvarTabla.Add("A80")
C.Valor = 79
C.Código = "o"
C.Patrón = "134111"
Set C = mvarTabla.Add("A81")
C.Valor = 80
C.Código = "p"
C.Patrón = "111242"
Set C = mvarTabla.Add("A82")
C.Valor = 81
C.Código = "q"
C.Patrón = "121142"
Set C = mvarTabla.Add("A83")
C.Valor = 82
C.Código = "r"
C.Patrón = "121241"
Set C = mvarTabla.Add("A84")
C.Valor = 83
C.Código = "s"
C.Patrón = "114212"
Set C = mvarTabla.Add("A85")
C.Valor = 84
C.Código = "t"
C.Patrón = "124112"
Set C = mvarTabla.Add("A86")
C.Valor = 85
C.Código = "u"
C.Patrón = "124211"
Set C = mvarTabla.Add("A87")
C.Valor = 86
C.Código = "v"
C.Patrón = "411212"
Set C = mvarTabla.Add("A88")
C.Valor = 87
C.Código = "w"
C.Patrón = "421112"
Set C = mvarTabla.Add("A89")
C.Valor = 88
C.Código = "x"
C.Patrón = "421211"
Set C = mvarTabla.Add("A90")
C.Valor = 89
C.Código = "y"
C.Patrón = "212141"
Set C = mvarTabla.Add("A91")
C.Valor = 90
C.Código = "z"
C.Patrón = "214121"
Set C = mvarTabla.Add("A92")
C.Valor = 91
C.Código = "{"
C.Patrón = "412121"
Set C = mvarTabla.Add("A93")
C.Valor = 92
C.Código = "|"
C.Patrón = "111143"
Set C = mvarTabla.Add("A94")
C.Valor = 93
C.Código = "}"
C.Patrón = "111341"
Set C = mvarTabla.Add("A95")
C.Valor = 94
C.Código = "~"
C.Patrón = "131141"
Set C = mvarTabla.Add("A96")
C.Valor = 95
C.Código = ""
C.Patrón = "114113"
Set C = mvarTabla.Add("A97")
C.Valor = 96
C.Código = ""
C.Patrón = "114311"
Set C = mvarTabla.Add("A98")
C.Valor = 97
C.Código = ""
C.Patrón = "411113"
Set C = mvarTabla.Add("A99")
C.Valor = 98
C.Código = ""
C.Patrón = "411311"
Set C = mvarTabla.Add("A100")
C.Valor = 99
C.Código = ""
C.Patrón = "113141"
Set C = mvarTabla.Add("A101")
C.Valor = 100
C.Código = ""
C.Patrón = "114131"
Set C = mvarTabla.Add("A102")
C.Valor = 101
C.Código = ""
C.Patrón = "311141"
Set C = mvarTabla.Add("A103")
C.Valor = 102
C.Código = ""
C.Patrón = "411131"
Set C = mvarTabla.Add("A104")
C.Valor = 103
C.Código = ""
C.Patrón = "211412"
Set C = mvarTabla.Add("A105")
C.Valor = 104
C.Código = ""
C.Patrón = "211214"
Set C = mvarTabla.Add("A106")
C.Valor = 105
C.Código = ""
C.Patrón = "211232"
Set C = mvarTabla.Add("A107")
C.Valor = 106
C.Código = ""
C.Patrón = "2331112"
'Set C = mvarTabla.Add("A108")
End Sub

Public Property Set Tabla(ByVal vData As colC128)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Tabla = Form1
    Set mvarTabla = vData
End Property

Public Property Get Tabla() As colC128
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tabla
    Set Tabla = mvarTabla
End Property

Private Sub Class_Initialize()
mvarFuente = "PF_C128.ttf"
Set mvarTabla = New colC128
InicializarTabla
End Sub


