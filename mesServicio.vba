Function MesServicioEspacio(celda As String) As Integer
Dim mesIzquierda As Integer
Dim mesDerecha As Integer
Dim mesValidado As Integer
Const IzqMesNoHallado As Integer = 13
Const DerMesNoHallado As Integer = 13
Const mesDiferente As Integer = 14
Const enero As String = "ene"
Const febrero As String = "feb"
Const marzo As String = "mar"
Const abril As String = "abr"
Const mayo As String = "may"
Const junio As String = "jun"
Const julio As String = "jul"
Const agosto As String = "ago"
Const septiembre As String = "sep"
Const octubre As String = "oct"
Const noviembre As String = "nov"
Const diciembre As String = "dic"

'corresponden al caracter hallado despues de las 3 primeras letras del mes, de derecha a izquierda
Dim caracterEnero As String
Dim caracterFebrero As String
Dim caracterMarzo As String
Dim caracterAbril As String
Dim caracterMayo As String
Dim caracterJunio As String
Dim caracterJulio As String
Dim caracterAgosto As String
Dim caracterSeptiembre As String
Dim caracterOctubre As String
Dim caracterNoviembre As String
Dim caracterDiciembre As String

'corresponden al caracter hallado despues de las 3 primeras letras del mes, de izquierda a derecha
Dim caracterIzqEnero As String
Dim caracterIzqFebrero As String
Dim caracterIzqMarzo As String
Dim caracterIzqAbril As String
Dim caracterIzqMayo As String
Dim caracterIzqJunio As String
Dim caracterIzqJulio As String
Dim caracterIzqAgosto As String
Dim caracterIzqSeptiembre As String
Dim caracterIzqOctubre As String
Dim caracterIzqNoviembre As String
Dim caracterIzqDiciembre As String

celda = LCase(celda)
caracterIzqEnero = Mid(celda, InStr(celda, enero) + 3, 1)
caracterIzqFebrero = Mid(celda, InStr(celda, febrero) + 3, 1)
caracterIzqMarzo = Mid(celda, InStr(celda, marzo) + 3, 1)
caracterIzqAbril = Mid(celda, InStr(celda, abril) + 3, 1)
caracterIzqMayo = Mid(celda, InStr(celda, mayo) + 3, 1)
caracterIzqJunio = Mid(celda, InStr(celda, junio) + 3, 1)
caracterIzqJulio = Mid(celda, InStr(celda, julio) + 3, 1)
caracterIzqAgosto = Mid(celda, InStr(celda, agosto) + 3, 1)
caracterIzqSeptiembre = Mid(celda, InStr(celda, septiembre) + 3, 1)
caracterIzqOctubre = Mid(celda, InStr(celda, octubre) + 3, 1)
caracterIzqNoviembre = Mid(celda, InStr(celda, noviembre) + 3, 1)
caracterIzqDiciembre = Mid(celda, InStr(celda, diciembre) + 3, 1)

caracterEnero = Mid(celda, InStrRev(celda, enero) + 3, 1)
caracterFebrero = Mid(celda, InStrRev(celda, febrero) + 3, 1)
caracterMarzo = Mid(celda, InStrRev(celda, marzo) + 3, 1)
caracterAbril = Mid(celda, InStrRev(celda, abril) + 3, 1)
caracterMayo = Mid(celda, InStrRev(celda, mayo) + 3, 1)
caracterJunio = Mid(celda, InStrRev(celda, junio) + 3, 1)
caracterJulio = Mid(celda, InStrRev(celda, julio) + 3, 1)
caracterAgosto = Mid(celda, InStrRev(celda, agosto) + 3, 1)
caracterSeptiembre = Mid(celda, InStrRev(celda, septiembre) + 3, 1)
caracterOctubre = Mid(celda, InStrRev(celda, octubre) + 3, 1)
caracterNoviembre = Mid(celda, InStrRev(celda, noviembre) + 3, 1)
caracterDiciembre = Mid(celda, InStrRev(celda, diciembre) + 3, 1)

If InStr(celda, enero) > 0 And (Mid(celda, InStr(celda, enero) + 3, 2) = "ro" Or caracterIzqEnero = "/" _
Or caracterIzqEnero = "-" Or caracterIzqEnero = " " Or caracterIzqEnero = "2" _
Or caracterIzqEnero = "1" Or caracterIzqEnero = "") Then
    mesIzquierda = 1
ElseIf InStr(celda, febrero) > 0 And (caracterIzqFebrero = "r" Or caracterIzqFebrero = "/" _
Or caracterIzqFebrero = "-" Or caracterIzqFebrero = " " Or caracterIzqFebrero = "2" _
Or caracterIzqFebrero = "1" Or caracterIzqFebrero = "") Then
    mesIzquierda = 2
ElseIf InStr(celda, marzo) > 0 And (caracterIzqMarzo = "z" Or caracterIzqMarzo = "/" _
Or caracterIzqMarzo = "-" Or caracterIzqMarzo = " " Or caracterIzqMarzo = "2" _
Or caracterIzqMarzo = "1" Or caracterIzqMarzo = "") Then
    mesIzquierda = 3
ElseIf InStr(celda, abril) > 0 And (caracterIzqAbril = "i" Or caracterIzqAbril = "/" _
Or caracterIzqAbril = "-" Or caracterIzqAbril = " " Or caracterIzqAbril = "2" _
Or caracterIzqAbril = "1" Or caracterIzqAbril = "") Then
    mesIzquierda = 4
ElseIf InStr(celda, mayo) > 0 And (caracterIzqMayo = "o" Or caracterIzqMayo = "/" _
Or caracterIzqMayo = "-" Or caracterIzqMayo = " " Or caracterIzqMayo = "2" _
Or caracterIzqMayo = "1" Or caracterIzqMayo = "") Then
    mesIzquierda = 5
ElseIf InStr(celda, junio) > 0 And (caracterIzqJunio = "i" Or caracterIzqJunio = "/" _
Or caracterIzqJunio = "-" Or caracterIzqJunio = " " Or caracterIzqJunio = "2" _
Or caracterIzqJunio = "1" Or caracterIzqJunio = "") Then
    mesIzquierda = 6
ElseIf InStr(celda, julio) > 0 And (caracterIzqJulio = "i" Or caracterIzqJulio = "/" _
Or caracterIzqJulio = "-" Or caracterIzqJulio = " " Or caracterIzqJulio = "2" _
Or caracterIzqJulio = "1" Or caracterIzqJulio = "") Then
    mesIzquierda = 7
ElseIf InStr(celda, agosto) > 0 And (caracterIzqAgosto = "s" Or caracterIzqAgosto = "/" _
Or caracterIzqAgosto = "-" Or caracterIzqAgosto = " " Or caracterIzqAgosto = "2" _
Or caracterIzqAgosto = "1" Or caracterIzqAgosto = "") Then
    mesIzquierda = 8
ElseIf InStr(celda, septiembre) > 0 And (caracterIzqSeptiembre = "t" Or caracterIzqSeptiembre = "/" _
Or caracterIzqSeptiembre = "-" Or caracterIzqSeptiembre = " " Or caracterIzqSeptiembre = "2" _
Or caracterIzqSeptiembre = "1" Or caracterIzqSeptiembre = "") Then
    mesIzquierda = 9
ElseIf InStr(celda, octubre) > 0 And (caracterIzqOctubre = "u" Or caracterIzqOctubre = "/" _
Or caracterIzqOctubre = "-" Or caracterIzqOctubre = " " Or caracterIzqOctubre = "2" _
Or caracterIzqOctubre = "1" Or caracterIzqOctubre = "") Then
    mesIzquierda = 10
ElseIf InStr(celda, noviembre) > 0 And (caracterIzqNoviembre = "i" Or caracterIzqNoviembre = "/" _
Or caracterIzqNoviembre = "-" Or caracterIzqNoviembre = " " Or caracterIzqNoviembre = "2" _
Or caracterIzqNoviembre = "1" Or caracterIzqNoviembre = "") Then
    mesIzquierda = 11
ElseIf InStr(celda, diciembre) > 0 And (caracterIzqDiciembre = "i" Or caracterIzqDiciembre = "/" _
Or caracterIzqDiciembre = "-" Or caracterIzqDiciembre = " " Or caracterIzqDiciembre = "2" _
Or caracterIzqDiciembre = "1" Or caracterIzqDiciembre = "") Then
    mesIzquierda = 12
Else
    mesIzquierda = IzqMesNoHallado
End If


If InStrRev(celda, diciembre) > 0 And (caracterDiciembre = "i" Or caracterDiciembre = "/" _
Or caracterDiciembre = "-" Or caracterDiciembre = " " Or caracterDiciembre = "2" _
Or caracterDiciembre = "1" Or caracterDiciembre = "") Then
    mesDerecha = 12
ElseIf InStrRev(celda, noviembre) > 0 And (caracterNoviembre = "i" Or caracterNoviembre = "/" _
Or caracterNoviembre = "-" Or caracterNoviembre = " " Or caracterNoviembre = "2" _
Or caracterNoviembre = "1" Or caracterNoviembre = "") Then
    mesDerecha = 11
ElseIf InStrRev(celda, octubre) > 0 And (caracterOctubre = "u" Or caracterOctubre = "/" _
Or caracterOctubre = "-" Or caracterOctubre = " " Or caracterOctubre = "2" _
Or caracterOctubre = "1" Or caracterOctubre = "") Then
    mesDerecha = 10
ElseIf InStrRev(celda, septiembre) > 0 And (caracterSeptiembre = "t" Or caracterSeptiembre = "/" _
Or caracterSeptiembre = "-" Or caracterSeptiembre = " " Or caracterSeptiembre = "2" _
Or caracterSeptiembre = "1" Or caracterSeptiembre = "") Then
    mesDerecha = 9
ElseIf InStrRev(celda, agosto) > 0 And (caracterAgosto = "s" Or caracterAgosto = "/" _
Or caracterAgosto = "-" Or caracterAgosto = " " Or caracterAgosto = "2" _
Or caracterAgosto = "1" Or caracterAgosto = "") Then
    mesDerecha = 8
ElseIf InStrRev(celda, julio) > 0 And (caracterJulio = "i" Or caracterJulio = "/" _
Or caracterJulio = "-" Or caracterJulio = " " Or caracterJulio = "2" _
Or caracterJulio = "1" Or caracterJulio = "") Then
    mesDerecha = 7
ElseIf InStrRev(celda, junio) > 0 And (caracterJunio = "i" Or caracterJunio = "/" _
Or caracterJunio = "-" Or caracterJunio = " " Or caracterJunio = "2" _
Or caracterJunio = "1" Or caracterJunio = "") Then
    mesDerecha = 6
ElseIf InStrRev(celda, mayo) > 0 And (caracterMayo = "o" Or caracterMayo = "/" _
Or caracterMayo = "-" Or caracterMayo = " " Or caracterMayo = "2" _
Or caracterMayo = "1" Or caracterMayo = "") Then
    mesDerecha = 5
ElseIf InStrRev(celda, abril) > 0 And (caracterAbril = "i" Or caracterAbril = "/" _
Or caracterAbril = "-" Or caracterAbril = " " Or caracterAbril = "2" _
Or caracterAbril = "1" Or caracterAbril = "") Then
    mesDerecha = 4
ElseIf InStrRev(celda, marzo) > 0 And (caracterMarzo = "z" Or caracterMarzo = "/" _
Or caracterMarzo = "-" Or caracterMarzo = " " Or caracterMarzo = "2" _
Or caracterMarzo = "1" Or caracterMarzo = "") Then
    mesDerecha = 3
ElseIf InStrRev(celda, febrero) > 0 And (caracterFebrero = "r" Or caracterFebrero = "/" _
Or caracterFebrero = "-" Or caracterFebrero = " " Or caracterFebrero = "2" _
Or caracterFebrero = "1" Or caracterFebrero = "") Then
    mesDerecha = 2
ElseIf InStrRev(celda, enero) > 0 And (Mid(celda, InStrRev(celda, enero) + 3, 2) = "ro" Or caracterEnero = "/" _
Or caracterEnero = "-" Or caracterEnero = " " Or caracterEnero = "2" _
Or caracterEnero = "1" Or caracterEnero = "") Then
    mesDerecha = 1
Else
    mesDerecha = DerMesNoHallado
End If

If (mesIzquierda = mesDerecha) Then
    MesServicioEspacio = mesIzquierda
ElseIf (mesIzquierda = 13 And mesDerecha <> 13) Then
    MesServicioEspacio = mesDerecha
ElseIf (mesDerecha = 13 And mesIzquierda <> 13) Then
    MesServicioEspacio = mesIzquierda
Else
    MesServicioEspacio = mesDiferente
End If

End Function


