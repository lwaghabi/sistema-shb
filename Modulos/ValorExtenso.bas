Attribute VB_Name = "ValorExtenso"
Option Explicit

Type ValorExtenso
     ValorPorExtenso As String
         
End Type

    
Function ValorExtenso(AcumulaValor) As AcumulaValor

Dim nValor As Currency
nValor = AcumulaValor
'Valida Argumento
If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then
Exit Function
End If

'Variáveis
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), aTexto(4) As String

'Matrizes de extensos (Parciais)
ReDim aUnid(19) As String
aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
aUnid(19) = "dezenove "

ReDim aDezena(9) As String
aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
aDezena(9) = "noventa "

ReDim aCentena(9) As String
aCentena(1) = "cento ": aCentena(2) = "duzentos "
aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
aCentena(9) = "novecentos "

'Separa valor em grupos
cValor = Format$(nValor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)

'Calcula cada grupo
For nContador = 1 To 4
  cParte = aGrupo(nContador)
  nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
  If nTamanho = 3 Then
    If Right$(cParte, 2) <> "00" Then
      aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
      nTamanho = 2
    Else
      aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", aCentena(Left(cParte, 1)))
    End If
  End If
  If nTamanho = 2 Then
    If Val(Right(cParte, 2)) < 20 Then
      aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
    Else
      aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
      If Right$(cParte, 1) <> "0" Then
        aTexto(nContador) = aTexto(nContador) + "e "
        nTamanho = 1
      End If
    End If
  End If
  If nTamanho = 1 Then
    aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
  End If
Next

'Final
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
  cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos")
Else
  cFinal = ""
  cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "milhões ", "milhão "), "")
  If Val(aGrupo(2) + aGrupo(3)) = 0 Then
    cFinal = cFinal + "de "
  Else
    cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "mil ", "")
  End If
  cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "real ", "reais ")
  cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos"), "")
End If

ValorPorExtenso = UCase$(cFinal)

End Function


