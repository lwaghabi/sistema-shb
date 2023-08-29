Attribute VB_Name = "DiaUtilAnterior"
Option Explicit

Type DataDiaUtil
     DiaUtilAnterior As Date
End Type

Function ObterDiaUtilAnterior() As DataDiaUtil

Dim DataObtida As DataDiaUtil
Dim FimPesquisaData As Integer
Dim DiadaSemana As Integer
Dim DataInformada As Date

FimPesquisaData = 0

DataInformada = Date - 1

Do While FimPesquisaData = 0
   DiadaSemana = Weekday(DataInformada)
   If DiadaSemana > 1 And DiadaSemana < 7 Then
      FimPesquisaData = 1
      DataObtida.DiaUtilAnterior = DataInformada
      ObterDiaUtilAnterior = DataObtida
   Else
      DataInformada = DataInformada - 1
   End If
Loop
End Function
