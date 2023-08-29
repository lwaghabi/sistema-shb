Attribute VB_Name = "CalculaDiaUtil"
Option Explicit

Type DataUtilAnterior
     DiaUtilAnterior As Date
     
End Type

Public Function ObterDataUtilAnterior(DataInformada As Date, NDias As Integer) As DataUtilAnterior


Dim FimPesquisaData As Integer
Dim DiadaSemana As Integer
Dim Limite As Integer

Dim DataObtida As DataUtilAnterior

If NDias = 0 Then
   Limite = 0
Else
   Limite = 1
End If

FimPesquisaData = 0

DataObtida.DiaUtilAnterior = Date - Limite

Do While FimPesquisaData = 0
   
   DiadaSemana = Weekday(DataObtida.DiaUtilAnterior)
   If DiadaSemana > 1 And DiadaSemana < 7 Then
      If Limite = NDias Then
         FimPesquisaData = 1
         ObterDiaUtilAnterior.DiaUtilAnterior = DataObtida.DiaUtilAnterior
      Else
         Limite = Limite + 1
         DataObtida.DiaUtilAnterior = DataObtida.DiaUtilAnterior - 1
      End If
   Else
      DataObtida.DiaUtilAnterior = DataObtida.DiaUtilAnterior - 1
   End If
      
Loop

MsgBox ("Data Final = "), , DataObtida.DiaUtilAnterior
End Function
