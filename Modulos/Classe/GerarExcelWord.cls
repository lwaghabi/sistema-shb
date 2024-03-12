VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GerarExcelWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ExportarContabilidade()
        Dim CaminhoNew As String
                
        CaminhoNew = "C:\Meus Documentos\SISTEMA SHB" & "\docPadrao\"
        
        Dim oApp As Excel.Application
        Dim oWB As Excel.Workbook
        Dim i As Integer
        Dim Ex As Object
        Set Ex = CreateObject("Excel.Application")

        i = 3
         On Error GoTo Erro
            'Create an Excel instance.
50          Set oApp = New Excel.Application

            'Open the desired workbook

60          If Dir(CaminhoNew & "ModelExcelWord.xlsx", vbArchive) = "" Then
70             MsgBox "Não foi possível gerar o documento porque" & vbCrLf & _
               "O arquivo padrão não foi localizado!", vbCritical
80             Exit Sub
90          End If
            
100         Set oWB = oApp.Workbooks.Open(FileName:=CaminhoNew & "ModelExcelWord.xlsx")
            
            'Do any modifications to the workbook.
            Rotina_AbrirBanco
            rs.Open "SELECT * FROM indenizequip", db, 3, 3
            Do Until rs.EOF
               oApp.Cells(i, 1) = rs!descEquip
               oApp.Cells(i, 2) = rs!valor
               rs.MoveNext
               i = i + 1
            Loop
110
          FechaDB

490       oWB.SaveAs FileName:=CaminhoNew & "ExcelWord.xlsx"

510       oWB.Close SaveChanges:=False
520       Set oWB = Nothing
530       oApp.Quit
540       Set oApp = Nothing

400       Ex.Workbooks.Open (CaminhoNew & "ExcelWord.xlsx")
410       Ex.Visible = True

Exit Sub
Erro:
MsgBox "Ocorreu o seguinte erro:" + vbNewLine & _
        Err.Description + vbNewLine & "Na linha: " & Erl
End Sub




