VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmGeraCustoExcel 
   Caption         =   "frmGeraCustoExcel"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   4560
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ComboBox cmbMesFim 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4920
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cmbAnoFim 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6240
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox cmbAnoInicio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox cmbMesInicio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdGerarExcel 
      BackColor       =   &H00FFFF00&
      Caption         =   "Gerar Custo"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox dtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Período Até:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Período De:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Gera Custo Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmGeraCustoExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j As Integer
Dim ContaReg As Integer

Private Sub cmdGerarExcel_Click()
   Screen.MousePointer = vbHourglass
   Call GerarExcel
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
   dtHoje = Date
   j = 1
   Do While j <= 12
      cmbMesInicio.AddItem j
      cmbMesFim.AddItem j
      j = j + 1
   Loop

   j = 2020

   Do While j <= Year(Date)
      cmbAnoInicio.AddItem j
      cmbAnoFim.AddItem j
      j = j + 1
   Loop
End Sub

Public Sub GerarExcel()

        Dim CaminhoNew As String
        Dim query As String
        Dim TotRegCtaRec As Integer
        
        Call Rotina_AbrirBanco

'        rs.Open "SELECT COUNT(*) as total FROM HistoricoNotaFiscalDetProd", db, 3, 3
'        TotRegCtaRec = rs!Total
'
'        Call Rotina_AbrirBanco
7

'        usu.Open "select usuEnderecoOneDrive from Usuario where  chnome = ('" & glbUsuario & "')", db, 3, 3
'
'        If usu!usuEnderecoOneDrive = Null Then
'         MsgBox ("Não autorizada a impressão de proposta")
'         FechaDB
'         Exit Sub
'        Else
'         CaminhoNew = usu!usuEnderecoOneDrive & "Sistema\PROPOSTA MODELO\"
'        End If

         CaminhoNew = "C:\Meus Documentos\SISTEMA SHB\"
        
        Dim oApp As Excel.Application
        Dim oWB As Excel.Workbook
        Dim i As Integer
        Dim Ex As Object
        Dim dif As Integer
        Dim mesAtual As Integer
        Dim anoAtual As Integer
        Dim dataInicio As String
        Dim dataFim As String
        Dim excelCelula As Excel.Worksheet
        Dim excelValores As Excel.Worksheet
        Dim excelRecebidos As Excel.Worksheet
        Dim excelPrevistos As Excel.Worksheet
        Set Ex = CreateObject("Excel.Application")

        i = 2
         
            'Create an Excel instance.
50          Set oApp = New Excel.Application

            'Open the desired workbook

60          If Dir(CaminhoNew & "CustoModelo.xlsx", vbArchive) = "" Then
70             MsgBox "Não foi possível gerar o documento porque" & vbCrLf & _
               "O arquivo padrão não foi localizado!", vbCritical
80             Exit Sub
90          End If
            
            dataInicio = cmbAnoInicio & "-" & Format$(cmbMesInicio, "00") & "-" & "01"
            dataFim = cmbAnoFim & "-" & Format$(cmbMesFim, "00") & "-" & "31"
            
100         Set oWB = oApp.Workbooks.Open(FileName:=CaminhoNew & "CustoModelo.xlsx")
            Set excelCelula = oWB.Worksheets("Periodos")
            Set excelValores = oWB.Worksheets("Valores")
            Set excelRecebidos = oWB.Worksheets("Receitas Recebidas")
            Set excelPrevistos = oWB.Worksheets(" Receitas Previstas")
            
            Call Rotina_AbrirBanco
            'Do any modifications to the workbook.
            rs.Open "SELECT HistoricoNotaFiscalDetProd.chPessoa," & _
             " HistoricoNotaFiscalDetProd.chNotaFiscalEntrada," & _
             " HistoricoContasPagar.chFatura,chDataVencito,ctpDataPagamento," & _
             " nfdValorParcela,nfdCentroDeCusto,nfdGrupoCentroDeCusto," & _
             " nfdSubGrupoCentroDeCusto FROM HistoricoNotaFiscalDetProd INNER JOIN" & _
             " HistoricoContasPagar ON HistoricoContasPagar.chPessoa =" & _
             " HistoricoNotaFiscalDetProd.chPessoa AND" & _
             " HistoricoNotaFiscalDetProd.chNotaFiscalEntrada=" & _
             " HistoricoContasPagar.chNotaFiscal WHERE nfdCentroDeCusto = 2 AND" & _
             " nfdGrupoCentroDeCusto > 00 AND HistoricoContasPagar.ctpStatus=1" & _
             " and ctpDataPagamento>= ('" & dataInicio & "') and ctpDataPagamento <= ('" & dataFim & "')" & _
             " UNION ALL SELECT NotaFiscalDetProd.chPessoa," & _
             " NotaFiscalDetProd.chNotaFiscalEntrada," & _
             " Contas_A_Pagar.chFatura,chDataVencito,ctpDataPagamento," & _
             " nfdValorParcela,nfdCentroDeCusto,nfdGrupoCentroDeCusto," & _
             " nfdSubGrupoCentroDeCusto FROM NotaFiscalDetProd INNER JOIN" & _
             " Contas_A_Pagar ON Contas_A_Pagar.chPessoa =" & _
             " NotaFiscalDetProd.chPessoa AND" & _
             " NotaFiscalDetProd.chNotaFiscalEntrada=" & _
             " Contas_A_Pagar.chNotaFiscal WHERE nfdCentroDeCusto = 2 AND" & _
             " nfdGrupoCentroDeCusto > 00 AND Contas_A_Pagar.ctpStatus=1" & _
             " and ctpDataPagamento>= ('" & dataInicio & "') and ctpDataPagamento <= ('" & dataFim & "')" & _
             " order by ctpDataPagamento", db, 3, 3
            
            ContaReg = 0
            
            Do Until rs.EOF
               excelValores.Cells(i, 1) = rs!chPessoa
               excelValores.Cells(i, 2) = rs!chNotaFiscalEntrada
               excelValores.Cells(i, 3) = rs!chFatura
               excelValores.Cells(i, 4) = rs!chDataVencito
               excelValores.Cells(i, 5) = rs!ctpDataPagamento
               excelValores.Cells(i, 6) = rs!nfdValorParcela
               excelValores.Cells(i, 7) = rs!nfdCentroDeCusto
               excelValores.Cells(i, 8) = rs!nfdGrupoCentroDeCusto
               excelValores.Cells(i, 9) = rs!nfdSubGrupoCentroDeCusto
               
               ContaReg = ContaReg + 1
               
               rs.MoveNext
               i = i + 1
            Loop
            rs.Close
            
            MsgBox ("Foram geradas ") & ContaReg & " linhas no Valores a pagar"
            
            
            rs.Open "SELECT chPessoa,chNotaFiscal,chFatura,ctrDataVencito,ctrDataVencitoOriginal,ctrDataRecebimento,ctrValorLart,ctrValorDaBoleta,ctrCentroDeCusto,ctrGrupoCentroDeCusto,ctrSubGrupoCentroDeCusto FROM Contas_A_Receber WHERE ctrStatus=1 AND ctrDataRecebimento>=('" & dataInicio & "') AND ctrDataRecebimento<=('" & dataFim & "')" & _
                    "UNION ALL SELECT chPessoa,chNotaFiscal,chFatura,ctrDataVencito,ctrDataVencOriginal,ctrDataRecebimento,ctrValorLart,ctrValorDaBoleta,ctrCentroDeCusto,ctrGrupoCentroDeCusto,ctrSubGrupoCentroDeCusto FROM HistoricoContasReceber WHERE ctrStatus = 1 AND ctrDataRecebimento>=('" & dataInicio & "') AND ctrDataRecebimento<=('" & dataFim & "') order by ctrDataRecebimento", db, 3, 3
            i = 2
            ContaReg = 0
            
            Do Until rs.EOF
               excelRecebidos.Cells(i, 1) = rs!chPessoa
               excelRecebidos.Cells(i, 2) = rs!chNotafiscal
               excelRecebidos.Cells(i, 3) = rs!chFatura
               excelRecebidos.Cells(i, 4) = rs!ctrDataVencito
               excelRecebidos.Cells(i, 5) = rs!ctrDataVencitoOriginal
               excelRecebidos.Cells(i, 6) = rs!ctrDataRecebimento
               excelRecebidos.Cells(i, 7) = rs!ctrValorLart
               excelRecebidos.Cells(i, 8) = rs!ctrValorDaBoleta
               excelRecebidos.Cells(i, 9) = rs!ctrCentroDeCusto
               excelRecebidos.Cells(i, 10) = rs!ctrGrupoCentroDeCusto
               excelRecebidos.Cells(i, 11) = rs!ctrSubGrupoCentroDeCusto
               
               ContaReg = ContaReg + 1
               
               rs.MoveNext
               i = i + 1
            Loop
            
            MsgBox ("Foram geradas ") & ContaReg & " linhas no Valores Recebidos"
            
            rs.Close
            
            rs.Open "select chPessoa,chNotaFiscal,chFatura,ctrDataVencito,ctrDataVencitoOriginal,ctrDataRecebimento,ctrValorLart,ctrValorDaBoleta,ctrCentroDeCusto,ctrGrupoCentroDeCusto,ctrSubGrupoCentroDeCusto from Contas_A_Receber where ctrDataVencito>=('" & dataInicio & "') and ctrDataVencito<=('" & dataFim & "')" & _
                    "UNION ALL SELECT chPessoa,chNotaFiscal,chFatura,ctrDataVencito,ctrDataVencOriginal,ctrDataRecebimento,ctrValorLart,ctrValorDaBoleta,ctrCentroDeCusto,ctrGrupoCentroDeCusto,ctrSubGrupoCentroDeCusto FROM HistoricoContasReceber WHERE ctrStatus = 1 AND ctrDataVencito>=('" & dataInicio & "') AND ctrDataVencito<=('" & dataFim & "') order by ctrDataVencito", db, 3, 3
            
            i = 2
            ContaReg = 0
            
            Do Until rs.EOF
               excelPrevistos.Cells(i, 1) = rs!chPessoa
               excelPrevistos.Cells(i, 2) = rs!chNotafiscal
               excelPrevistos.Cells(i, 3) = rs!chFatura
               excelPrevistos.Cells(i, 4) = rs!ctrDataVencito
               excelPrevistos.Cells(i, 5) = rs!ctrDataVencitoOriginal
               excelPrevistos.Cells(i, 6) = rs!ctrDataRecebimento
               excelPrevistos.Cells(i, 7) = rs!ctrValorLart
               excelPrevistos.Cells(i, 8) = rs!ctrValorDaBoleta
               excelPrevistos.Cells(i, 9) = rs!ctrCentroDeCusto
               excelPrevistos.Cells(i, 10) = rs!ctrGrupoCentroDeCusto
               excelPrevistos.Cells(i, 11) = rs!ctrSubGrupoCentroDeCusto
               
               ContaReg = ContaReg + 1
               
               rs.MoveNext
               i = i + 1
            Loop
            
            MsgBox ("Foram geradas ") & ContaReg & " linhas no Valores Previstos"
            
            rs.Close
          If cmbAnoFim = cmbAnoInicio + 1 Then
            dif = 12 - cmbMesInicio + cmbMesFim
          Else
            dif = cmbMesFim - cmbMesInicio
          End If
          
          j = 0
          i = 2
          
          Do While j <= dif
            mesAtual = cmbMesInicio + j
            anoAtual = cmbAnoInicio
            If mesAtual > 12 Then
               mesAtual = mesAtual - 12
               anoAtual = cmbAnoFim
            End If
            excelCelula.Cells(i, 3) = mesAtual
            excelCelula.Cells(i, 2) = anoAtual
            excelCelula.Cells(i, 1) = i - 1
            j = j + 1
            i = i + 1
          Loop
110
          FechaDB

490       oWB.SaveAs FileName:=CaminhoNew & "CustoPeriodo1.xlsx"

510       oWB.Close SaveChanges:=False
520       Set oWB = Nothing
530       oApp.Quit
540       Set oApp = Nothing

Exit Sub

End Sub

