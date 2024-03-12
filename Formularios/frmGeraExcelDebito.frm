VERSION 5.00
Begin VB.Form frmGeraExcelDebito 
   Caption         =   "frmGeraExcelDebito"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Parâmetros"
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   12135
      Begin VB.ComboBox cmbAno 
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
         Left            =   3360
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cmbMes 
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
         Left            =   1560
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdGerarExcel 
         BackColor       =   &H00FFFF00&
         Caption         =   "Gerar"
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
   End
   Begin VB.Label Label3 
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
      Left            =   9840
      TabIndex        =   6
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblHoje 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Gerar Excel Para Contabilidade - Débito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmGeraExcelDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Ind As Integer
Dim ano As Integer
Dim mes As Integer
Dim Dia As Integer
Dim AnoInicioOperacao
Dim DataHoje As Date
Dim AnoHoje As Integer
Dim MesHoje As Integer
Dim DataInicioInvertida As String
Dim DataInicioOperacao As Date
Dim DataFinalInvertida As String
Dim MesProximo As Integer
Dim Periodo As Integer
Dim ValorAcumulado As Currency

Dim TipoDeLancamento As String

Dim dataInicio As String
Dim dataFim As String

Private Sub cmdGerarExcel_Click()
Dim contabilidadepagto As String
Call DeletaTabExcel

ValorAcumulado = 0

If AnoHoje = cmbAno Then
   If cmbMes > MesHoje Then
      MsgBox ("Mês para consulta inválido. Maior que o mês da data atual."), vbInformation
      Exit Sub
   End If
End If

Call CriaDatasPesquisa

If Year(DataInicioOperacao) < Year(Date) Then
   Periodo = 1
Else
   If Month(DataInicioOperacao) < Month(Date) Then
      Periodo = 1
   Else
      Periodo = 0
   End If
End If

Call Rotina_AbrirBanco

If Periodo = 0 Then
   ctp.Open "Select * from contas_a_pagar where ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & DataInicioInvertida & "') and ctpDataPagamento < ('" & DataFinalInvertida & "')", db, 3, 3
Else
   ctp.Open "Select * from historicocontaspagar where ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & DataInicioInvertida & "') and ctpDataPagamento < ('" & DataFinalInvertida & "')", db, 3, 3
End If

If ctp.EOF Then
   MsgBox ("Sem Contas a Pagar para o periodo solicitado"), vbInformation
   Call FechaDB
   Exit Sub
End If

ctp.MoveFirst

Do While Not ctp.EOF

   If contab.State = 1 Then
      contab.Close: Set contab = Nothing
      acContab = 0
   End If
   
   contab.Open "Select * from contabilidadepagto where fornecedorSacado = ('" & ctp!chPessoa & "') and NDocumento = ('" & ctp!chNotaFiscal & "')", db, 3, 3
   If contab.EOF Then

      contab.AddNew
   
      contab!TipoLancamento = TipoDeLancamento
      contab!TipoTransacao = ctp!ctpTipoLancamentoDesc
      contab!Descricao = ctp!ctpdescricaooperacao
      contab!FornecedorSacado = ctp!chPessoa
      contab!NDocumento = ctp!chNotaFiscal
      contab!valor = Format$(ctp!ctpValorDaBoleta, "##,##0.00")
      contab!DataPagamento = ctp!ctpDataPagamento
      
      If nfd.State = 1 Then
         nfd.Close: Set nfd = Nothing
      End If
      
      
      If Periodo = 0 Then
         nfd.Open "Select * from notafiscaldetprod where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotaFiscal & "')", db, 3, 3
      Else
         nfd.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotaFiscal & "')", db, 3, 3
      End If
      If nfd.EOF Then
         MsgBox ("Não encontrei o Centro de Custo no Detalhe Produto."), vbInformation
         contab!Categoria = Empty
      Else
         contab!Categoria = nfd!chProdutoFabrica
         contab!Grupo = nfd!nfdCentroDeCusto & nfd!nfdGrupoCentroDeCusto & "00"
         contab!Despesa = nfd!nfdCentroDeCusto & nfd!nfdGrupoCentroDeCusto & nfd!nfdSubGrupoCentroDeCusto
      End If
      
      contab.Update
      
      ValorAcumulado = ValorAcumulado + ctp!ctpValorDaBoleta
      
   End If
   
   ctp.MoveNext
Loop

MsgBox ("Valor Total Gerado = ") & Format$(ValorAcumulado, "##,##0.00")

Call FechaDB

Call ExportarContabilidade

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Periodo = 0
lblHoje = Date

TipoDeLancamento = "DESPESA"

AnoHoje = Year(Date)
MesHoje = Month(Date)
lblHoje = Date

AnoInicioOperacao = 2019

ano = Year(Date)

For Ind = 1 To 12
    cmbMes.AddItem Format$(Ind, "00")
Next

cmbMes.ListIndex = 0

Do While (ano + 1) > AnoInicioOperacao
   cmbAno.AddItem ano
   ano = ano - 1
Loop

cmbAno.ListIndex = 0

End Sub

Public Sub CriaDatasPesquisa()

mes = Format$(cmbMes, "00")
ano = cmbAno
Dia = Format$(1, "00")

DataHoje = Format$(Dia, "00") & "/" & Format$(mes, "00") & "/" & ano
DataInicioOperacao = DataHoje
DataHoje = DataHoje - 1
Dia = Day(DataHoje)
mes = Month(DataHoje)
ano = Year(DataHoje)

DataInicioInvertida = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")

DataHoje = DataHoje + 1
MesProximo = Month(DataHoje)

Do While Month(DataHoje) = MesProximo
   DataHoje = DataHoje + 1
   'MesProximo = Format$(Month(DataHoje), "00")
Loop

DataFinalInvertida = Format$(DataHoje, "yyyy-mm-dd")
DataHoje = Date

End Sub

Public Sub DeletaTabExcel()

Call Rotina_AbrirBanco

contab.Open "Select * from contabilidadepagto", db, 3, 3
If Not contab.EOF Then

   contab.MoveFirst
   
   Do While Not contab.EOF
      contab.Delete
      contab.MoveNext
   Loop
End If

Call FechaDB

End Sub
