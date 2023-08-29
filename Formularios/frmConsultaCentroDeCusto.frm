VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaCentroDeCusto 
   Caption         =   "frmConsultaCentroDeCusto"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19710
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   19710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Seleciona Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   16
      Top             =   6840
      Width           =   12495
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00C0C000&
         Caption         =   "Consulta"
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cmbAno 
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
         Left            =   5280
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cmbMes 
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
         Left            =   3720
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Ano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Mês"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   16440
      TabIndex        =   14
      Top             =   6840
      Width           =   3135
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalhe do Subgrupo de Centro de Custos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   13080
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid grdDetalhe 
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   8281
         _Version        =   393216
         FixedCols       =   0
         FormatString    =   "Detalhe do Subgrupo                                  |Valor              "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label txtTotalDetalhe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Total do Detalhe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   5160
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Custo do Subgrupo de Centro de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   6600
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid grdSubGrupo 
         Height          =   4695
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "Subgrup de Centro de Custo                      |Valor              |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Total Subgrupo de Centro de Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   5160
         Width           =   4335
      End
      Begin VB.Label lblTotalGrupoCentroDeCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   6120
         Width           =   15
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custo Consolidado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid grdCentroDeCusto 
         Height          =   4695
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "Grupo deCentro de Custo                         |Valor             |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTotalCentroDeCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Total Centro DeCusto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   5160
         Width           =   2775
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
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
      Left            =   17280
      TabIndex        =   19
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblHoje 
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
      Left            =   17400
      TabIndex        =   18
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Consulta de Despesas por  Centro de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmConsultaCentroDeCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DescricaoCentroDeCusto(10) As String
Dim ValorCentroDeCusto(10) As Currency
Dim GrupoCentroDeCusto(10) As String
Dim DescricaoGrupoCentroDeCusto(50) As String
Dim SubGrupoCentroDeCusto(50) As String
Dim Ind As Integer
Dim IndTab As Integer
Dim AcumulaCentroDeCusto As Currency
Dim NotaFiscal As String
Dim Pessoa As String
Dim Status As Integer
Dim ValorGrupoCentroDeCusto(50) As Currency
Dim AcumulaGrupoCentroDeCusto As Currency
Dim AcumulaDetalhe As Currency
Dim ChaveGrupo As String
Dim ChaveSubGrupo As String
Dim AnoInicioOperacao As String
Dim ano As Integer
Dim Mes As Integer
Dim AnoHoje As Integer
Dim MesHoje As Integer
Dim Mes12 As String
Dim ChavePeriodo As Integer
Dim InicioPeriodo As String
Dim FimPeriodo As String
Dim DataInicioPeriodo As Date
Dim DataInicioInvertida As Date
Dim Contador As Integer


Private Sub cmdConsulta_Click()
If AnoHoje = cmbAno Then
   If cmbMes > MesHoje Then
      MsgBox ("Mês para consulta inválido. Maior que o mês da data atual."), vbInformation
      Exit Sub
   Else
      If cmbMes = MesHoje Then
         ChavePeriodo = 0
      Else
         ChavePeriodo = 1
      End If
   End If
Else
   ChavePeriodo = 1
End If

lblTotalGrupoCentroDeCusto = 0
txtTotalDetalhe = 0

If grdSubGrupo.Rows > 1 Then
   grdSubGrupo.TextMatrix(0, 0) = "Subgrup de Centro de Custo"
   grdSubGrupo.TextMatrix(1, 0) = Empty
   grdSubGrupo.TextMatrix(1, 1) = Empty
   grdSubGrupo.Rows = 1
End If
If grdDetalhe.Rows > 1 Then
   grdDetalhe.TextMatrix(0, 0) = "Detalhe do Subgrupo"
   grdDetalhe.TextMatrix(1, 0) = Empty
   grdDetalhe.TextMatrix(1, 1) = Empty
   grdDetalhe.Rows = 1
End If
ano = cmbAno
Mes = Format$(cmbMes, "00")

Call GeraDataInicioDataFim

Call CarregaGridCentroDeCusto

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

AnoHoje = Year(Date)
MesHoje = Month(Date)
lblHoje = Date

AnoInicioOperacao = 2022

ano = Year(Date)

For Ind = 1 To 12
    cmbMes.AddItem Format$(Ind, "00")
Next

Do While (ano + 1) > AnoInicioOperacao
   cmbAno.AddItem ano
   ano = ano - 1
Loop

Ind = Month(Date)

cmbMes.ListIndex = Ind - 1

cmbAno.ListIndex = 0

Call CarregaGridCentroDeCusto

End Sub

Public Sub VerificaStatus()

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If

Status = 0

If ChavePeriodo = 0 Then
   ctp.Open "Select * from Contas_A_Pagar where chPessoa = ('" & Pessoa & "') and chNotaFiscal = ('" & NotaFiscal & "') and ctpStatus = 1", db, 3, 3
   If Not ctp.EOF Then
      Status = ctp!ctpStatus
'      If ctp!ctpStatus = 0 Then
'         MsgBox ("Valor - ") & ctp!ctpValorDaBoleta
'      End If
   Else
      Status = 0
   End If
Else
   ctp.Open "Select * from HistoricoContasPagar where chPessoa = ('" & Pessoa & "') and chNotaFiscal = ('" & NotaFiscal & "')", db, 3, 3
   If Not ctp.EOF Then
      Status = ctp!ctpStatus
   Else
      Status = 0
   End If
End If

End Sub

Private Sub grdCentroDeCusto_Click()

If grdCentroDeCusto.TextMatrix(grdCentroDeCusto.Row, 0) = Empty Then
   MsgBox ("Clicar somente em linha com conteúdo."), vbInformation
   Exit Sub
End If

Call Rotina_AbrirBanco

For Ind = 0 To 50
   ValorGrupoCentroDeCusto(Ind) = 0
Next

If grdDetalhe.Rows > 1 Then
   grdDetalhe.TextMatrix(0, 0) = "Detalhe do Subgrupo"
   grdDetalhe.TextMatrix(1, 0) = Empty
   grdDetalhe.TextMatrix(1, 1) = Empty
   grdDetalhe.Rows = 1
End If

txtTotalDetalhe = 0

AcumulaGrupoCentroDeCusto = 0

ChaveGrupo = grdCentroDeCusto.TextMatrix(grdCentroDeCusto.Row, 2)

ccc.Open "Select * from CentroDeCusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ChaveGrupo & "') and chSubGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
If ccc.EOF Then
   MsgBox ("ERRO: Tabela de grupo de centro de custo Vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

ccc.MoveFirst

Do While Not ccc.EOF
   DescricaoGrupoCentroDeCusto(ccc!chSubGrupoCentroDeCusto) = ccc!DescricaoCentroDeCusto
   SubGrupoCentroDeCusto(ccc!chSubGrupoCentroDeCusto) = ccc!chSubGrupoCentroDeCusto
   ccc.MoveNext
Loop

   If ChavePeriodo = 0 Then
      If Prod.State = 1 Then
         Prod.Close: Set Prod = Nothing
      End If
      Prod.Open "Select * from NotaFiscalDetProd where nfdCentroDeCusto = ('" & "2" & "') and nfdGrupoCentroDeCusto = ('" & ChaveGrupo & "') and nfdSubGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
      If Prod.EOF Then
         Contador = Contador + 1
      End If
   Else
      If Prod.State = 1 Then
         Prod.Close: Set Prod = Nothing
      End If
      Prod.Open "Select * from HistoricoNotaFiscalDetProd where nfdDataPagamento > ('" & InicioPeriodo & "') and nfdDataPagamento < ('" & FimPeriodo & "') and nfdCentroDeCusto = ('" & "2" & "') and nfdGrupoCentroDeCusto = ('" & ChaveGrupo & "') and nfdSubGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
      If Prod.EOF Then
         MsgBox ("Nota fiscal sem movimento no período solicitado."), vbInformation
         Call FechaDB
         Exit Sub
      End If
   End If

   Prod.MoveFirst

   Do While Not Prod.EOF
      NotaFiscal = Prod!chNotaFiscalEntrada
      Pessoa = Prod!chPessoa
      Call VerificaStatus
      If Status = 1 Then
         ValorGrupoCentroDeCusto(Prod!nfdSubGrupoCentroDeCusto) = ValorGrupoCentroDeCusto(Prod!nfdSubGrupoCentroDeCusto) + Prod!nfdValorParcela
      End If
     
      Prod.MoveNext

   Loop

grdSubGrupo.Rows = 1
Ind = 1

grdSubGrupo.TextMatrix(0, 0) = grdCentroDeCusto.TextMatrix(grdCentroDeCusto.Row, 0)

For IndTab = 0 To 50
    If ValorGrupoCentroDeCusto(IndTab) > 0 Then
       grdSubGrupo.Rows = Ind + 1
       grdSubGrupo.TextMatrix(Ind, 0) = DescricaoGrupoCentroDeCusto(IndTab)
       grdSubGrupo.TextMatrix(Ind, 1) = Format$(ValorGrupoCentroDeCusto(IndTab), "##,###,##0.00")
       grdSubGrupo.TextMatrix(Ind, 2) = SubGrupoCentroDeCusto(IndTab)
       AcumulaGrupoCentroDeCusto = AcumulaGrupoCentroDeCusto + ValorGrupoCentroDeCusto(IndTab)
       Ind = Ind + 1
    End If
Next

lblTotalGrupoCentroDeCusto = Format$(AcumulaGrupoCentroDeCusto, "##,###,##0.00")

End Sub

Private Sub grdSubGrupo_Click()

If grdSubGrupo.TextMatrix(1, 0) = Empty Then
   MsgBox ("Clicar somente em linha com conteúdo."), vbInformation
   Exit Sub
End If

ChaveGrupo = grdCentroDeCusto.TextMatrix(grdCentroDeCusto.Row, 2)
ChaveSubGrupo = grdSubGrupo.TextMatrix(grdSubGrupo.Row, 2)

Call Rotina_AbrirBanco

   If ChavePeriodo = 0 Then
      If Prod.State = 1 Then
         Prod.Close: Set Prod = Nothing
      End If
      Prod.Open "Select * from NotaFiscalDetProd where nfdCentroDeCusto = ('" & "2" & "') and nfdGrupoCentroDeCusto = ('" & ChaveGrupo & "') and nfdSubGrupoCentroDeCusto = ('" & ChaveSubGrupo & "')", db, 3, 3
      If Prod.EOF Then
         MsgBox ("Nota fiscal sem movimento Not período."), vbInformation
      End If
   Else
      If Prod.State = 1 Then
         Prod.Close: Set Prod = Nothing
      End If
      Prod.Open "Select * from HistoricoNotaFiscalDetProd where nfdDataPagamento > ('" & InicioPeriodo & "') and nfdDataPagamento < ('" & FimPeriodo & "') and nfdCentroDeCusto = ('" & "2" & "') and nfdGrupoCentroDeCusto = ('" & ChaveGrupo & "') and nfdSubGrupoCentroDeCusto = ('" & ChaveSubGrupo & "')", db, 3, 3
      If Prod.EOF Then
         MsgBox ("Nota fiscal sem movimento no período solicitado."), vbInformation
         Call FechaDB
         Exit Sub
      End If
   End If

   grdDetalhe.Rows = 1
   Ind = 1
   AcumulaDetalhe = 0
   
   Prod.MoveFirst
 
   grdDetalhe.TextMatrix(0, 0) = grdSubGrupo.TextMatrix(grdSubGrupo.Row, 0)

   Do While Not Prod.EOF
      NotaFiscal = Prod!chNotaFiscalEntrada
      Pessoa = Prod!chPessoa
      Call VerificaStatus
      If Status = 1 Then
         AcumulaDetalhe = AcumulaDetalhe + Prod!nfdValorParcela
         grdDetalhe.Rows = Ind + 1
         grdDetalhe.TextMatrix(Ind, 0) = Prod!chCodProduto
         grdDetalhe.TextMatrix(Ind, 1) = Format$(Prod!nfdValorParcela, "###,##0.00")
         Ind = Ind + 1
      End If
      Prod.MoveNext
   Loop

txtTotalDetalhe = Format$(AcumulaDetalhe, "##,###,##0.00")


End Sub

Public Sub CarregaGridCentroDeCusto()

Call Rotina_AbrirBanco

For Ind = 0 To 10
   ValorCentroDeCusto(Ind) = 0
Next

AcumulaCentroDeCusto = 0

ccc.Open "Select * from CentroDeCusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto > ('" & "00" & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If ccc.EOF Then
   MsgBox ("ERRO: Tabela de centro de custo Vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

ccc.MoveFirst

Do While Not ccc.EOF
   DescricaoCentroDeCusto(ccc!chGrupoCentroDeCusto) = ccc!DescricaoCentroDeCusto
   GrupoCentroDeCusto(ccc!chGrupoCentroDeCusto) = ccc!chGrupoCentroDeCusto
   ccc.MoveNext
Loop



If ChavePeriodo = 0 Then
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   Prod.Open "Select * from NotaFiscalDetProd where nfdCentroDeCusto = ('" & "2" & "') and nfdGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("Nota fiscal sem movimento Not período."), vbInformation
      Call FechaDB
      Exit Sub
    End If
Else
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   Prod.Open "Select * from HistoricoNotaFiscalDetProd where nfdDataPagamento > ('" & InicioPeriodo & "') and nfdDataPagamento < ('" & FimPeriodo & "') and nfdCentroDeCusto = ('" & "2" & "') and nfdGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("Nota fiscal sem movimento no período solicitado."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

Prod.MoveFirst

Do While Not Prod.EOF
   NotaFiscal = Prod!chNotaFiscalEntrada
   Pessoa = Prod!chPessoa
   Call VerificaStatus
   If Status = 1 Then
      ValorCentroDeCusto(Prod!nfdGrupoCentroDeCusto) = ValorCentroDeCusto(Prod!nfdGrupoCentroDeCusto) + Prod!nfdValorParcela
   End If
   Prod.MoveNext
Loop

grdCentroDeCusto.Rows = 1

AcumulaCentroDeCusto = 0

For Ind = 0 To 10
    If Not DescricaoCentroDeCusto(Ind) = Empty Then
       grdCentroDeCusto.Rows = Ind + 1
       grdCentroDeCusto.TextMatrix(Ind, 0) = DescricaoCentroDeCusto(Ind)
       grdCentroDeCusto.TextMatrix(Ind, 1) = Format$(ValorCentroDeCusto(Ind), "##,###,##0.00")
       grdCentroDeCusto.TextMatrix(Ind, 2) = GrupoCentroDeCusto(Ind)
       AcumulaCentroDeCusto = AcumulaCentroDeCusto + ValorCentroDeCusto(Ind)
    End If
Next

'ctp.MoveNext
   
lblTotalCentroDeCusto = Format$(AcumulaCentroDeCusto, "##,###,##0.00")

End Sub

Public Sub GeraDataInicioDataFim()
Dim MesProximo As Integer

Dia = Format$(1, "00")

DataInicioInvertida = Format$(ano & "-" & Mes & "-" & Dia, "dd/mm/yyyy")

DataInicioInvertida = DataInicioInvertida - 1

InicioPeriodo = Year(DataInicioInvertida) & "-" & Format$(Month(DataInicioInvertida), "00") & "-" & Format$(Day(DataInicioInvertida), "00")

MesProximo = Format$(Mes + 1, "00")
Dia = Format$(1, "00")

If MesProximo > 12 Then
   MesProximo = 1
   ano = ano + 1
End If

FimPeriodo = ano & "-" & Format$(MesProximo, "00") & "-" & Format$(Dia, "00")

End Sub


Public Sub RotinaOriginal()
Call Rotina_AbrirBanco

For Ind = 0 To 10
   ValorCentroDeCusto(Ind) = 0
Next

AcumulaCentroDeCusto = 0

ccc.Open "Select * from CentroDeCusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto > ('" & "00" & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If ccc.EOF Then
   MsgBox ("ERRO: Tabela de centro de custo Vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

ccc.MoveFirst

Do While Not ccc.EOF
   DescricaoCentroDeCusto(ccc!chGrupoCentroDeCusto) = ccc!DescricaoCentroDeCusto
   GrupoCentroDeCusto(ccc!chGrupoCentroDeCusto) = ccc!chGrupoCentroDeCusto
   ccc.MoveNext
Loop


If ChavePeriodo = 0 Then
   Prod.Open "Select * from NotaFiscalDetProd", db, 3, 3
   If Prod.EOF Then
      MsgBox ("Nota fiscal sem movimento Not período."), vbInformation
      Call FechaDB
      Exit Sub
    End If
Else
   Prod.Open "Select * from HistoricoNotaFiscalDetProd where nfdDataPagamento > ('" & InicioPeriodo & "') and nfdDataPagamento < ('" & FimPeriodo & "')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("Nota fiscal sem movimento no período solicitado."), vbInformation
      Call FechaDB
      Exit Sub
      
   End If
End If

Prod.MoveFirst

Do While Not Prod.EOF
   NotaFiscal = Prod!chNotaFiscalEntrada
   Pessoa = Prod!chPessoa
   If ChavePeriodo = 0 Then
      Call VerificaStatus
   Else
      Status = 1
   End If
   If Status = 1 Then
      ValorCentroDeCusto(Prod!nfdGrupoCentroDeCusto) = ValorCentroDeCusto(Prod!nfdGrupoCentroDeCusto) + Prod!nfdValorParcela
   End If
   Prod.MoveNext
Loop

grdCentroDeCusto.Rows = 1

For Ind = 0 To 10
    If Not DescricaoCentroDeCusto(Ind) = Empty Then
       grdCentroDeCusto.Rows = Ind + 1
       grdCentroDeCusto.TextMatrix(Ind, 0) = DescricaoCentroDeCusto(Ind)
       grdCentroDeCusto.TextMatrix(Ind, 1) = Format$(ValorCentroDeCusto(Ind), "##,###,##0.00")
       grdCentroDeCusto.TextMatrix(Ind, 2) = GrupoCentroDeCusto(Ind)
       AcumulaCentroDeCusto = AcumulaCentroDeCusto + ValorCentroDeCusto(Ind)
    End If
Next


lblTotalCentroDeCusto = Format$(AcumulaCentroDeCusto, "##,###,##0.00")

End Sub
