VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCtaPagar 
   Caption         =   "frmCtaPagar"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17010
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   9795
   ScaleWidth      =   17010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaldo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   17
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox txtTotalPago 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   16
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox txtTotalAReceber 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdCtaRecebidas 
      BackColor       =   &H00C0C000&
      Caption         =   "Imprime Contas a Receber"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton drCtaPagarReceber 
      BackColor       =   &H008080FF&
      Caption         =   "Imprime Contas a Pagar"
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   1815
   End
   Begin VB.PictureBox GridCtaPag 
      Appearance      =   0  'Flat
      DataMember      =   "cmdCtaPag"
      Height          =   6495
      Left            =   8280
      ScaleHeight     =   6465
      ScaleWidth      =   8505
      TabIndex        =   8
      Top             =   1080
      Width           =   8535
      Begin MSFlexGridLib.MSFlexGrid gridCtaPagas 
         Height          =   6255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   11033
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777088
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FormatString    =   "|Data Pagto |Descrição                        |Num. Doc.    |Valor Pago "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H000000FF&
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
      Height          =   1095
      Left            =   14640
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.PictureBox GridCtaReceber 
      Appearance      =   0  'Flat
      DataMember      =   "cmdCtaRec"
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6465
      ScaleWidth      =   8265
      TabIndex        =   7
      Top             =   1080
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid gridCtaRecebidas 
         Height          =   6375
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11245
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777088
         BackColorBkg    =   16777152
         FormatString    =   "|Data Receb.|Operação        |Fatura/Doc        |Valor Recebido"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtHoje 
      Height          =   495
      Left            =   14640
      TabIndex        =   14
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Saldo Consultado"
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
      Left            =   12120
      TabIndex        =   11
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Total a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   10
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Total a Receber"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Contas a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   6
      Top             =   720
      Width           =   2130
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Contas a Receber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contas a Pagar e a Receber"
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
      TabIndex        =   4
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   14880
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCtaPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim tabValor(15) As Currency
Dim tabDesc(15) As String

Dim Resp As String
Dim Linha As Single
Dim DataVenc As Date
Dim IndConf As Integer
Dim AcumulaRecebidas As Currency
Dim AcumulaPagar As Currency
Dim TotalCentroDeCusto As Currency
Dim indice As Byte
Dim DataAcesso As Date
Dim PrimeiraVez As Byte

'Area de trabalho para gerar grid

'Contas a Receber
Dim wsctrDataRecebimento As Date
Dim wsctpchPessoa As String
Dim wsctrchNotaFiscal As String
Dim wsctrvalordaboleta As Currency

'Contas a Pagar

Dim wsctpdatapagamento As Date
Dim wsctrchPessoa As String
Dim wsctpchNotaFiscal As String
Dim wsctpvalordaboleta As Currency

Dim DataInicioInvertida As String
Dim DataFinalInvertida As String
Dim DataHoje As Date
Dim Relatorio As String

Dim dtInicio As Date
Dim dtFim As Date

Private Sub cmdConsulta_Click()

Call LimparGrid

Call RotinaInverterData

Call Rotina_AbrirBanco

hctr.Open "Select * from historicocontasreceber where ctrDataRecebimento > ('" & DataInicioInvertida & "') and ctrDataRecebimento < ('" & DataFinalInvertida & "')", db, 3, 3
If hctr.EOF Then
   MsgBox ("Sem Lançamentos a credito confirmados no período informado"), vbInformation
Else
   hctr.MoveFirst
   Call GerenciaContasaReceberPeriodo
End If

hctp.Open "Select * from historicocontaspagar where ctpDataPagamento > ('" & DataInicioInvertida & "') and ctpDataPagamento < ('" & DataFinalInvertida & "')", db, 3, 3
If hctp.EOF Then
   MsgBox ("Sem Lançamentos a débito confirmados no período informado"), vbInformation
Else
   hctp.MoveFirst
   Call GerenciaContasaPagarPeriodo
End If

txtSaldo = Format$(AcumulaRecebidas - AcumulaPagar, "##,##0.00")

If AcumulaPagar > AcumulaRecebidas Then
   txtSaldo.ForeColor = vbRed
Else
   txtSaldo.ForeColor = vbBlue
End If


End Sub

Private Sub cmdCtaRecebidas_Click()
Dim sql As String
Dim Rel As Object

Call RotinaInverterData

'If optMesAtual = True Then

   Call GeraDataInicioDataFim

   Call Rotina_AbrirBanco
   Relatorio = "drCtaReceber"
   db.BeginTrans
   gge.Open "Select * from geradorgeral where chAlfaNumerica = ('" & Relatorio & " ')", db, 3, 3
   If gge.EOF Then
      gge.AddNew
   End If

   gge!chAlfaNumerica = Relatorio
   gge!ggeDataHoje = Date
   gge!ggeDataIni = DataInicioInvertida
   gge!ggeDataFim = DataFinalInvertida
   gge.Update

   db.CommitTrans

   Set Rel = drCtaReceber
   sql = "Select gge.ggeDataHoje, gge.ggeDataIni, gge.ggeDataFim, ctr.ctrDataVencito, ctr.chNotaFiscal, pes.pesRazaoSocial, ctr.ctrValorDaBoleta "
   sql = sql & " From contas_a_receber ctr, geradorgeral gge, pessoa pes "
   sql = sql & "WHERE ctr.ctrStatus = 0 and gge.chAlfaNumerica = ('" & Relatorio & "') and ctr.chPessoa = pes.chPessoa order by ctr.ctrDataVencito"

AbrirRelatorio sql, Rel
 
End Sub

Private Sub drCtaPagarReceber_Click()

Dim sql As String
Dim Rel As Object
Dim Sec As Integer
Dim Rel1 As String

Call Rotina_AbrirBanco

db.BeginTrans

   Call GeraDataInicioDataFim
   
   db.Execute ("DELETE FROM geradorgeral WHERE chAlfaNumerica = 'drCtaPagar'")
   db.Execute ("INSERT INTO geradorgeral (chAlfaNumerica, ggeDataHoje,ggeDataIni, ggeDataFim) " & _
  "VALUES('drCtaPagar', '" & Format(DataHoje, "yyyy-MM-dd") & "','" & DataInicioInvertida & "','" & DataFinalInvertida & "')")

db.CommitTrans
'Para ser usado somente quando houver erro na busca por ser muito longa
'db.Execute "SET SQL_BIG_SELECTS=1"
Set Rel = drCtaPagar
sql = "Select gge.ggeDataHoje, gge.ggeDataIni, gge.ggeDataFim, ctp.chDataVencito, ctp.chNotaFiscal, ctp.chPessoa, ctp.ctpDescricaoOperacao, ctp.ctpValorDaBoleta "
sql = sql & " FROM geradorgeral gge, contas_a_pagar ctp "
sql = sql & "WHERE gge.chAlfaNumerica = 'drCtaPagar' AND ctp.ctpStatus = 0 ORDER BY ctp.chDataVencito, ctp.chPessoa"
   
AbrirRelatorio sql, Rel

'MsgBox ("Relatorio")
Call FechaDB


End Sub

Private Sub cmdSair_Click()

Unload Me

End Sub

Private Sub dtFim_LostFocus()
If Not Month(dtInicio) = Month(dtFim) Then
   MsgBox ("O período tem que iniciar e finalizar no mesmo mês"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If
End Sub

Private Sub dtInicio_LostFocus()
If Month(dtInicio) = Month(Date) And Year(dtInicio) = Year(Date) Then
   MsgBox ("Data início e fim tem que ser anterior a data atual"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

Call RotinaInverterData

End Sub



Private Sub Form_Load()

txtHoje = Date
dtInicio = Date
dtFim = Date

PrimeiraVez = 1

Call LimpaTabela

Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber where ctrStatus = ('" & 0 & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Sem Lançamentos a credito confirmados até a presente data"), vbInformation
Else
   ctr.MoveFirst
   Call GerenciaContasaReceber
End If

ctp.Open "Select * from contas_a_pagar where ctpStatus = ('" & 0 & "')", db, 3, 3
If ctp.EOF Then
   MsgBox ("Sem Lançamentos a débito confirmados até a presente data"), vbInformation
Else
   ctp.MoveFirst
   Call GerenciaContasaPagar
End If

txtSaldo = Format$(AcumulaRecebidas - AcumulaPagar, "##,##0.00")

If AcumulaPagar > AcumulaRecebidas Then
   txtSaldo.ForeColor = vbRed
Else
   txtSaldo.ForeColor = vbBlue
End If

'optMesAtual = True
'optMesAnter = False
'fraPeriodo.Visible = False

usu.Open "Select * from usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Erro no acesso a usuario."), vbCritical
   Exit Sub
End If
   
'If usu!usuRelAnalitico = 1 Then
'   drCtaPagarReceber.Enabled = True
'Else
'   drCtaPagarReceber.Enabled = False
'End If


frmCtaPagar.Show

Call FechaDB

End Sub

Public Sub GerenciaContasaReceber()

IndConf = 0
AcumulaRecebidas = 0

Do While Not ctr.EOF
   
   If ctr("ctrstatus") = 0 Then
      wsctrDataRecebimento = ctr!ctrDataVencito
      wsctrchPessoa = ctr!chPessoa
      wsctrchNotaFiscal = ctr!chNotafiscal
      wsctrvalordaboleta = ctr!ctrValorDaBoleta

      Call CarregaContasaReceber
      AcumulaRecebidas = AcumulaRecebidas + ctr!ctrValorDaBoleta
   End If

   ctr.MoveNext

Loop

gridCtaRecebidas.Col = 1
gridCtaRecebidas.ColSel = 1
     
gridCtaRecebidas.Row = 1
'gridCtaRecebidas.RowSel = IndConf
        
If IndConf > 1 Then
   gridCtaRecebidas.Sort = 1
End If

gridCtaRecebidas.Col = 0
gridCtaRecebidas.ColSel = 0
gridCtaRecebidas.Row = 0
gridCtaRecebidas.RowSel = 0

txtTotalAReceber = Format$(AcumulaRecebidas, "##,##0.00")


End Sub

Public Sub CarregaContasaReceber()

gridCtaRecebidas.Row = 1

IndConf = IndConf + 1

gridCtaRecebidas.Rows = IndConf + 1
       
gridCtaRecebidas.TextMatrix(IndConf, 1) = wsctrDataRecebimento
gridCtaRecebidas.TextMatrix(IndConf, 2) = wsctrchPessoa
gridCtaRecebidas.TextMatrix(IndConf, 3) = wsctrchNotaFiscal
gridCtaRecebidas.TextMatrix(IndConf, 4) = Format(wsctrvalordaboleta, "#,###.00")

End Sub

Public Sub GerenciaContasaPagar()
Dim Status As Integer
Dim IndSalvo As Integer

IndConf = 0
AcumulaPagar = 0

Do While Not ctp.EOF
   
   If ctp!ctpStatus = 0 Then
   
      wsctpdatapagamento = ctp!chDataVencito
      wsctpchPessoa = ctp!chPessoa
      wsctpchNotaFiscal = ctp!chNotafiscal
      wsctpvalordaboleta = ctp!ctpValorDaBoleta
      
      Call CarregaContasaPagar
      AcumulaPagar = AcumulaPagar + ctp!ctpValorDaBoleta
   End If
   
   ctp.MoveNext

Loop

gridCtaPagas.Col = 1
gridCtaPagas.ColSel = 1
     
gridCtaPagas.Row = 1
gridCtaPagas.RowSel = IndConf
        
If IndConf > 1 Then
   gridCtaPagas.Sort = 1
End If

TotalCentroDeCusto = 0

For indice = 1 To 15
    tabDesc(indice) = Empty
    tabValor(indice) = 0
Next

gridCtaPagas.Col = 0
gridCtaPagas.ColSel = 0
gridCtaPagas.Row = 0
gridCtaPagas.RowSel = 0

txtTotalPago = Format$(AcumulaPagar, "##,##0.00")

End Sub

Public Sub CarregaContasaPagar()

gridCtaPagas.Row = 1

IndConf = IndConf + 1

gridCtaPagas.Rows = IndConf + 1

gridCtaPagas.TextMatrix(IndConf, 1) = wsctpdatapagamento
gridCtaPagas.TextMatrix(IndConf, 2) = wsctpchPessoa
gridCtaPagas.TextMatrix(IndConf, 3) = wsctpchNotaFiscal
gridCtaPagas.TextMatrix(IndConf, 4) = Format(wsctpvalordaboleta, "#,###.00")

End Sub

'Private Sub optMesAnter_Click()
'optMesAtual = False
'optMesAnter = True
'Call LimparGrid

'fraPeriodo.Visible = True
'End Sub

'Private Sub optMesAtual_Click()
'optMesAnter = False
'optMesAtual = True
'fraPeriodo.Visible = False

'If PrimeiraVez = 1 Then
'   PrimeiraVez = 0
'   Exit Sub
'End If

'Call LimparGrid
'
'Call Rotina_AbrirBanco
'
'ctr.Open "Select * from contas_a_receber where ctrStatus = ('" & 1 & "')", db, 3, 3
'If ctr.EOF Then
'   MsgBox ("Sem Lançamentos a credito confirmados até a presente data"), vbInformation
'Else
'   ctr.MoveFirst
'   Call GerenciaContasaReceber
'End If

'ctp.Open "Select * from contas_a_pagar where ctpStatus = ('" & 1 & "')", db, 3, 3
'If ctp.EOF Then
'   MsgBox ("Sem Lançamentos a débito confirmados até a presente data"), vbInformation
'Else
'   ctp.MoveFirst
'   Call GerenciaContasaPagar
'End If
'
'txtSaldo = Format$(AcumulaRecebidas - AcumulaPagar, "##,##0.00")
'
'If AcumulaPagar > AcumulaRecebidas Then
'   txtSaldo.ForeColor = vbRed
'Else
'   txtSaldo.ForeColor = vbBlue
'End If
'
'optMesAtual = True
'optMesAnter = False
'fraPeriodo.Visible = False
'
'frmCtaReceb.Show
'
'Call FechaDB
'
'
'End Sub

Public Sub GerenciaContasaReceberPeriodo()
IndConf = 0
AcumulaRecebidas = 0

Do While Not hctr.EOF
   
   wsctrDataRecebimento = hctr!ctrDataRecebimento
   wsctrchPessoa = hctr!chPessoa
   wsctrchNotaFiscal = hctr!chNotafiscal
   wsctrvalordaboleta = hctr!ctrValorDaBoleta

   Call CarregaContasaReceber
   AcumulaRecebidas = AcumulaRecebidas + hctr!ctrValorDaBoleta
   

   hctr.MoveNext

Loop

gridCtaRecebidas.Col = 1
gridCtaRecebidas.ColSel = 1
     
gridCtaRecebidas.Row = 1
'gridCtaRecebidas.RowSel = IndConf
        
If IndConf > 1 Then
   gridCtaRecebidas.Sort = 1
End If

gridCtaRecebidas.Col = 0
gridCtaRecebidas.ColSel = 0
gridCtaRecebidas.Row = 0
gridCtaRecebidas.RowSel = 0

txtTotalAReceber = Format$(AcumulaRecebidas, "##,##0.00")


End Sub

Public Sub GerenciaContasaPagarPeriodo()
IndConf = 0
AcumulaPagar = 0

Do While Not hctp.EOF
 
   wsctpdatapagamento = hctp!ctpDataPagamento
   wsctpchPessoa = hctp!chPessoa
   wsctpchNotaFiscal = hctp!chNotafiscal
   wsctpvalordaboleta = hctp!ctpValorDaBoleta
   
   Call CarregaContasaPagar
   
   AcumulaPagar = AcumulaPagar + hctp!ctpValorDaBoleta

   hctp.MoveNext

Loop

gridCtaPagas.Col = 1
gridCtaPagas.ColSel = 1
     
gridCtaPagas.Row = 1
gridCtaPagas.RowSel = IndConf
        
If IndConf > 1 Then
   gridCtaPagas.Sort = 1
End If

gridCtaPagas.Col = 0
gridCtaPagas.ColSel = 0
gridCtaPagas.Row = 0
gridCtaPagas.RowSel = 0

txtTotalPago = Format$(AcumulaPagar, "##,##0.00")
End Sub

Public Sub RotinaInverterData()
DataInicioInvertida = Format$((dtInicio - 1), "yyyy-mm-dd")
DataFinalInvertida = Format$((dtFim + 1), "yyyy-mm-dd")
End Sub

Public Sub LimparGrid()

Dim Ind As Integer

gridCtaRecebidas.Rows = 2

Ind = 1
gridCtaRecebidas.TextMatrix(Ind, 0) = Empty
gridCtaRecebidas.TextMatrix(Ind, 1) = Empty
gridCtaRecebidas.TextMatrix(Ind, 2) = Empty
gridCtaRecebidas.TextMatrix(Ind, 3) = Empty
gridCtaRecebidas.TextMatrix(Ind, 4) = Empty
txtTotalAReceber = Empty

AcumulaRecebidas = 0

gridCtaPagas.Rows = 2

gridCtaPagas.TextMatrix(Ind, 0) = Empty
gridCtaPagas.TextMatrix(Ind, 1) = Empty
gridCtaPagas.TextMatrix(Ind, 2) = Empty
gridCtaPagas.TextMatrix(Ind, 3) = Empty
gridCtaPagas.TextMatrix(Ind, 4) = Empty
txtTotalPago = Empty

AcumulaPagar = 0

txtSaldo = Empty

End Sub

Public Sub GeraDataInicioDataFim()
Dim MesProximo As Integer

mes = Format$(Month(Date), "00")
ano = Year(Date)
Dia = Format$(1, "00")

DataInicioInvertida = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")

MesProximo = Format$(mes, "00")
DataHoje = Date
Do While mes = MesProximo
   DataHoje = DataHoje + 1
   MesProximo = Format$(Month(DataHoje), "00")
Loop
DataHoje = DataHoje - 1
DataFinalInvertida = Format$(DataHoje, "yyyy-mm-dd")
DataHoje = Date

End Sub

Public Sub LimpaTabela()
For indice = 1 To 15
    tabValor(indice) = 0
    tabDesc(indice) = Empty
Next
End Sub

