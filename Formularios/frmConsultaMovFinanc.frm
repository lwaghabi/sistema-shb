VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaMovFinanc 
   Caption         =   "frmConsultaMovFinanc"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16350
   LinkTopic       =   "Form3"
   ScaleHeight     =   8160
   ScaleWidth      =   16350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   9615
      Begin VB.ComboBox cmbDebCre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7920
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Movimentação Financeira "
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
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   7575
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   9720
      TabIndex        =   13
      Top             =   360
      Width           =   6495
      Begin MSComCtl2.DTPicker txtDataConsulta 
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   124649473
         CurrentDate     =   38966
      End
      Begin VB.CommandButton cmdConsultar 
         BackColor       =   &H00FFFF00&
         Caption         =   "Consultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbFiltro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDataHoje 
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hoje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   15
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resumo Financeiro da Operação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7920
      TabIndex        =   6
      Top             =   6000
      Width           =   8295
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton Command1 
            BackColor       =   &H000000FF&
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
            Height          =   1215
            Left            =   120
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label txtTotalValor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Total Processado no Dia............"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   4620
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "dias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label txtMediaFaturamento 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prazo Médio de Faturamento..............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   4665
      End
      Begin VB.Label txtValorTotalDaOperacao 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Total da Operação....................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4560
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Condições Financeiras da Negociação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   7920
      TabIndex        =   5
      Top             =   1560
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid GridFinanc 
         Height          =   4095
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16777152
         FormatString    =   "Cod. da Fatura    |Vencimento  |Desc Operação    |Valor da Fatura|Status"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Negociações na data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7695
      Begin MSFlexGridLib.MSFlexGrid GridNeg 
         Height          =   4095
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16777152
         FormatString    =   "Cliente                                             |Num Pedido |Comp|Nota Fiscal|"
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
   End
End
Attribute VB_Name = "frmConsultaMovFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Indice As Byte
Dim IndPrz As Byte
Dim IndFab As Byte
Dim IndDebCre As Byte
Dim NotaFiscal As String
Dim AcumulaValor As Currency
Dim AcumulaTotal As Currency
Dim DataHoje As Date
Dim DiaHoje As Integer
Dim DataPrazoMedio As Date
Dim DataConsulta As Date
Dim DataCombo As Date
Dim DiaCombo As Integer
Dim PrazoMedio As Integer
Dim PrazoAdicional As Integer
Dim PessoaAnterior As String
Dim NotaAnterior As String
Dim Pessoa As String



Private Sub cmdConsultar_Click()

Call Rotina_040_Limpa_GridNeg
Call Rotina_045_Limpa_GridFinan

IndFab = cmbFiltro.ListIndex

Rotina_00_Principal

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

DataHoje = Date
txtDataConsulta = Date
DiaHoje = Day(DataHoje)
DiaCombo = 1

txtDataConsulta = Date

cmbFiltro.Clear
cmbFiltro.AddItem "Geral"

Call Rotina_AbrirBanco

Bco.Open "Select * from Banco", db, 3, 3
If Bco.EOF Then
   MsgBox ("tabela de bancos vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If


Bco.MoveFirst

Do While Not Bco.EOF
   cmbFiltro.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbFiltro.ListIndex = 0
IndFab = 0
cmbDebCre.AddItem "Crédito"
cmbDebCre.AddItem "Débito"

cmbDebCre.ListIndex = 0

txtDataHoje = Date

Call Rotina_00_Principal
End Sub

Public Sub Rotina_00_Principal()

Call Rotina_040_Limpa_GridNeg
Call Rotina_045_Limpa_GridFinan

NotaAnterior = Empty
AcumulaTotal = 0
txtTotalValor = Empty
txtValorTotalDaOperacao = Empty
txtMediaFaturamento = Empty

If cmbDebCre.ListIndex = 0 Then
   Call Rotina_010_Credito
Else
   Call Rotina_015_Debito
End If
End Sub

Public Sub Rotina_010_Credito()
If Not Month(txtDataConsulta) = Month(Date) Then
   MsgBox ("Entre com uma data fornecida nesta caixa de rolagem")
   txtDataConsulta.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber", db, 3, 3
If ctr.EOF Then
   MsgBox ("Não há lançamentos a Crédito até a apresente data."), vbInformation
   Call FechaDB
   Exit Sub
End If


ctr.MoveFirst
Indice = 0
Do While Not ctr.EOF
    If Not (ctr!chFatura) = "CREDITO" Then
       If ctr!ctrDataEmissao = txtDataConsulta Then
          If (cmbFiltro = "Geral") Or (cmbFiltro = ctr!chCodBcoLart) Then
             AcumulaTotal = AcumulaTotal + ctr!ctrvalordaboleta
             If ((ctr!chPessoa = PessoaAnterior) And (ctr!chNotaFiscal = NotaAnterior)) Then
                PessoaAnterior = ctr!chPessoa
                NotaAnterior = ctr!chNotaFiscal
             Else
                If neg.State = 1 Then
                   neg.Close: Set neg = Nothing
                End If
                neg.Open "Select * from Negociacao where chNumPedido = ('" & ctr!chNumPedido & "') and chNumPedidoComp = ('" & ctr!chNumPedidoComp & "')", db, 3, 3
                If Not (neg.EOF) Then
                   Indice = Indice + 1
                   GridNeg.Rows = Indice + 1
                   GridNeg.TextMatrix(Indice, 0) = neg!chPessoa
                   GridNeg.TextMatrix(Indice, 1) = neg!chNumPedido
                   GridNeg.TextMatrix(Indice, 2) = neg!chNumPedidoComp
                   GridNeg.TextMatrix(Indice, 3) = neg!negNotaFiscal
                   GridNeg.TextMatrix(Indice, 4) = ctr!chFatura
                   PessoaAnterior = ctr!chPessoa
                   NotaAnterior = ctr!chNotaFiscal
                End If
             End If
          End If
       End If
    End If
    
ctr.MoveNext

Loop

txtTotalValor = Format$(AcumulaTotal, "##,##0.00")

If Indice > 1 Then
   GridNeg.Col = 0
   GridNeg.ColSel = 0
   GridNeg.Row = 1
   GridNeg.RowSel = Indice
   GridNeg.Sort = 1
   GridNeg.Col = 0
   GridNeg.ColSel = 0
   GridNeg.Row = 0
   GridNeg.RowSel = 0
End If

Call FechaDB

End Sub
Public Sub Rotina_015_Debito()
Dim nnnn As String
   
Call Rotina_AbrirBanco

ctp.Open "Select * from Contas_A_Pagar", db, 3, 3
   If ctp.EOF Then
      MsgBox ("Não há contas a pagar até o presente momento"), vbInformation
      Call FechaDB
      Exit Sub
    End If

ctp.MoveFirst
Indice = 0
Do While Not ctp.EOF
If Not (ctp!chFatura = "Comissão") Then
   If ctp!ctpdatalanc = txtDataConsulta Then
      If (cmbFiltro = "Geral") Or (cmbFiltro = ctp!chCodBcoLart) Then
         AcumulaTotal = AcumulaTotal + ctp!ctpvalordaboleta
         If Not ((ctp!chPessoa = PessoaAnterior) And (ctp!chNotaFiscal = NotaAnterior)) Then
            PessoaAnterior = ctp!chPessoa
            NotaAnterior = ctp!chNotaFiscal
            Indice = Indice + 1
            GridNeg.Rows = Indice + 1
            GridNeg.TextMatrix(Indice, 0) = ctp!chPessoa
            GridNeg.TextMatrix(Indice, 1) = "N/Inf."
            GridNeg.TextMatrix(Indice, 2) = "N/Inf."
            GridNeg.TextMatrix(Indice, 3) = ctp!chNotaFiscal
         End If
      End If
   End If
End If
   
ctp.MoveNext
      
Loop
nnnn = GridNeg.TextMatrix(1, 2)
txtTotalValor = Format$(AcumulaTotal, "##,##0.00")

If Indice > 1 Then
   GridNeg.Col = 0
   GridNeg.ColSel = 0
   GridNeg.Row = 1
   GridNeg.RowSel = Indice
   GridNeg.Sort = 1
   GridNeg.Col = 0
   GridNeg.ColSel = 0
   GridNeg.Row = 0
   GridNeg.RowSel = 0
End If

Call FechaDB

End Sub
Private Sub cmbFiltro_lostfocus()

NotaAnterior = Empty
AcumulaTotal = 0
txtTotalValor = Empty
txtValorTotalDaOperacao = Empty
txtMediaFaturamento = Empty

Indice = cmbFiltro.ListIndex

End Sub

Public Sub Rotina_040_Limpa_GridNeg()
GridNeg.Rows = 2
Indice = 1
GridNeg.TextMatrix(Indice, 0) = Empty
GridNeg.TextMatrix(Indice, 1) = Empty
GridNeg.TextMatrix(Indice, 2) = Empty
GridNeg.TextMatrix(Indice, 3) = Empty
GridNeg.TextMatrix(Indice, 4) = Empty

End Sub
Public Sub Rotina_045_Limpa_GridFinan()
GridFinanc.Rows = 2
Indice = 1
GridFinanc.TextMatrix(Indice, 0) = Empty
GridFinanc.TextMatrix(Indice, 1) = Empty
GridFinanc.TextMatrix(Indice, 2) = Empty
GridFinanc.TextMatrix(Indice, 3) = Empty
GridFinanc.TextMatrix(Indice, 4) = Empty
End Sub



Private Sub GridNeg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Rotina_045_Limpa_GridFinan

IndFab = cmbFiltro.ListIndex

IndDebCre = cmbDebCre.ListIndex

If IndFab = 0 Then
   MsgBox ("Favor utilizar o FILTRO e indicar o Banco de faturamento")
   cmbFiltro.SetFocus
   Exit Sub
End If

If IndDebCre = 1 Then
   Call Rotina_050_Carga_Debito
Else
   Call Rotina_060_Carga_Credito
End If

GridFinanc.Col = 1
GridFinanc.ColSel = 1
GridFinanc.Sort = 5

End Sub

Public Sub Rotina_050_Carga_Debito()
Indice = GridNeg.Row
IndPrz = 0

txtMediaFaturamento = "N/Disp."

NotaFiscal = GridNeg.TextMatrix(Indice, 3)
Pessoa = GridNeg.TextMatrix(Indice, 0)

Indice = 0
AcumulaValor = 0

Call Rotina_AbrirBanco

ctp.Open "Select * from Contas_A_Pagar", db, 3, 3
If ctp.EOF Then
   MsgBox ("Não há contas a pagar até o presente momento."), vbInformation
   Call FechaDB
   Exit Sub
End If


ctp.MoveFirst

Do While Not ctp.EOF
   If ctp!chNotaFiscal = NotaFiscal And ctp!chPessoa = Pessoa Then
      Indice = Indice + 1
      GridFinanc.Rows = Indice + 1
      GridFinanc.TextMatrix(Indice, 0) = ctp!chFatura
      GridFinanc.TextMatrix(Indice, 1) = ctp!chdatavencito
      GridFinanc.TextMatrix(Indice, 2) = ctp!ctpdescricaooperacao
      GridFinanc.TextMatrix(Indice, 3) = Format$(ctp!ctpvalordaboleta, "##,##0.00")
      If ctp!ctpstatus = 1 Then
         GridFinanc.TextMatrix(Indice, 4) = "Pg."
      Else
         If ctp!chdatavencito < Date - 1 Then
            GridFinanc.TextMatrix(Indice, 4) = "Atr."
         Else
            GridFinanc.TextMatrix(Indice, 4) = Empty
         End If
      End If
      AcumulaValor = AcumulaValor + Format$(ctp!ctpvalordaboleta, "##,##0.00")
   End If
      
   ctp.MoveNext

Loop
txtValorTotalDaOperacao = Format$(AcumulaValor, "##,##0.00")

Call FechaDB

End Sub

Public Sub Rotina_060_Carga_Credito()

Indice = GridNeg.Row
IndPrz = 0

If GridNeg.TextMatrix(Indice, 1) = Empty Then
   MsgBox ("Clicar somente em linha com conteúdo")
   cmdConsultar.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

neg.Open "Select * from Negociacao where chNumPedido = ('" & GridNeg.TextMatrix(Indice, 1) & "') and chNumPedidoComp = ('" & GridNeg.TextMatrix(Indice, 2) & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Erro no acesso a Tabela de Negociacao"), vbCritical
   Call FechaDB
   Exit Sub
End If

PrazoMedio = 0
DataPrazoMedio = txtDataConsulta
DataPrazoMedio = DataPrazoMedio + neg!negAPartirDe
DataConsulta = txtDataConsulta

If neg!negFaturamento > 1 Then
   For IndPrz = 1 To neg!negFaturamento
       PrazoAdicional = DataPrazoMedio - DataConsulta
       PrazoMedio = PrazoMedio + PrazoAdicional
       DataPrazoMedio = DataPrazoMedio + neg!negIntervaloFatura
   Next
Else
   PrazoMedio = neg!negAPartirDe
End If

If neg!negFaturamento = 0 Then
   If neg!negboletafrete = 0 Then
      PrazoMedio = 0
   Else
      PrazoMedio = neg!negboletafrete
   End If
Else
   PrazoMedio = PrazoMedio / neg!negFaturamento
End If

txtMediaFaturamento = PrazoMedio

ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & GridNeg.TextMatrix(Indice, 0) & "') and chNotaFiscal = ('" & GridNeg.TextMatrix(Indice, 3) & "')", db, 3, 3
If ctr.EOF Then
   ctr.Close: Set ctr = Nothing
   ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & GridNeg.TextMatrix(Indice, 0) & "') and chNotaFiscal = ('" & GridNeg.TextMatrix(Indice, 3) & "') and chFatura = ('" & 1 & "')", db, 3, 3
   If ctp.EOF Then
      MsgBox ("Não encontrei contas a receber"), vbCritical
      Call FechaDB
      Exit Sub
   End If
End If

NotaFiscal = ctr!chNotaFiscal
Indice = 0
AcumulaValor = 0

Do While Not ctr.EOF
   If ctr!chNotaFiscal = NotaFiscal And (cmbFiltro = ctr!chCodBcoLart) Then
      Indice = Indice + 1
      GridFinanc.Rows = Indice + 1
      GridFinanc.TextMatrix(Indice, 0) = ctr!chNotaFiscal & "/" & ctr!chFatura
      GridFinanc.TextMatrix(Indice, 1) = ctr!ctrDataVencito
      GridFinanc.TextMatrix(Indice, 2) = ctr!ctrDescricaoOperacao
      GridFinanc.TextMatrix(Indice, 3) = Format$(ctr!ctrvalordaboleta, "##,##0.00")
      If ctr!ctrstatus = 1 Then
         GridFinanc.TextMatrix(Indice, 4) = "Pg."
      Else
         If ctr!ctrDataVencito < Date - 1 Then
            GridFinanc.TextMatrix(Indice, 4) = "Atr."
         Else
            GridFinanc.TextMatrix(Indice, 4) = Empty
         End If
      End If
      AcumulaValor = AcumulaValor + Format$(ctr!ctrvalordaboleta, "##,##0.00")
   End If
   
ctr.MoveNext

Loop
txtValorTotalDaOperacao = Format$(AcumulaValor, "##,##0.00")

Call FechaDB

End Sub

