VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPagamentosRecebimentos 
   Caption         =   "frmPagamentosRecebimentos"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   9855
      Begin MSComCtl2.DTPicker dtFim 
         Height          =   495
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         Format          =   113770497
         CurrentDate     =   44538
      End
      Begin MSComCtl2.DTPicker dtInicio 
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62717953
         CurrentDate     =   44538
      End
      Begin VB.Label Label5 
         Caption         =   "Até"
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
         Left            =   5520
         TabIndex        =   19
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "De"
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
         Left            =   2160
         TabIndex        =   18
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.TextBox txtTotalPago 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
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
      Left            =   11005
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid GridPessoa 
      Height          =   4575
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      FormatString    =   "|Nome/Desc. Operação  |Documento     |Data Vencimento| Data Pagto     | Valor             |Sub Total            |"
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
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H000000FF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   1575
   End
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
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parâmetros de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   13455
      Begin VB.Frame Frame2 
         Caption         =   "Parâmetros de Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   10455
         Begin VB.ComboBox cmbTipopagto 
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
            Left            =   5760
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   4575
         End
         Begin VB.ComboBox cmbClienteColaborador 
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
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   5415
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo de Despesa"
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
            Left            =   5760
            TabIndex        =   14
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2535
         Begin VB.ComboBox cmbTipoConsulta 
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
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   2175
         End
      End
   End
   Begin VB.TextBox dtHoje 
      Alignment       =   2  'Center
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
      Left            =   12000
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
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
      TabIndex        =   16
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   12000
      TabIndex        =   8
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblTipoPagto 
      Caption         =   "Consulta a Pagamentos Realizados"
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
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmPagamentosRecebimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PessoaAnter As String
Dim fim As Integer
Dim Indice As Integer
Dim TipoDesp(100) As String
Dim TotalPago As Currency
Dim TotalParc As Currency
Dim AcumulaValor As Currency
Dim Encontrei As Integer
Dim EncontreiHist As Integer
Dim dataInicio As String
Dim dataFim As String
Dim AjustarGrid As Integer
Dim ChaveMes As Integer
Dim ChaveHistorico As Integer
Dim Ajustar As Integer

Private Sub cmbClienteColaborador_LostFocus()

If cmbTipoConsulta = "PESSOA" Then
   Call CargaPESSOA
Else
   Call CargaDespesa
End If
End Sub

Private Sub cmbTipoConsulta_LostFocus()

Call Rotina_AbrirBanco

cmbClienteColaborador.Clear

If cmbTipoConsulta = "PESSOA" Then
   pes.Open "Select * from Pessoa", db, 3, 3
   pes.MoveFirst
   Do While Not pes.EOF
      cmbClienteColaborador.AddItem pes!chPessoa
      pes.MoveNext
   Loop
Else
   ProdFornec.Open "Select * from ProdutoFornecedor", db, 3, 3
   ProdFornec.MoveFirst
   PessoaAnter = Empty
   Do While Not ProdFornec.EOF
      If Not PessoaAnter = ProdFornec!chTipoProduto Then
         If pes.State = 1 Then
            pes.Close: Set pes = Nothing
         End If
      
         pes.Open "Select * from Pessoa where chPessoa = ('" & ProdFornec!chTipoProduto & "')", db, 3, 3
         If pes.EOF Then
            cmbClienteColaborador.AddItem ProdFornec!chTipoProduto
         End If
            
         PessoaAnter = ProdFornec!chTipoProduto
         
      End If
      ProdFornec.MoveNext
   Loop
End If
   
End Sub

Private Sub cmdConsulta_Click()

dataInicio = dtInicio - 1
dataFim = dtFim + 1

dataInicio = Format$((dtInicio - 1), "yyyy" & "-" & "mm" & "-" & "dd")
dataFim = Format$((dtFim + 1), "yyyy" & "-" & "mm" & "-" & "dd")

If cmbTipoConsulta = "PESSOA" Then
   Call CargaGridPessoa
Else
   Call CargaGridDespesa
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
dtHoje = Date
dtInicio = Date
dtFim = Date

cmbTipoConsulta.AddItem "PESSOA"
cmbTipoConsulta.AddItem "DESPESA"

End Sub

Public Sub CargaPESSOA()

PessoaAnter = Empty

fim = 0
Indice = 0

For Indice = 0 To 100
    TipoDesp(Indice) = Empty
Next

cmbTipopagto.Clear
cmbTipopagto.AddItem " Todos"

Call Rotina_AbrirBanco

ctp.Open "Select * from Contas_A_Pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not ctp.EOF Then
   Indice = 1
   ctp.MoveFirst
   TipoDesp(Indice) = ctp!ctpdescricaooperacao
   cmbTipopagto.AddItem ctp!ctpdescricaooperacao
   Do While Not ctp.EOF
      Do While fim = 0
         If TipoDesp(Indice) = ctp!ctpdescricaooperacao Then
            fim = 1
         Else
            Indice = Indice + 1
            If TipoDesp(Indice) = Empty Then
               fim = 1
               cmbTipopagto.AddItem ctp!ctpdescricaooperacao
               TipoDesp(Indice) = ctp!ctpdescricaooperacao
            End If
         End If
      Loop
   ctp.MoveNext
   fim = 0
   Indice = 1
   Loop
End If

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If

ctp.Open "Select * from HistoricoContasPagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not ctp.EOF Then
   Indice = 1
   ctp.MoveFirst
   Do While Not ctp.EOF
      Do While fim = 0
         If TipoDesp(Indice) = ctp!ctpdescricaooperacao Then
            fim = 1
         Else
            Indice = Indice + 1
            If TipoDesp(Indice) = Empty Then
               fim = 1
               cmbTipopagto.AddItem ctp!ctpdescricaooperacao
               TipoDesp(Indice) = ctp!ctpdescricaooperacao
            End If
         End If
      Loop
   ctp.MoveNext
   fim = 0
   Indice = 1
   Loop
End If

cmbTipopagto.ListIndex = 0

End Sub

Public Sub CargaDespesa()

PessoaAnter = Empty

fim = 0
Indice = 0

For Indice = 0 To 100
    TipoDesp(Indice) = Empty
Next

cmbTipopagto.Clear
cmbTipopagto.AddItem " Todos"

Call Rotina_AbrirBanco

dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not dnfe.EOF Then
   Indice = 0
   dnfe.MoveFirst
   TipoDesp(Indice) = Empty
   Do While Not dnfe.EOF
      Do While fim = 0
         If TipoDesp(Indice) = dnfe!chCodProduto Then
            fim = 1
         Else
            Indice = Indice + 1
            If TipoDesp(Indice) = Empty Then
               fim = 1
               cmbTipopagto.AddItem dnfe!chCodProduto
               TipoDesp(Indice) = dnfe!chCodProduto
            End If
         End If
      Loop
   dnfe.MoveNext
   fim = 0
   Indice = 1
   Loop
End If

If dnfe.State = 1 Then
   dnfe.Close: Set dnfe = Nothing
End If

dnfe.Open "Select * from HistoricoNotaFiscalDetProd where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not dnfe.EOF Then
   Indice = 1
   dnfe.MoveFirst
   Do While Not dnfe.EOF
      Do While fim = 0
         If TipoDesp(Indice) = dnfe!chCodProduto Then
            fim = 1
         Else
            Indice = Indice + 1
            If TipoDesp(Indice) = Empty Then
               fim = 1
               cmbTipopagto.AddItem dnfe!chCodProduto
               TipoDesp(Indice) = dnfe!chCodProduto
            End If
         End If
      Loop
   dnfe.MoveNext
   fim = 0
   Indice = 1
   Loop
End If

cmbTipopagto.ListIndex = 0

End Sub

Public Sub CargaGridPessoa()

Call Rotina_AbrirBanco

TotalPago = 0

Indice = 1

If cmbTipopagto = " Todos" Then
   ctp.Open "Select * from Contas_A_Pagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
Else
   ctp.Open "Select * from Contas_A_Pagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpDescricaoOperacao = ('" & cmbTipopagto & "') and ctpStatus = ('" & 1 & "')and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
End If
If Not ctp.EOF Then
   ctp.MoveFirst
   Encontrei = 1
   Call EncheGridPessoa
Else
   Encontrei = 0
End If

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If

If cmbTipopagto = " Todos" Then
   ctp.Open "Select * from HistoricoContasPagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpStatus = ('" & 1 & "')and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
Else
   ctp.Open "Select * from HistoricoContasPagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpDescricaoOperacao = ('" & cmbTipopagto & "') and ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
End If

If Not ctp.EOF Then
   ctp.MoveFirst
   EncontreiHist = 1
   Call EncheGridPessoa
Else
   EncontreiHist = 0
End If

GridPessoa.Col = 0
GridPessoa.ColSel = 0
    
'GridPessoa.Row = 1

txtTotalPago = Format$(TotalPago, "###,##0.00")

GridPessoa.Sort = 8

Ajustar = 1

If Not (Encontrei = 0 And EncontreiHist = 0) Then
   Call AjustaGrid
End If

If TotalPago = 0 Then
   GridPessoa.Rows = 1
End If

Call FechaDB
End Sub

Public Sub EncheGridPessoa()
AjustarGrid = 1
Do While Not ctp.EOF
   GridPessoa.Rows = Indice + 1
   If cmbTipopagto = " Todos" Then
      GridPessoa.TextMatrix(Indice, 0) = ctp!ctpdescricaooperacao & Format$(ctp!ctpDataPagamento, "yyyymmdd")
   Else
      GridPessoa.TextMatrix(Indice, 0) = Format$(ctp!ctpDataPagamento, "yyyymmdd")
   End If
   GridPessoa.TextMatrix(Indice, 1) = ctp!ctpdescricaooperacao
   GridPessoa.TextMatrix(Indice, 2) = ctp!chNotafiscal
   GridPessoa.TextMatrix(Indice, 3) = ctp!chDataVencito
   GridPessoa.TextMatrix(Indice, 4) = ctp!ctpDataPagamento
   GridPessoa.TextMatrix(Indice, 5) = Format$(ctp!ctpValorDaBoleta, "###,##0.00")
   GridPessoa.TextMatrix(Indice, 6) = Empty
   GridPessoa.TextMatrix(Indice, 7) = ctp!ctpdescricaooperacao
   TotalPago = TotalPago + ctp!ctpValorDaBoleta
   Indice = Indice + 1

   ctp.MoveNext
Loop

End Sub

Public Sub CargaGridDespesa()

TotalPago = 0

Indice = 1

Call Rotina_AbrirBanco

If dnfe.State = 1 Then
   dnfe.Close: Set dnfe = Nothing
End If

If cmbTipopagto = " Todos" Then
   dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
Else
   dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbClienteColaborador & "') and chCodProduto = ('" & cmbTipopagto & "')", db, 3, 3
End If
If Not dnfe.EOF Then
   Encontrei = 1
   dnfe.MoveFirst
   ChaveMes = 1
   Call EncheGridDespesa
   ChaveMes = 0
Else
   Encontrei = 0
End If


If Not (Year(Date) = Year(dtInicio)) Or (Not (Month(Date) = Month(dtInicio))) Then
'If Not (Month(Date) = Month(dtInicio)) Then

   If dnfe.State = 1 Then
      dnfe.Close: Set dnfe = Nothing
   End If
   
   If cmbTipopagto = " Todos" Then
      dnfe.Open "Select * from HistoricoNotaFiscalDetProd where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
   Else
      dnfe.Open "Select * from HistoricoNotaFiscalDetProd where chPessoa = ('" & cmbClienteColaborador & "') and chCodProduto = ('" & cmbTipopagto & "')", db, 3, 3
   End If
   If Not dnfe.EOF Then
      dnfe.MoveFirst
      EncontreiHist = 1
      ChaveHistorico = 1
      Call EncheGridDespesa
      ChaveHistorico = 0
   Else
      EncontreiHist = 0
   End If
End If

GridPessoa.Col = 0
GridPessoa.ColSel = 0
    
'GridPessoa.Row = 1

txtTotalPago = Format$(TotalPago, "###,##0.00")
If cmbTipopagto = " Todos" Then
   GridPessoa.Sort = 1
Else
   GridPessoa.Sort = 8
End If

If Not (Encontrei = 0 And EncontreiHist = 0) Then
   Call AjustaGrid
End If

If TotalPago = 0 Then
   GridPessoa.Rows = 1
End If

Call FechaDB
End Sub

Public Sub EncheGridDespesa()

Encontrei = 0
Ajustar = 0

Do While Not dnfe.EOF
   
   If ctp.State = 1 Then
      ctp.Close: Set ctp = Nothing
   End If
   
   If ChaveMes = 1 Then
      ctp.Open "Select * from Contas_A_Pagar  where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscal = ('" & dnfe!chNotaFiscalEntrada & "') and ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
      If Not ctp.EOF Then
         Encontrei = 1
         Ajustar = 1
      Else
         Encontrei = 0
         
         If ctp.State = 1 Then
            ctp.Close: Set ctp = Nothing
         End If
      End If
   End If
    
   If ChaveHistorico = 1 Then
      ctp.Open "Select * from HistoricoContasPagar where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscal = ('" & dnfe!chNotaFiscalEntrada & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
      If Not ctp.EOF Then
         Encontrei = 1
         Ajustar = 1
      Else
         Encontrei = 0
      End If
   End If
   
   If Encontrei = 1 Then
      
      GridPessoa.Rows = Indice + 1
      If cmbTipopagto = " Todos" Then
         GridPessoa.TextMatrix(Indice, 0) = dnfe!chCodProduto & Format$(ctp!ctpDataPagamento, "yyyymmdd")
      Else
         GridPessoa.TextMatrix(Indice, 0) = Format$(ctp!ctpDataPagamento, "yyyymmdd")
      End If
      GridPessoa.TextMatrix(Indice, 1) = dnfe!chCodProduto
      GridPessoa.TextMatrix(Indice, 2) = ctp!chNotafiscal
      GridPessoa.TextMatrix(Indice, 3) = ctp!chDataVencito
      GridPessoa.TextMatrix(Indice, 4) = ctp!ctpDataPagamento
      GridPessoa.TextMatrix(Indice, 5) = Format$(dnfe!nfdValorParcela, "###,##0.00")
      GridPessoa.TextMatrix(Indice, 6) = Empty
      GridPessoa.TextMatrix(Indice, 7) = dnfe!chCodProduto
      TotalPago = TotalPago + dnfe!nfdValorParcela
      Indice = Indice + 1
      
   End If
   dnfe.MoveNext
Loop

End Sub

Public Sub AjustaGrid()
Indice = 1
If Ajustar = 1 Then
   Do While Indice < GridPessoa.Rows
      If Indice = 1 Then
         AcumulaValor = GridPessoa.TextMatrix(Indice, 5)
      Else
         If GridPessoa.TextMatrix(Indice, 7) = GridPessoa.TextMatrix(Indice - 1, 7) Then
            GridPessoa.TextMatrix(Indice, 1) = Empty
            GridPessoa.TextMatrix(Indice, 6) = Empty
            AcumulaValor = AcumulaValor + GridPessoa.TextMatrix(Indice, 5)
         Else
            GridPessoa.TextMatrix(Indice - 1, 6) = Format$(AcumulaValor, "###,##0.00")
            AcumulaValor = GridPessoa.TextMatrix(Indice, 5)
         End If
      End If
      Indice = Indice + 1
   Loop
   GridPessoa.TextMatrix(Indice - 1, 6) = Format$(AcumulaValor, "###,##0.00")
   AcumulaValor = 0
End If
Encontrei = 0
End Sub
