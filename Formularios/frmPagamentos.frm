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
         Format          =   392757249
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
         Format          =   392757249
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
Dim pessoaAnter As String
Dim Fim As Integer
Dim indice As Integer
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
   pes.Open "Select * from pessoa", db, 3, 3
   pes.MoveFirst
   Do While Not pes.EOF
      cmbClienteColaborador.AddItem pes!chPessoa
      pes.MoveNext
   Loop
Else
   ProdFornec.Open "Select * from produtofornecedor", db, 3, 3
   If Not ProdFornec.EOF Then
      ProdFornec.MoveFirst
      pessoaAnter = Empty
      Do While Not ProdFornec.EOF
         If Not pessoaAnter = ProdFornec!chTipoProduto Then
            If pes.State = 1 Then
               pes.Close: Set pes = Nothing
            End If
         
            pes.Open "Select * from pessoa where chPessoa = ('" & ProdFornec!chTipoProduto & "')", db, 3, 3
            If pes.EOF Then
               cmbClienteColaborador.AddItem ProdFornec!chTipoProduto
            End If
               
            pessoaAnter = ProdFornec!chTipoProduto
            
         End If
         ProdFornec.MoveNext
      Loop
   End If
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

pessoaAnter = Empty

Fim = 0
indice = 0

For indice = 0 To 100
    TipoDesp(indice) = Empty
Next

cmbTipopagto.Clear
cmbTipopagto.AddItem " Todos"

Call Rotina_AbrirBanco

ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not ctp.EOF Then
   indice = 1
   ctp.MoveFirst
   TipoDesp(indice) = ctp!ctpdescricaooperacao
   cmbTipopagto.AddItem ctp!ctpdescricaooperacao
   Do While Not ctp.EOF
      Do While Fim = 0
         If TipoDesp(indice) = ctp!ctpdescricaooperacao Then
            Fim = 1
         Else
            indice = indice + 1
            If TipoDesp(indice) = Empty Then
               Fim = 1
               cmbTipopagto.AddItem ctp!ctpdescricaooperacao
               TipoDesp(indice) = ctp!ctpdescricaooperacao
            End If
         End If
      Loop
   ctp.MoveNext
   Fim = 0
   indice = 1
   Loop
End If

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If

ctp.Open "Select * from historicocontaspagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not ctp.EOF Then
   indice = 1
   ctp.MoveFirst
   Do While Not ctp.EOF
      Do While Fim = 0
         If TipoDesp(indice) = ctp!ctpdescricaooperacao Then
            Fim = 1
         Else
            indice = indice + 1
            If TipoDesp(indice) = Empty Then
               Fim = 1
               cmbTipopagto.AddItem ctp!ctpdescricaooperacao
               TipoDesp(indice) = ctp!ctpdescricaooperacao
            End If
         End If
      Loop
   ctp.MoveNext
   Fim = 0
   indice = 1
   Loop
End If

cmbTipopagto.ListIndex = 0

End Sub

Public Sub CargaDespesa()

pessoaAnter = Empty

Fim = 0
indice = 0

For indice = 0 To 100
    TipoDesp(indice) = Empty
Next

cmbTipopagto.Clear
cmbTipopagto.AddItem " Todos"

Call Rotina_AbrirBanco

dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not dnfe.EOF Then
   indice = 0
   dnfe.MoveFirst
   TipoDesp(indice) = Empty
   Do While Not dnfe.EOF
      Do While Fim = 0
         If TipoDesp(indice) = dnfe!chCodProduto Then
            Fim = 1
         Else
            indice = indice + 1
            If TipoDesp(indice) = Empty Then
               Fim = 1
               cmbTipopagto.AddItem dnfe!chCodProduto
               TipoDesp(indice) = dnfe!chCodProduto
            End If
         End If
      Loop
   dnfe.MoveNext
   Fim = 0
   indice = 1
   Loop
End If

If dnfe.State = 1 Then
   dnfe.Close: Set dnfe = Nothing
End If

dnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
If Not dnfe.EOF Then
   indice = 1
   dnfe.MoveFirst
   Do While Not dnfe.EOF
      Do While Fim = 0
         If TipoDesp(indice) = dnfe!chCodProduto Then
            Fim = 1
         Else
            indice = indice + 1
            If TipoDesp(indice) = Empty Then
               Fim = 1
               cmbTipopagto.AddItem dnfe!chCodProduto
               TipoDesp(indice) = dnfe!chCodProduto
            End If
         End If
      Loop
   dnfe.MoveNext
   Fim = 0
   indice = 1
   Loop
End If

cmbTipopagto.ListIndex = 0

End Sub

Public Sub CargaGridPessoa()

Call Rotina_AbrirBanco

TotalPago = 0

indice = 1

If cmbTipopagto = " Todos" Then
   ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
Else
   ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpDescricaoOperacao = ('" & cmbTipopagto & "') and ctpStatus = ('" & 1 & "')and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
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
   ctp.Open "Select * from historicocontaspagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpStatus = ('" & 1 & "')and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
Else
   ctp.Open "Select * from historicocontaspagar where chPessoa = ('" & cmbClienteColaborador & "') and ctpDescricaoOperacao = ('" & cmbTipopagto & "') and ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
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
   GridPessoa.Rows = indice + 1
   If cmbTipopagto = " Todos" Then
      GridPessoa.TextMatrix(indice, 0) = ctp!ctpdescricaooperacao & Format$(ctp!ctpDataPagamento, "yyyymmdd")
   Else
      GridPessoa.TextMatrix(indice, 0) = Format$(ctp!ctpDataPagamento, "yyyymmdd")
   End If
   GridPessoa.TextMatrix(indice, 1) = ctp!ctpdescricaooperacao
   GridPessoa.TextMatrix(indice, 2) = ctp!chNotafiscal
   GridPessoa.TextMatrix(indice, 3) = ctp!chDataVencito
   GridPessoa.TextMatrix(indice, 4) = ctp!ctpDataPagamento
   GridPessoa.TextMatrix(indice, 5) = Format$(ctp!ctpValorDaBoleta, "###,##0.00")
   GridPessoa.TextMatrix(indice, 6) = Empty
   GridPessoa.TextMatrix(indice, 7) = ctp!ctpdescricaooperacao
   TotalPago = TotalPago + ctp!ctpValorDaBoleta
   indice = indice + 1

   ctp.MoveNext
Loop

End Sub

Public Sub CargaGridDespesa()

TotalPago = 0

indice = 1

Call Rotina_AbrirBanco

If dnfe.State = 1 Then
   dnfe.Close: Set dnfe = Nothing
End If

If cmbTipopagto = " Todos" Then
   dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
Else
   dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbClienteColaborador & "') and chCodProduto = ('" & cmbTipopagto & "')", db, 3, 3
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
      dnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & cmbClienteColaborador & "')", db, 3, 3
   Else
      dnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & cmbClienteColaborador & "') and chCodProduto = ('" & cmbTipopagto & "')", db, 3, 3
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
      ctp.Open "Select * from contas_a_pagar  where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscal = ('" & dnfe!chNotaFiscalEntrada & "') and ctpStatus = ('" & 1 & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
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
      ctp.Open "Select * from historicocontaspagar where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscal = ('" & dnfe!chNotaFiscalEntrada & "') and ctpDataPagamento > ('" & dataInicio & "') and ctpDataPagamento < ('" & dataFim & "')", db, 3, 3
      If Not ctp.EOF Then
         Encontrei = 1
         Ajustar = 1
      Else
         Encontrei = 0
      End If
   End If
   
   If Encontrei = 1 Then
      
      GridPessoa.Rows = indice + 1
      If cmbTipopagto = " Todos" Then
         GridPessoa.TextMatrix(indice, 0) = dnfe!chCodProduto & Format$(ctp!ctpDataPagamento, "yyyymmdd")
      Else
         GridPessoa.TextMatrix(indice, 0) = Format$(ctp!ctpDataPagamento, "yyyymmdd")
      End If
      GridPessoa.TextMatrix(indice, 1) = dnfe!chCodProduto
      GridPessoa.TextMatrix(indice, 2) = ctp!chNotafiscal
      GridPessoa.TextMatrix(indice, 3) = ctp!chDataVencito
      GridPessoa.TextMatrix(indice, 4) = ctp!ctpDataPagamento
      GridPessoa.TextMatrix(indice, 5) = Format$(dnfe!nfdValorParcela, "###,##0.00")
      GridPessoa.TextMatrix(indice, 6) = Empty
      GridPessoa.TextMatrix(indice, 7) = dnfe!chCodProduto
      TotalPago = TotalPago + dnfe!nfdValorParcela
      indice = indice + 1
      
   End If
   dnfe.MoveNext
Loop

End Sub

Public Sub AjustaGrid()
indice = 1
If Ajustar = 1 Then
   Do While indice < GridPessoa.Rows
      If indice = 1 Then
         AcumulaValor = GridPessoa.TextMatrix(indice, 5)
      Else
         If GridPessoa.TextMatrix(indice, 7) = GridPessoa.TextMatrix(indice - 1, 7) Then
            GridPessoa.TextMatrix(indice, 1) = Empty
            GridPessoa.TextMatrix(indice, 6) = Empty
            AcumulaValor = AcumulaValor + GridPessoa.TextMatrix(indice, 5)
         Else
            GridPessoa.TextMatrix(indice - 1, 6) = Format$(AcumulaValor, "###,##0.00")
            AcumulaValor = GridPessoa.TextMatrix(indice, 5)
         End If
      End If
      indice = indice + 1
   Loop
   GridPessoa.TextMatrix(indice - 1, 6) = Format$(AcumulaValor, "###,##0.00")
   AcumulaValor = 0
End If
Encontrei = 0
End Sub
