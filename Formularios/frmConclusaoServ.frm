VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConclusaoServ 
   Caption         =   "frmConclusaoServ"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemoveDaLista 
      Caption         =   "Remove da Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12000
      TabIndex        =   25
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtNotaFiscal 
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
      Left            =   4560
      TabIndex        =   24
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cmbFornecedor 
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
      Left            =   1560
      TabIndex        =   22
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   21
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdProcessa 
      Caption         =   "Processa Conclusão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   20
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdInserirNaLista 
      Caption         =   "Inserir na Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12000
      TabIndex        =   19
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtTotalCalculado 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   9600
      TabIndex        =   17
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox txtValorTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   9600
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtQuantidade 
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
      Left            =   8160
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtPreco 
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
      Left            =   6960
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtServico 
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
      Left            =   1080
      TabIndex        =   9
      Top             =   3240
      Width           =   5895
   End
   Begin MSFlexGridLib.MSFlexGrid tblServicos 
      Height          =   2775
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "Linha|Serviços                                                  |Preço    |Quantid. |Val. Total"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtValorNotaFiscal 
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
      Height          =   480
      Left            =   8880
      TabIndex        =   7
      Top             =   2040
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker dtDataConlcusao 
      Height          =   480
      Left            =   6600
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   847
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
      Format          =   238813185
      CurrentDate     =   45268
   End
   Begin VB.ComboBox cmbPO 
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
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Fornecedor"
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
      TabIndex        =   23
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "Valor Calculado"
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
      Left            =   7560
      TabIndex        =   18
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Valor Total"
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
      Left            =   9720
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Quantidade"
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
      Left            =   8160
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Preço"
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
      Left            =   6960
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Serviços"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label5 
      Caption         =   "Valor da Nota Fiscal"
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
      Left            =   8880
      TabIndex        =   6
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Data Conclusão"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Nota Fiscal"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "PO"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Conclusão de Serviços"
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
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "frmConclusaoServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPO_LostFocus()
   Call carregaInfo
End Sub

Private Sub cmdInserirNaLista_Click()
   Call insereNaLista
   txtTotalCalculado = Format(calculaTotal, "##,##0.00")
   txtPreco = Empty
   txtQuantidade = Empty
   txtValorTotal = Empty
   txtServico = Empty
End Sub

Private Sub cmdProcessa_Click()
   Call ProcessaConlusao
End Sub

Private Sub cmdRemoveDaLista_Click()
   Call removeDaLista
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Call carregaPO
   Call carregaFornecedor
   dtDataConlcusao = Date
End Sub

Public Sub carregaPO()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM servpo WHERE status = 1", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Não existem POs cadastradas"), vbInformation
      rs.Close
      Exit Sub
   End If
   
   Do While Not rs.EOF
   
      cmbPO.AddItem rs!Id
      rs.MoveNext
      
   Loop
   
   rs.Close
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar PO: " & Err.Description), vbInformation
rs.Close
End Sub

Public Sub carregaFornecedor()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   rs.Open "Select chPessoa from pessoa where pesTipoPessoa=2", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Não existem fonecedores cadastradas"), vbInformation
      rs.Close
      Exit Sub
   End If
   
   Do While Not rs.EOF
   
      cmbFornecedor.AddItem rs!chPessoa
      rs.MoveNext
      
   Loop
   
   rs.Close
Exit Sub
Erro: MsgBox ("Erro ao carregar fonecedores: " & Err.Description), vbInformation
rs.Close
End Sub

Public Sub carregaInfo()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM servpo WHERE id = ('" & cmbPO & "')", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Não existe registro da PO selecionada"), vbCritical
      FechaDB
      Exit Sub
   End If
   
   cmbFornecedor = rs!fornecedor
   txtTotalCalculado = 0
   
   rs.Close
   
   Call carregaTabela
   
Exit Sub
Erro: MsgBox ("Erro ao carregar informações de PO: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaTabela()
   Dim i As Integer
   i = 1
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM servpodetalhe spd INNER JOIN servservico ss ON ss.grupo=spd.grupo AND ss.classe=spd.classe AND ss.codServ=spd.codServ WHERE id = ('" & cmbPO & "') and spd.status=0", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Não existe registro da PO selecionada"), vbCritical
      FechaDB
      Exit Sub
   End If
   
   tblServicos.Rows = 1
   
   Do While Not rs.EOF
      tblServicos.AddItem i & vbTab & rs!Descricao & vbTab & Format(rs!valorServ, "##,##0.00") & vbTab & Empty & vbTab & 0
      rs.MoveNext
      i = i + 1
   Loop
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar tabela: " & Err.Description), vbInformation
rs.Close
End Sub

Public Sub insereNaLista()
   On Error GoTo Erro
   
   tblServicos.TextMatrix(tblServicos.Row, 2) = txtPreco
   tblServicos.TextMatrix(tblServicos.Row, 3) = txtQuantidade
   tblServicos.TextMatrix(tblServicos.Row, 4) = txtValorTotal
   
Exit Sub
Erro: MsgBox ("Erro ao verificar linha")
End Sub

Private Sub tblServicos_Click()
   txtServico = tblServicos.TextMatrix(tblServicos.Row, 1)
   txtPreco = tblServicos.TextMatrix(tblServicos.Row, 2)
End Sub

Private Sub txtQuantidade_LostFocus()
   If txtQuantidade <> Empty And txtPreco <> Empty Then
      txtValorTotal = Format(txtQuantidade * txtPreco, "##,##0.00")
   End If
End Sub

Public Function calculaTotal() As Currency
   On Error GoTo Erro
   Dim i As Integer
   Dim total As Currency
   i = 1
   total = 0
   Do While i < tblServicos.Rows
      total = total + CCur(tblServicos.TextMatrix(i, 4))
      i = i + 1
   Loop
   calculaTotal = total
Exit Function
Erro: MsgBox ("Erro ao calcular total: " & Err.Description)
End Function

Public Sub removeDaLista()
   On Error GoTo Erro
   tblServicos.RemoveItem (tblServicos.Row)
Exit Sub
Erro: MsgBox ("Erro ao remover da lista: " & Err.Description), vbInformation
End Sub

Public Sub ProcessaConlusao()
   Dim i As Integer
   Dim diff As Integer
   diff = 0
   i = 1
   
   Call Rotina_AbrirBanco
   
   If txtNotaFiscal = Empty Then
      MsgBox ("Nota Fiscal não informado!"), vbInformation
      FechaDB
      Exit Sub
   End If
   
   If txtValorNotaFiscal = Empty Then
      MsgBox ("Valor da nota fiscal não informado!"), vbInformation
      FechaDB
      Exit Sub
   End If
   If CCur(txtValorNotaFiscal) = CCur(txtTotalCalculado) Then
      db.BeginTrans
      
      rs.Open "SELECT * FROM servpodetalhe WHERE id = '" & cmbPO & "' AND status=0", db, 3, 3
         
      Do While Not rs.EOF And i < tblServicos.Rows
         
         If rs!quantidade < rs!quantidadeAtendida + CInt(tblServicos.TextMatrix(i, 3)) Then
            MsgBox ("Erro de quantidade na linha: " & tblServicos.TextMatrix(i, 0))
            GoTo Cancela
         Else
            rs!quantidadeAtendida = rs!quantidadeAtendida + CInt(tblServicos.TextMatrix(i, 3))
            If rs!quantidade = rs!quantidadeAtendida Then
               rs!Status = 1
            End If
            rs.Update
         End If
         If i + 1 < tblServicos.Rows Then
            Do While diff < CInt(tblServicos.TextMatrix(i + 1, 0)) - CInt(tblServicos.TextMatrix(i, 0))
               rs.MoveNext
               diff = diff + 1
            Loop
         End If
         diff = 0
         i = i + 1
      Loop
   
      Call gerarfinanceiro
   
      db.CommitTrans
   End If
   
   FechaDB
   MsgBox ("Serviços foram concluídos com sucesso!"), vbInformation
   Call verificaPO
   
Exit Sub
Erro: MsgBox ("Erro ao processar conclusão de serviço: " & Err.Description), vbInformation
db.RollbackTrans
Exit Sub
Cancela: db.RollbackTrans
FechaDB
End Sub
Public Sub verificaPO()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   If Not servicosPendentes Then
      db.Execute ("UPDATE servpo SET status=2 WHERE id = '" & cmbPO & "'")
      MsgBox ("PO foi concluída com sucesso!"), vbInformation
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao verificar PO: " & Err.Description), vbInformation
FechaDB
End Sub

Public Function servicosPendentes() As Boolean
   Dim result As Boolean
   On Error GoTo Erro:
   result = True
   rs.Open "SELECT count(*) as qtd FROM servpodetalhe WHERE id = ('" & cmbPO & "')", db, 3, 3
   Prod.Open "SELECT count(*) as qtd FROM servpodetalhe WHERE id = ('" & cmbPO & "') AND status = 1", db, 3, 3
   If rs!qtd - Prod!qtd = 0 Then
      result = False
   Else
      result = True
   End If
   rs.Close
   Prod.Close
   servicosPendentes = result
Exit Function
Erro: MsgBox ("Erro ao verificar pendencia de serviços: " & Err.Description), vbInformation
rs.Close
Prod.Close
End Function

Public Sub gerarfinanceiro()
   Dim i As Integer
   
   On Error GoTo Erro
   
   If CCur(txtValorNotaFiscal) > 0 Then
   
      If rs.State = 1 Then
         rs.Close: Set rs = Nothing
      End If
   
      rs.Open "Select * from notafiscalentrada where chPessoa=('" & txtFornecedor & "') and chNotaFiscalEntrada=('" & txtNotaFiscal & "')", db, 3, 3
      
      If rs.EOF Then
      
         rs.AddNew
      
      End If
      
      rs!chPessoa = cmbFornecedor
      rs!chNotaFiscalEntrada = txtNotaFiscal
      rs!nfeFinalidadePagto = 2
      rs!nfeDataEmissao = Date
      rs!nfedataLanc = Date
      rs!nfeValorDaNota = CCur(txtValorNotaFiscal)
      rs!nfeValorFrete = 0
      rs!nfePagtoFrete = 0
      rs!nfeValorICMS = 0
      rs!nfeValorIPI = 0
      rs!nfeNF_Boleto = 3
      Prod.Open "SELECT indice from tipolancamento where chTipoDocumento = (SELECT metodoPagamento from suppedidodecompra where id=('" & txtNumPO & "'))", db, 3, 3
      If Not Prod.EOF Then
         rs!nfeTipoLancamento = Prod!indice
      End If
      Prod.Close
      rs!nfeStatus = 0
      rs.Update
      
      rs.Close
      
      i = 1
      
      Do While i < CInt(tblServicos.Rows)
         
            rs.Open "Select * from notafiscaldetprod where chPessoa=('" & cmbFornecedor & "') and chNotaFiscalEntrada=('" & txtNotaFiscal & "') and chCodProduto=('" & tblServicos.TextMatrix(i, 1) & "')", db, 3, 3
            
            If rs.EOF Then
            
               rs.AddNew
            
            End If
            rs!chPessoa = cmbFornecedor
            rs!chNotaFiscalEntrada = txtNotaFiscal
            rs!chCodProduto = tblServicos.TextMatrix(i, 1)
            'rs!chFatura = 1
            rs!nfdCentroDeCusto = "2"
            
            Prod.Open "Select GrupoCentroDeCusto,SubGrupoCentroDeCusto from servservico where descricao = ('" & tblServicos.TextMatrix(i, 1) & "')", db, 3, 3
               rs!nfdGrupoCentroDeCusto = Format$(Prod!GrupoCentroDeCusto, "00")
               rs!nfdSubGrupoCentroDeCusto = Format$(Prod!SubGrupoCentroDeCusto, "00")
               pes.Open "SELECT DescricaoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and chGrupoCentroDeCusto=('" & Format(Prod!GrupoCentroDeCusto, "00") & "') and chSubGrupoCentroDeCusto= '00' ", db, 3, 3
               rs!chProdutoFabrica = pes!DescricaoCentroDeCusto
               pes.Close
            Prod.Close
            rs!nfdQtd = tblServicos.TextMatrix(i, 3)
            rs!nfdPU = tblServicos.TextMatrix(i, 2)
            rs!nfdValorDaCompra = tblServicos.TextMatrix(i, 4)
            
            'rs!nfdValorDaParcela = NumParcelas
            rs!nfdStatusPagto = 0
            rs.Update
            i = i + 1
            rs.Close
      Loop
      
      i = 1
      
   '   Do While i < CInt(tblFaturas.Rows)
   '
   '      rs.Open "SELECT * FROM notafiscaldesdobramento WHERE chPessoa = ('" & txtFornecedor & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & Format$(tblFaturas.TextMatrix(i, 1), "yyyy-MM-dd") & "')", db, 3, 3
   '
   '      If rs.EOF Then
   '
   '         rs.AddNew
   '
   '      End If
   '
   '      rs!chPessoa = txtFornecedor
   '      rs!chNotaFiscalEntrada = txtNotaFiscal
   '      rs!chDataVencimento = Format$(tblFaturas.TextMatrix(i, 1), "yyyy-MM-dd")
   '      rs!nfdDataVencoriginal = Format$(tblFaturas.TextMatrix(i, 1), "yyyy-MM-dd")
   '      rs!nfdFaturaNumero = tblFaturas.TextMatrix(i, 0)
   '      rs!nfdValorDaFatura = tblFaturas.TextMatrix(i, 2)
   '      rs!nfdStatus = 0
   '      rs!nfdStatusPagto = 0
   '      rs!nfdOrdemBoleto = 0
   '      rs.Update
   '
   '      rs.Close
   '      i = i + 1
   '   Loop
   End If
Exit Sub
Erro: MsgBox ("Erro ao gerar financeiro: " & Err.Description), vbInformation
db.RollbackTrans
End Sub
