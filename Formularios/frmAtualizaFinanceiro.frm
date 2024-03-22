VERSION 5.00
Begin VB.Form frmAtualizaFinanceiro 
   Caption         =   "frmAtualizaFinanceiro"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbTipoDespesa 
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
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdAtualizaFinanc 
      Caption         =   "Atualizar"
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Atualiza Lançamentos de Contas Bancarias"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmAtualizaFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualizaFinanc_Click()
Dim file As Long, Linha As String
Dim AcumulaValor As Currency

On Error GoTo Erro:

file = FreeFile
Call Rotina_AbrirBanco

db.BeginTrans

rs.Open "SELECT usuEnderecoOneDrive FROM usuario WHERE chNome = '" & glbUsuario & "'", db, 3, 3

Open rs!usuEnderecoOneDrive & "Sistema\Arquivos Financeiros\" & cmbTipoDespesa & "\CONTACORRENTE110324.txt" For Input As file

rs.Close

AcumulaValor = 0

db.Execute ("INSERT INTO notafiscalentrada (chPessoa,chNotaFiscalEntrada,nfeFinalidadePagto,nfeDataEmissao,nfeDataLanc,nfeValorDaNota,nfeNF_Boleto,nfeDesbobramento,nfeTipoLancamento,nfeStatus) VALUES ('" & cmbTipoDespesa & "','" & Format$(Date, "dd-MM-yy") & "','" & 8 & "','" & Format$(Date, "yyyy-MM-dd") & "','" & Format$(Date, "yyyy-MM-dd") & "'," & AcumulaValor & ",,,,)")

Do While Not EOF(file)

    Line Input #file, Linha
    
    Nome = Mid(Linha, 44, 30)
    valor = Mid(Linha, 122, 13)
    Status = Mid(Linha, 230, 1)
    If Status = "0" Then
        rs.Open "SELECT chPessoa FROM pessoa WHERE pesRazaoSocial LIKE '" & Nome & "'", db, 3, 3
        If Not rs.EOF Then
            db.Execute ("INSERT INTO notafiscaldetprod (chPessoa,chNotaFiscalEntrada,chCodProduto,chProdutoFabrica,nfdCentroDeCusto,nfdGrupoCentroDeCusto,nfdSubGrupoCentroDeCusto,nfdQtd,nfdPU,nfdQtdParcelas,nfdValorParcela,nfdStatusPagto,nfdDataPagamento) VALUES ('" & cmbTipoDespesa & "','" & Format$(Date, "dd-MM-yy") & "','" & rs!chPessoa & "','','2','03','01'," & 1 & "," & CCur(valor) / 100 & "," & 1 & "," & CCur(valor) / 100 & "," & 1 & "," & Format$(Date, "yyyy-MM-dd") & ")")
            AcumulaValor = AcumulaValor + CCur(valor) / 100
        Else
            MsgBox ("funcionário: " & Nome & " de valor: " & CCur(valor) / 100 & " não foi encontrado!!!"), vbCritical
        End If
        rs.Close
    End If
    
    If Mid(Linha, 1, 8) = "34100015" Then
        
        If CCur(Mid(Linha, 37, 15)) / 100 = AcumulaValor Then
            
            db.Execute ("UPDATE notafiscalentrada SET nfeValorDaNota = " & AcumulaValor & " WHERE chPessoa = '" & cmbTipoDespesa & "' AND chNotaFiscalEntrada = '" & Format$(Date, "dd-MM-yy") & "'")
            
        End If
        
    End If
       
Loop

        

    db.Execute ("INSERT INTO notafiscaldesdobramento (chPessoa,chNotaFiscalEntrada,chCodProduto,chProdutoFabrica,nfdCentroDeCusto,nfdGrupoCentroDeCusto,nfdSubGrupoCentroDeCusto,nfdQtd,nfdPU,nfdQtdParcelas,nfdValorParcela,nfdStatusPagto,nfdDataPagamento) VALUES ('" & cmbTipoDespesa & "','','" & Nome & "','','2','03','01'," & 1 & "," & CCur(valor) / 100 & "," & 1 & "," & CCur(valor) / 100 & "," & 1 & "," & Date & ")")

    MsgBox AcumulaValor
    
db.CommitTrans
Exit Sub
Erro: MsgBox ("Erro ao atualizar o financeiro"), vbInformation
db.RollbackTrans
End Sub
