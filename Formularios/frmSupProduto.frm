VERSION 5.00
Begin VB.Form frmSupProduto 
   Caption         =   "frmSupProduto"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14820
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   14820
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEstoqueMaximo 
      Height          =   420
      Left            =   9240
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtPontoDePedido 
      Height          =   420
      Left            =   8040
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.ComboBox cmbProduto 
      Height          =   420
      Left            =   5400
      TabIndex        =   2
      Top             =   1560
      Width           =   9255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Classificação de Centro de Custo"
      Height          =   2055
      Left            =   240
      TabIndex        =   24
      Top             =   7200
      Width           =   4935
      Begin VB.ComboBox cmbGrupoCentroDeCusto 
         Height          =   420
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   3615
      End
      Begin VB.ComboBox cmbSubGrupoCentroDeCusto 
         Height          =   420
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label11 
         Caption         =   "Grupo"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Sub Grupo"
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFlag 
      Height          =   495
      Left            =   10440
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   975
      Left            =   13320
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmbExcluir 
      Caption         =   "Excluir"
      Height          =   975
      Left            =   11760
      TabIndex        =   12
      Top             =   8040
      Width           =   1215
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   420
      Left            =   5640
      TabIndex        =   10
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   975
      Left            =   10320
      TabIndex        =   11
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox txtEspecTec 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3960
      Width           =   14295
   End
   Begin VB.TextBox txtQtdUnid 
      Height          =   420
      Left            =   5520
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox cmbUnidProd 
      Height          =   420
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cmbClasse 
      Height          =   420
      Left            =   3000
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox cmbGrupo 
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label13 
      Caption         =   "Registro e Atualização de Produtos em Estoque"
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
      TabIndex        =   29
      Top             =   240
      Width           =   9855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Estoque Máximo"
      Height          =   615
      Left            =   9120
      TabIndex        =   28
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ponto de Pedido"
      Height          =   615
      Left            =   7800
      TabIndex        =   27
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Status"
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Especificação Técnica"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   3600
      Width           =   9495
   End
   Begin VB.Label Label7 
      Caption         =   "Quantidade da Unidade"
      Height          =   615
      Left            =   5520
      TabIndex        =   20
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Unidade Produto"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Classe"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Grupo"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Produto"
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label dtHoje 
      Alignment       =   2  'Center
      Caption         =   "Label3"
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
      Left            =   12840
      TabIndex        =   15
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
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
      Left            =   13080
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSupProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Resp As String
Dim flagInclusao As Boolean
Dim Grupo As String
Dim Classe As String
Dim Produto As String
Private Sub cmbClasse_LostFocus()
   
   cmbProduto.Clear
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select nomeProd from supproduto where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') order by nomeProd", db, 3, 3
   
   If Prod.EOF Then
   
      MsgBox ("Não há produtos cadastrados nessa categoria"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbProduto.AddItem Prod!nomeProd
      Prod.MoveNext
   
   Loop
   
   cmbProduto = Empty
   cmbUnidProd.ListIndex = 0
   txtQtdUnid = Empty
   txtEspecTec = Empty
   cmbGrupoCentroDeCusto.ListIndex = 0
   
   FechaDB

End Sub

Private Sub cmbExcluir_Click()
Call Rotina_AbrirBanco

On Error GoTo TE

db.Execute ("DELETE FROM supproduto WHERE grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe=('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and nomeProd=('" & cmbProduto & "')")
MsgBox ("Excluido com Sucesso!"), vbInformation
Call limpaTela

Exit Sub

TE: 'Tratamento de Exceções
    MsgBox "Verificar se há pedidos de compra ou produto em estoque antes da exclusão."

FechaDB
End Sub

Private Sub cmbGrupo_LostFocus()
   Call Rotina_AbrirBanco
   neg.Open "Select * from supgrupoclasse where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe != '000' ", db, 3, 3
   If neg.EOF Then
      MsgBox "Erro: Não existem classes nesse grupo", vbCritical
      FechaDB
      Exit Sub
   End If
   neg.MoveFirst
   
   cmbClasse.Clear
   
   Do While Not neg.EOF
      cmbClasse.AddItem neg!Descricao
      neg.MoveNext
   Loop
   
   neg.Close
   FechaDB

   cmbClasse.ListIndex = 0
   cmbProduto = Empty
   cmbUnidProd.ListIndex = 0
   txtQtdUnid = Empty
   txtEspecTec = Empty
   txtPontoDePedido = 0
   txtEstoqueMaximo = 0
   cmbGrupoCentroDeCusto.ListIndex = 0
   
End Sub

Private Sub cmbGrupoCentroDeCusto_LostFocus()
   Call Rotina_AbrirBanco
      Call carregadoSubGrupo
   FechaDB
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   db.BeginTrans
   
   Dim i As Integer
   
   rs.Open "Select * from supproduto where nomeProd=('" & cmbProduto & "') and (grupo!=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') or classe!=('" & Format$(cmbClasse.ListIndex + 1, "000") & "'))", db, 3, 3
   If Not rs.EOF Then
      MsgBox "Produto com mesmo nome já existe em outro grupo ou classe", vbCritical
      FechaDB
      Exit Sub
   End If
   If flagInclusao = True Then
      pes.Open "Select MAX(codProd) as codigo from supproduto where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') AND classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "')", db, 3, 3
      
      If Not IsNull(pes!codigo) Then
         
         Dim codigoNumerico As Integer
         
         codigoNumerico = pes!codigo
         
         Produto = codigoNumerico + 1

      
      Else
      
         Produto = "00001"
      
      End If
       
      pes.Close
      rs.Close
   Else
      rs.Close
      rs.Open "Select codProd from supproduto where nomeProd=('" & cmbProduto & "') and grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe=('" & Format$(cmbClasse.ListIndex + 1, "000") & "')"
         Produto = rs!codProd
      rs.Close

   End If
   
   
   Grupo = cmbGrupo.ListIndex + 1
   Classe = cmbClasse.ListIndex + 1
   Grupo = Format$(Grupo, "00")
   Classe = Format$(Classe, "000")
   Produto = Format$(Produto, "00000")
   
   Prod.Open "Select * from supproduto where grupo = ('" & Grupo & "') and classe = ('" & Classe & "') and codProd = ('" & Produto & "')", db, 3, 3
   
   If Prod.EOF Then

      Prod.AddNew
   
   End If
      
   Prod!Grupo = Grupo
   Prod!Classe = Classe
   Prod!codProd = Produto
   Prod!nomeProd = cmbProduto
   Prod!unidadeProd = cmbUnidProd.ListIndex
   Prod!qtdUnidade = txtQtdUnid
   Prod!pontoDePedido = txtPontoDePedido
   Prod!estoqueMaximo = txtEstoqueMaximo
   Prod!especificacaoTecnica = txtEspecTec
   Prod!Status = cmbStatus.ListIndex
   Prod!centrodecusto = "2"
   Prod!GrupoCentroDeCusto = Format$(cmbGrupoCentroDeCusto.ListIndex + 1, "00")
   Prod!SubGrupoCentroDeCusto = Format$(cmbSubGrupoCentroDeCusto.ListIndex + 1, "00")
   Prod.Update
   
   MsgBox "Salvo com sucesso!"
   
   If txtFlag = 1 Then
      txtFlag = 0
      Unload Me
   End If
   
   Call geraEstoque
   
   Prod.Close
   
   rs.Open "SELECT * FROM supestoque WHERE grupo = ('" & Grupo & "') and classe = ('" & Classe & "') and codProd = ('" & Produto & "')", db, 3, 3
   If Not rs.EOF And txtPontoDePedido <> Empty And txtEstoqueMaximo <> Empty Then
      Prod.Open "SELECT * FROM suprequisicaocompra WHERE nomeProd=('" & cmbProduto & "') and status = 0", db, 3, 3
      If Prod.EOF Then
         If Not (rs!estoqueMinimo > CInt(txtPontoDePedido)) And Not (rs!qtdEmEstoque > CInt(txtPontoDePedido)) Then
            Call geraRequisicao(txtEstoqueMaximo, rs!qtdEmEstoque)
         End If
      Else
         MsgBox ("Requisição já feita para este produto, se necessário efetue alteração na requisição de compra"), vbInformation
      End If
      rs!estoqueMinimo = txtPontoDePedido
      rs!estoqueMaximo = txtEstoqueMaximo
      rs.Update
   Else
      MsgBox ("Produto não existe no estoque!"), vbInformation
   End If
   rs.Close
   
   db.CommitTrans
   
   Call limpaTela
    
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao salvar produto: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub Form_Load()
   dtHoje = Date
   cmbGrupo = cmbGrupo
   cmbClasse = cmbClasse
   txtFlag = 0
   
   Call Rotina_AbrirBanco
   Prod.Open "Select * from supgrupoclasse where classe = '000'", db, 3, 3
         If Prod.EOF Then
            MsgBox ("ERRO: Arquivo vazio."), vbCritical
            Call FechaDB
            Exit Sub
         End If
         
         Prod.MoveFirst
         
         Do While Not Prod.EOF
            cmbGrupo.AddItem Prod!Descricao
            Prod.MoveNext
         Loop
   Prod.Close
   
   cmbGrupo.ListIndex = 0
   
   Prod.Open "Select * from unidadeembalagem", db, 3, 3
   If Prod.EOF Then
      MsgBox "Erro: Unidades de embalagens não cadastradas", vbCritical
      FechaDB
      Exit Sub
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbUnidProd.AddItem Prod!AbreviaturaUnidadeEmbalagem
      Prod.MoveNext
   
   Loop
   Prod.Close
   
   cmbStatus.AddItem "Inativo"
   cmbStatus.AddItem "Ativo"
   
   rs.Open "Select DescricaoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and chGrupoCentroDeCusto>'00' and chSubGrupoCentroDeCusto='00' ", db, 3, 3
   
   rs.MoveFirst
   
   Do While Not rs.EOF
   
      cmbGrupoCentroDeCusto.AddItem rs!DescricaoCentroDeCusto
      rs.MoveNext
      
   Loop
   
   
   rs.Close
   
   FechaDB
End Sub

Private Sub cmbProduto_LostFocus()
Dim Verifica As String

   Verifica = Empty
   Verifica = Mid$(cmbProduto, 51, 5)
   If Not Verifica = Empty Then
      MsgBox ("Código do Produto Informado ultrapassa 50 caracteres.")
      cmdSair.SetFocus
      Exit Sub
   End If
   
   Call Rotina_AbrirBanco
   flagInclusao = False
   
   If cmbProduto <> Empty Then
         
      rs.Open "Select * from supproduto where nomeProd = ('" & cmbProduto & "')", db, 3, 3
      
      If rs.EOF Then
      
         Resp = MsgBox("Inclusão de Produto. Confirma???", vbExclamation + vbYesNo)
         If Resp = vbYes Then
         
            flagInclusao = True
            
         End If
      Else
      
         Call encherTela
      
      End If
      
      rs.Close
   End If
   
   FechaDB
End Sub

Public Sub encherTela()
   Dim grupoCustoInt As Integer
   Dim subGrupoCustoInt As Integer
   cmbUnidProd.ListIndex = rs!unidadeProd
   txtQtdUnid = rs!qtdUnidade
   txtEspecTec = rs!especificacaoTecnica
   txtPontoDePedido = rs!pontoDePedido
   txtEstoqueMaximo = rs!estoqueMaximo
   cmbStatus.ListIndex = rs!Status
   If Not IsNull(rs!GrupoCentroDeCusto) Then
      grupoCustoInt = rs!GrupoCentroDeCusto
      cmbGrupoCentroDeCusto.ListIndex = grupoCustoInt - 1
   Else
      cmbGrupoCentroDeCusto = Empty
   End If
   If Not IsNull(rs!SubGrupoCentroDeCusto) Then
      subGrupoCustoInt = rs!SubGrupoCentroDeCusto
      Call carregadoSubGrupo
      cmbSubGrupoCentroDeCusto.ListIndex = subGrupoCustoInt - 1
   Else
      cmbSubGrupoCentroDeCusto = Empty
   End If
   
   cmdSalvar.SetFocus
End Sub

Public Sub limpaTela()
   cmbGrupo.ListIndex = 0
   cmbClasse = Empty
   cmbUnidProd = Empty
   txtQtdUnid = Empty
   txtEspecTec = Empty
   cmbStatus = Empty
   cmbSubGrupoCentroDeCusto.Clear
   cmbGrupoCentroDeCusto.ListIndex = 0

End Sub

Public Sub carregadoSubGrupo()
   Dim grupodecusto As String
   
   Prod.Open "Select chGrupoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and DescricaoCentroDeCusto=('" & cmbGrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto='00'", db, 3, 3
   
   grupodecusto = Prod!chGrupoCentroDeCusto
   
   Prod.Close
   
   pes.Open "Select DescricaoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and chGrupoCentroDeCusto=('" & grupodecusto & "') and chSubGrupoCentroDeCusto>'00' ", db, 3, 3
   
   pes.MoveFirst
   
   cmbSubGrupoCentroDeCusto.Clear
   
   Do While Not pes.EOF
   
      cmbSubGrupoCentroDeCusto.AddItem pes!DescricaoCentroDeCusto
      pes.MoveNext
      
   Loop
   
   
   pes.Close

End Sub

Public Sub geraRequisicao(estoqueMaximo As Integer, qtdEstoque As Integer)
   
   On Error GoTo Erro
   
   pes.Open "SELECT * FROM suprequisicaocompra WHERE nomeProd = ('" & cmbProduto & "') and idRequisicao = ('" & 0 & "')", db, 3, 3
   
   If pes.EOF Then
   
      pes.AddNew
      
   End If
   
   pes!nomeProd = cmbProduto
   pes!idRequisicao = 0
   pes!qtdRequisitada = 0
   pes!qtdEmEstoque = qtdEstoque
   pes!qtdPendente = 0
   pes!estoqueMaximo = estoqueMaximo
   pes!qtdComprar = estoqueMaximo - qtdEstoque
   pes!Status = 0
   pes.Update
   
   pes.Close
   
   MsgBox ("Requisição de compra gerada pelo sistema!"), vbInformation
Exit Sub
Erro: MsgBox ("Erro ao gerar requisição: " & Err.Description), vbInformation
pes.Close
End Sub

Public Sub geraEstoque()
   On Error GoTo Erro
   pes.Open "SELECT * FROM supestoque WHERE grupo = ('" & Grupo & "') and classe = ('" & Classe & "') and codProd = ('" & Produto & "')", db, 3, 3

   If pes.EOF Then
      
      pes.AddNew
      pes!Grupo = Grupo
      pes!Classe = Classe
      pes!codProd = Produto
      pes!qtdEmEstoque = 0
      pes!qtdReservado = 0
      pes!estoqueMinimo = txtPontoDePedido
      pes!estoqueMaximo = txtEstoqueMaximo
      pes!dataUltimaAtualizacao = Date
      pes.Update
      MsgBox ("Estoque foi gerado para novo produto!"), vbInformation

   End If
   
   pes.Close
Exit Sub
Erro:    MsgBox ("Erro ao gerar estoque: " & Err.Description), vbInformation
pes.Close
End Sub
