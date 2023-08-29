VERSION 5.00
Begin VB.Form frmSupProduto 
   Caption         =   "frmSupProduto"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14565
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
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Fornecedores"
      Height          =   2055
      Left            =   5280
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdRetira 
         Caption         =   "Ret"
         Height          =   495
         Left            =   3720
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Inc"
         Height          =   495
         Left            =   3000
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin VB.ListBox lstFornecedores 
         Height          =   960
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   4095
      End
      Begin VB.ComboBox cmbForncedores 
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
         TabIndex        =   30
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Classificação de Centro de Custo"
      Height          =   2055
      Left            =   240
      TabIndex        =   26
      Top             =   7200
      Width           =   4935
      Begin VB.ComboBox cmbGrupoCentroDeCusto 
         Height          =   420
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   3615
      End
      Begin VB.ComboBox cmbSubGrupoCentroDeCusto 
         Height          =   420
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label11 
         Caption         =   "Grupo"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Sub Grupo"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFlag 
      Height          =   495
      Left            =   7560
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   975
      Left            =   13200
      TabIndex        =   23
      Top             =   8280
      Width           =   1215
   End
   Begin VB.ListBox lstProdutos 
      Height          =   5460
      Left            =   10080
      TabIndex        =   4
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton cmbExcluir 
      Caption         =   "Excluir"
      Height          =   975
      Left            =   11640
      TabIndex        =   13
      Top             =   8280
      Width           =   1215
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   420
      Left            =   10080
      TabIndex        =   11
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   975
      Left            =   10200
      TabIndex        =   12
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox txtEspecTec 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3960
      Width           =   9615
   End
   Begin VB.TextBox txtDescricao 
      Height          =   420
      Left            =   4800
      TabIndex        =   7
      Top             =   2760
      Width           =   4935
   End
   Begin VB.TextBox txtQtdUnid 
      Height          =   420
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox cmbUnidProd 
      Height          =   420
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.ComboBox cmbClasse 
      Height          =   420
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox cmbGrupo 
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtProduto 
      Height          =   420
      Left            =   5640
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label9 
      Caption         =   "Status"
      Height          =   375
      Left            =   10080
      TabIndex        =   25
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Especificação Técnica"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Width           =   9495
   End
   Begin VB.Label lblDesc 
      Caption         =   "Descrição"
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label7 
      Caption         =   "Quantidade da Unidade"
      Height          =   615
      Left            =   2640
      TabIndex        =   20
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Unidade Produto"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2280
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
      Left            =   11760
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
      Left            =   11760
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Registro e Atualização de Produtos"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7815
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
Dim grupo As String
Dim classe As String
Dim Produto As String
Private Sub cmbClasse_LostFocus()
   lstProdutos.Clear
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select nomeProd from supproduto where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') order by nomeProd", db, 3, 3
   
   If Prod.EOF Then
   
      MsgBox ("Não há produtos cadastrados nessa categoria"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      lstProdutos.AddItem Prod!nomeProd
      Prod.MoveNext
   
   Loop
   
   txtProduto = Empty
   cmbUnidProd.ListIndex = 0
   txtQtdUnid = Empty
   txtDescricao = Empty
   txtEspecTec = Empty
   cmbGrupoCentroDeCusto.ListIndex = 0
   
   FechaDB

End Sub

Private Sub cmbExcluir_Click()
Call Rotina_AbrirBanco

On Error GoTo TE

db.Execute ("DELETE FROM supproduto WHERE grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe=('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and nomeProd=('" & txtProduto & "')")
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
   txtProduto = Empty
   cmbUnidProd.ListIndex = 0
   txtQtdUnid = Empty
   txtDescricao = Empty
   txtEspecTec = Empty
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
   Call Rotina_AbrirBanco
   Dim i As Integer
   
   rs.Open "Select * from supproduto where nomeProd=('" & txtProduto & "') and (grupo!=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') or classe!=('" & Format$(cmbClasse.ListIndex + 1, "000") & "'))", db, 3, 3
   If Not rs.EOF Then
      MsgBox "Produto com mesmo nome já existe em outro grupo ou classe", vbCritical
      FechaDB
      Exit Sub
   End If
   If flagInclusao = True Then
      pes.Open "Select MAX(codProd) as codigo from supProduto where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') AND classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "')", db, 3, 3
      
      If Not IsNull(pes!codigo) Then
         
         Dim codigoNumerico As Integer
         
         codigoNumerico = pes!codigo
         
         Produto = codigoNumerico + 1

      
      Else
      
         Produto = "00001"
      
      End If
       
      pes.Close
   
   Else
      rs.Close
      rs.Open "Select codProd from supproduto where nomeProd=('" & txtProduto & "') and grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe=('" & Format$(cmbClasse.ListIndex + 1, "000") & "')"
         Produto = rs!codProd
      rs.Close

   End If
   
   grupo = cmbGrupo.ListIndex + 1
   classe = cmbClasse.ListIndex + 1
   grupo = Format$(grupo, "00")
   classe = Format$(classe, "000")
   Produto = Format$(Produto, "00000")
   
   Prod.Open "Select * from supProduto where grupo = ('" & grupo & "') and classe = ('" & classe & "') and codProd = ('" & Produto & "')", db, 3, 3
   If Prod.EOF Then

      Prod.AddNew
   
   End If
      
   Prod!grupo = grupo
   Prod!classe = classe
   Prod!codProd = Produto
   Prod!nomeProd = txtProduto
   Prod!unidadeProd = cmbUnidProd.ListIndex
   Prod!qtdUnidade = txtQtdUnid
   Prod!Descricao = txtDescricao
   Prod!especificacaoTecnica = txtEspecTec
   Prod!Status = cmbStatus.ListIndex
   Prod!CentroDeCusto = "2"
   Prod!GrupoCentroDeCusto = Format$(cmbGrupoCentroDeCusto.ListIndex + 1, "00")
   Prod!SubGrupoCentroDeCusto = Format$(cmbSubGrupoCentroDeCusto.ListIndex + 1, "00")
   Prod.Update
   
   MsgBox "Salvo com sucesso!"
   
   If txtFlag = 1 Then
      txtFlag = 0
      Unload Me
      'frmPO.cmbGrupo.SetFocus
   End If
   
'   db.Execute ("Update supprodutofornecedor set grupo=('" & Grupo & "'),classe=('" & Classe & "'),codProd=('" & Produto & "') WHERE chTipoProduto = ('" & txtProduto & "');")
   
'   i = 0
'   Do While i < lstFornecedores.ListCount
'      rs.Open "SELECT * FROM supprodutofornecedor WHERE chPessoa = ('" & lstFornecedores.List(i) & "') and grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and codProd = ('" & Produto & "')", db, 3, 3
'
'      If rs.EOF Then
'
'         rs.AddNew
'
'      End If
'
'      rs!chPessoa = lstFornecedores.List(i)
'      rs!chTipoProduto = txtProduto
'      rs!Grupo = Format$(cmbGrupo.ListIndex + 1, "00")
'      rs!Classe = Format$(cmbClasse.ListIndex + 1, "000")
'      rs!codProd = Produto
'      rs.Update
'      rs.Close
'      i = i + 1
'   Loop
   Call limpaTela
    
   FechaDB
End Sub

Private Sub cmdIncluir_Click()
   lstFornecedores.AddItem cmbForncedores
End Sub

Private Sub cmdRetira_Click()
   Call Rotina_AbrirBanco
   
   db.Execute ("DELETE FROM supprodutofornecedor WHERE chPessoa = ('" & lstFornecedores.List(lstFornecedores.ListIndex) & "') and grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and chTipoProduto = ('" & txtProduto & "')")
   
   FechaDB
   
   lstFornecedores.RemoveItem lstFornecedores.ListIndex
   
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
   
   Prod.Open "Select * from unidadedemedida", db, 3, 3
   If Prod.EOF Then
      MsgBox "Erro: Unidades de medidas não cadastradas", vbCritical
      FechaDB
      Exit Sub
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbUnidProd.AddItem Prod!AbreviaturaUnidadeMedida
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
   
   rs.Open "SELECT chPessoa from Pessoa where pesTipoPessoa = 2 and pesStatusPessoa = 0", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não existem fornecedores cadastrados"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   rs.MoveFirst
   
   Do While Not rs.EOF
   
      cmbForncedores.AddItem rs!chPessoa
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   FechaDB
End Sub

Private Sub lstProdutos_Click()
   Call Rotina_AbrirBanco

   grupo = Format$(cmbGrupo.ListIndex + 1, "00")
   classe = Format$(cmbClasse.ListIndex + 1, "000")


   rs.Open "Select * from supProduto where nomeProd = ('" & lstProdutos & "') and grupo=('" & grupo & "') and classe = ('" & classe & "')", db, 3, 3
  
   If rs.EOF Then
   
      FechaDB
      Exit Sub
   
   End If
   Call encherTela
   
   rs.Close
   
   txtQtdUnid.SetFocus
   
   FechaDB
End Sub

Private Sub txtProduto_LostFocus()
   Call Rotina_AbrirBanco
   flagInclusao = False
   
   If txtProduto <> Empty Then
         
      rs.Open "Select * from supProduto where nomeProd = ('" & txtProduto & "')", db, 3, 3
      
      If rs.EOF Then
      
         Resp = MsgBox("Inclusão de Produto. Confirma???", vbExclamation + vbYesNo)
         If Resp = vbYes Then
         
            flagInclusao = True
            
         End If
      End If
      
      txtQtdUnid.SetFocus
      
      rs.Close
   End If
   
   FechaDB
End Sub

Public Sub encherTela()
   Dim grupoCustoInt As Integer
   Dim subGrupoCustoInt As Integer
   txtProduto = rs!nomeProd
   cmbGrupo.ListIndex = rs!grupo - 1
   cmbClasse.ListIndex = rs!classe - 1
   cmbUnidProd.ListIndex = rs!unidadeProd
   txtQtdUnid = rs!qtdUnidade
   txtDescricao = rs!Descricao
   txtEspecTec = rs!especificacaoTecnica
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
   
   lstFornecedores.Clear
   
   Prod.Open "Select chPessoa from supprodutofornecedor where grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and chTipoProduto = ('" & txtProduto & "')", db, 3, 3
   
      If Not Prod.EOF Then
      
         Do While Not Prod.EOF
         
            lstFornecedores.AddItem Prod!chPessoa
            Prod.MoveNext
         
         Loop
      
      End If
   
   Prod.Close
   cmdSalvar.SetFocus
End Sub

Public Sub limpaTela()
   txtProduto = Empty
   cmbGrupo.ListIndex = 0
   cmbClasse = Empty
   cmbUnidProd = Empty
   txtQtdUnid = Empty
   txtDescricao = Empty
   txtEspecTec = Empty
   cmbStatus = Empty
   lstProdutos.Clear
   cmbSubGrupoCentroDeCusto.Clear
   cmbGrupoCentroDeCusto.ListIndex = 0
   lstFornecedores.Clear
   cmbForncedores = Empty
End Sub

Public Sub carregadoSubGrupo()
   Dim grupodecusto As String
   
   Prod.Open "Select chGrupoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and DescricaoCentroDeCusto=('" & cmbGrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto='00'", db, 3, 3
   
   grupodecusto = Prod!chGrupoCentroDeCusto
   
   Prod.Close
   
   pes.Open "Select DescricaoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and chGrupoCentroDeCusto=('" & grupodecusto & "') and chSubGrupoCentroDeCusto>'00' ", db, 3, 3
   
   pes.MoveFirst
   
   Do While Not pes.EOF
   
      cmbSubGrupoCentroDeCusto.AddItem pes!DescricaoCentroDeCusto
      pes.MoveNext
      
   Loop
   
   
   pes.Close

End Sub
