VERSION 5.00
Begin VB.Form frmInventario 
   Caption         =   "frmInventario"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Filtrar"
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
      Left            =   8040
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cmbClasse 
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
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.ComboBox cmbGrupo 
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
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton cmdSair 
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
      Height          =   735
      Left            =   10800
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox cmbProduto 
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
      TabIndex        =   3
      Top             =   1680
      Width           =   9975
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
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
      Height          =   480
      Left            =   10320
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Classe"
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
      Left            =   3840
      TabIndex        =   11
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Grupo"
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
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Quantidade"
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
      Left            =   10320
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Produto"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   3975
   End
End
Attribute VB_Name = "frmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbGrupo_LostFocus()
   On Error GoTo Erro:
   
   Call Rotina_AbrirBanco
   
   rs.Open "Select descricao from supgrupoclasse where grupo = ('" & Format(cmbGrupo.ListIndex + 1, "00") & "') and classe > 0", db, 3, 3

   If rs.EOF Then

      MsgBox ("Não existem classes registradas")
      FechaDB
      Exit Sub
   
   End If

   rs.MoveFirst
   cmbClasse.Clear
   Do While Not rs.EOF

      cmbClasse.AddItem rs!Descricao
      rs.MoveNext

   Loop

   rs.Close
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar classes" & Err.Description), vbInformation
FechaDB
End Sub

Private Sub cmbProduto_LostFocus()
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM supestoque se INNER JOIN supproduto sp ON se.grupo=sp.grupo and se.classe=sp.classe and se.codProd=sp.codProd WHERE nomeProd = ('" & cmbProduto & "')", db, 3, 3
   If Not rs.EOF Then
      txtQuantidade = rs!qtdEmEstoque
   End If
   rs.Close
   
   FechaDB
End Sub

Private Sub cmdExcluir_Click()
   Call Rotina_AbrirBanco
   
   Prod.Open "SELECT grupo,classe,codProd FROM supproduto WHERE nomeProd = ('" & cmbProduto & "')", db, 3, 3
   db.Execute ("DELETE FROM supestoque WHERE grupo = ('" & Prod!Grupo & "') and classe = ('" & Prod!Classe & "') and codProd = ('" & Prod!codProd & "')")
   Prod.Close
   FechaDB
   
   MsgBox ("Exluido com sucesso!"), vbInformation
   
   cmbProduto = Empty
   txtQuantidade = Empty
   
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
   
   On Error GoTo Erro:
   
   Call Rotina_AbrirBanco
   
   db.BeginTrans
   
   Prod.Open "SELECT grupo,classe,codProd,pontoDePedido,estoqueMaximo FROM supproduto WHERE nomeProd = ('" & cmbProduto & "')", db, 3, 3
   
   rs.Open "SELECT * FROM supestoque WHERE grupo = ('" & Prod!Grupo & "') and classe = ('" & Prod!Classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!Grupo = Prod!Grupo
   rs!Classe = Prod!Classe
   rs!codProd = Prod!codProd
   rs!qtdEmEstoque = txtQuantidade
   rs!estoqueMinimo = Prod!pontoDePedido
   rs!estoqueMaximo = Prod!estoqueMaximo
   rs!dataUltimaAtualizacao = Date
   rs.Update
   
   rs.Close
   
   db.Execute ("UPDATE suprequisicaocompra SET qtdEmEstoque = '" & txtQuantidade & "' WHERE nomeProd = '" & cmbProduto & "'")
   
   pes.Open "SELECT SUM(qtdRequisitada) AS acumulado,qtdEmEstoque,estoqueMaximo FROM suprequisicaocompra WHERE nomeProd = ('" & cmbProduto & "') AND STATUS=0 AND idRequisicao=0", db, 3, 3
   
   db.Execute ("DELETE FROM suprequisicaocompra WHERE nomeProd = '" & cmbProduto & "' and status = 0 and idRequisicao = 0")
   
   If (pes!acumulado + pes!qtdEmEstoque) < pes!estoqueMaximo And pes!qtdEmEstoque < Prod!pontoDePedido Then
      db.Execute ("INSERT INTO suprequisicaocompra(nomeProd,idRequisicao,qtdRequisitada,qtdEmEstoque,qtdPendente,estoqueMaximo,qtdComprar,status) VALUES ('" & cmbProduto & "',0,'" & pes!estoqueMaximo - (pes!acumulado + pes!qtdEmEstoque) & "','" & pes!qtdEmEstoque & "',0,'" & pes!estoqueMaximo & "','" & pes!estoqueMaximo - (pes!acumulado + pes!qtdEmEstoque) & "',0) ")
   End If
   
   Prod.Close
   pes.Close
   
   db.CommitTrans
   
   MsgBox ("Incluido com sucesso!"), vbInformation
   
   cmbProduto = Empty
   txtQuantidade = Empty
   
Exit Sub
   
Erro: MsgBox ("Erro ao alterar inventário: " & Err.Description), vbInformation
FechaDB

End Sub

Private Sub cmdFiltrar_Click()
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   If cmbGrupo <> Empty And cmbClasse <> Empty Then
   
      rs.Open "SELECT nomeProd FROM supproduto WHERE grupo = ('" & Format(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format(cmbClasse.ListIndex + 1, "000") & "')", db, 3, 3
      If rs.EOF Then
         
         MsgBox ("Não existem produtos cadastrados desse grupo e classe"), vbInformation
         FechaDB
      
      Else
         
         cmbProduto.Clear
         rs.MoveFirst
         
         Do While Not rs.EOF
         
            cmbProduto.AddItem rs!nomeProd
            rs.MoveNext
         
         Loop
         
         cmbProduto.ListIndex = 0
      End If
      
      rs.Close
   
   End If
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao filtrar produtos" & Err.Description), vbInformation
FechaDB
cmdSair.SetFocus
End Sub

Private Sub Form_Load()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM supproduto ORDER BY nomeProd", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não há produtos cadastrados"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   rs.MoveFirst
   
   Do While Not rs.EOF
   
      cmbProduto.AddItem rs!nomeProd
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   rs.Open "Select descricao from supgrupoclasse where classe = 0", db, 3, 3

   If rs.EOF Then

      MsgBox ("Não existem grupo registrados")
      FechaDB
      Exit Sub
   
   End If

   rs.MoveFirst
   Do While Not rs.EOF

      cmbGrupo.AddItem rs!Descricao
      rs.MoveNext

   Loop

   rs.Close
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar tela: " & Err.Description), vbInformation
FechaDB
End Sub
