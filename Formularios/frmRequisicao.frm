VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRequisicao 
   Caption         =   "frmRequisicao"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17550
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   17550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelaRequisicao 
      Caption         =   "Cancela Requisição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16080
      TabIndex        =   22
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
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
      Left            =   15360
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox cmbUnidOper 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   4575
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
      Height          =   615
      Left            =   16080
      TabIndex        =   11
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdGeraRequisicao 
      Caption         =   "Gera Requisição"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16080
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid tblProdutos 
      Height          =   4575
      Left            =   480
      TabIndex        =   17
      Top             =   4440
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      FormatString    =   $"frmRequisicao.frx":0000
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
   Begin VB.TextBox txtQtdProd 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   6
      Top             =   3840
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
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   9570
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
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   4455
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
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
   End
   Begin VB.ComboBox cmbIdReq 
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
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   16080
      TabIndex        =   19
      Top             =   3720
      Width           =   1335
      Begin VB.CommandButton cmdVerificaEstoque 
         Caption         =   "Verifica Estoque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdRetiraDaLista 
         Caption         =   "Retira da lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdJogaNaLista 
         Caption         =   "Jogar na lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label8 
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
      Left            =   15360
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Unid Oper"
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
      Left            =   2160
      TabIndex        =   18
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label6 
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
      Left            =   9840
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Produto"
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
      Left            =   480
      TabIndex        =   15
      Top             =   3480
      Width           =   3735
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
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   2520
      Width           =   2295
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
      Left            =   480
      TabIndex        =   13
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Num Req."
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
      Left            =   480
      TabIndex        =   12
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Requisição ao Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmRequisicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim codigo As String

Private Sub cmbClasse_LostFocus()
   Call Rotina_AbrirBanco
      
         pes.Open "Select nomeProd from supproduto where grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') and classe = ('" & Format$((cmbClasse.ListIndex + 1), "000") & "') order by nomeProd", db, 3, 3

         If pes.EOF Then
      
            MsgBox ("Não existem produtos para essa classe")
            FechaDB
            cmdSair.SetFocus
            Exit Sub
         
         End If
      
         pes.MoveFirst
         cmbProduto.Clear
      
         Do While Not pes.EOF
      
            cmbProduto.AddItem pes!nomeProd
            pes.MoveNext
      
         Loop
      
         pes.Close
            
      FechaDB
End Sub

Private Sub cmbGrupo_LostFocus()
   Call Rotina_AbrirBanco
   
   pes.Open "Select descricao from supgrupoclasse where grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') and classe != 0", db, 3, 3

   If pes.EOF Then

      MsgBox ("Não existem classes para esse grupo")
      FechaDB
      Exit Sub
   
   End If

   pes.MoveFirst
   cmbClasse.Clear

   Do While Not pes.EOF

      cmbClasse.AddItem pes!Descricao
      pes.MoveNext

   Loop

   pes.Close
   
   FechaDB
End Sub

Private Sub cmbIdReq_Click()
   
   cmdJogaNaLista.Enabled = False
   cmdRetiraDaLista.Enabled = False
   'cmdVerificaEstoque.Enabled = False
   
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT unidadeoperacional FROM suprequisicao WHERE id = ('" & cmbIdReq & "')", db, 3, 3
   
   If Not rs.EOF Then
   
      cmbUnidOper = rs!unidadeOperacional
   
   End If
   
   rs.Close
   
   rs.Open "SELECT * FROM suprequisicaodetalhe inner join suprequisicao on suprequisicaodetalhe.id = suprequisicao.id WHERE suprequisicao.id = ('" & cmbIdReq & "')", db, 3, 3
   
      If rs.EOF Then
      
         FechaDB
         Exit Sub
      
      End If
      
      tblProdutos.Rows = 1
      rs.MoveFirst
      
      Do While Not rs.EOF
         
         If rs!Status = 0 Then
            Prod.Open "SELECT nomeProd FROM supproduto WHERE grupo = ('" & rs!Grupo & "') AND classe = ('" & rs!Classe & "') AND codProd = ('" & rs!codProd & "')", db, 3, 3
            
            If rs!chPessoa = glbUsuario Then
               
               tblProdutos.AddItem Prod!nomeProd & vbTab & rs!quantidade & vbTab & rs!quantidadeAtendida & vbTab & rs!quantidade - rs!quantidadeAtendida & vbTab & rs!codigo
            
            Else
            
               tblProdutos.AddItem Prod!nomeProd & vbTab & rs!quantidade & vbTab & rs!quantidadeAtendida & vbTab & rs!quantidade - rs!quantidadeAtendida
            
            End If
            Prod.Close
         End If
         
         rs.MoveNext
      
      Loop
   
   rs.Close
   
   FechaDB
End Sub

Private Sub cmbUnidOper_LostFocus()
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM suprequisicao WHERE id = ('" & cmbIdReq & "') ", db, 3, 3
   If Not rs.EOF Then
      
      rs!unidadeOperacional = cmbUnidOper
      rs.Update
      
   End If
   rs.Close
   FechaDB
End Sub

Private Sub cmdCancelaRequisicao_Click()
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   db.BeginTrans
   
   rs.Open "SELECT * FROM suprequisicaodetalhe WHERE id = ('" & cmbIdReq & "')", db, 3, 3
   
   If Not rs.EOF Then
   
   rs.MoveFirst
   
   Do While Not rs.EOF
   
      Prod.Open "SELECT * FROM supestoque WHERE grupo = ('" & rs!Grupo & "') and classe = ('" & rs!Classe & "') and codProd = ('" & rs!codProd & "')", db, 3, 3
      If Not Prod.EOF Then
         If (Prod!qtdReservado - rs!quantidade) >= 0 Then
            Prod!qtdReservado = Prod!qtdReservado - rs!quantidade
         Else
            Prod!qtdReservado = 0
         End If
         Prod.Update
      End If
      Prod.Close
      rs.MoveNext
   
   Loop
   
   End If
   
   rs.Close
   
   db.Execute ("DELETE FROM suprequisicao WHERE id=('" & cmbIdReq & "')")
   
   db.CommitTrans
   
   MsgBox ("Requisição cancelada com sucesso!"), vbInformation
   
   FechaDB
Exit Sub
Erro: MsgBox ("Ocorreu um erro ao cancelar a requisição : " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub cmdGeraRequisicao_Click()
   Call Rotina_AbrirBanco
   
   Dim i As Integer
   Dim flagCod As Integer
   
   
   db.BeginTrans
   
   rs.Open "SELECT * from suprequisicao WHERE id=('" & cmbIdReq & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!Id = cmbIdReq
   rs!chPessoa = glbUsuario
   rs!maquina = glbMaquina
   rs!dataReq = Date
   'rs!justificativa = ""
   rs!unidadeOperacional = cmbUnidOper
   
   rs.Update
   
   rs.Close
   
   i = 1
   
   On Error GoTo Erro:
   
   Do While i < tblProdutos.Rows
      
      Prod.Open "SELECT * FROM supproduto WHERE nomeProd=('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
      
      rs.Open "SELECT * FROM suprequisicaodetalhe WHERE id=('" & cmbIdReq & "') and grupo=('" & Prod!Grupo & "') and classe = ('" & Prod!Classe & "') and codProd=('" & Prod!codProd & "')", db, 3, 3
      
      If rs.EOF Then
      
         rs.AddNew
      
      End If
      
      
      rs!Id = cmbIdReq
      
      rs!Grupo = Prod!Grupo
      
      rs!Classe = Prod!Classe
      
      rs!codProd = Prod!codProd
      
      rs!quantidade = tblProdutos.TextMatrix(i, 1)
      
      rs!quantidadeAtendida = tblProdutos.TextMatrix(i, 2)
      
      rs!QtdEntregue = rs!QtdEntregue + tblProdutos.TextMatrix(i, 2)
      
      pes.Open "SELECT * FROM supestoque WHERE grupo=('" & Prod!Grupo & "') and classe = ('" & Prod!Classe & "') and codProd=('" & Prod!codProd & "')", db, 3, 3
      
      If Not pes.EOF Then
         
         pes!qtdReservado = pes!qtdReservado + tblProdutos.TextMatrix(i, 2)
         pes.Update
      
      End If
         

      If CInt(tblProdutos.TextMatrix(i, 2)) > 0 And tblProdutos.TextMatrix(i, 4) = Empty Then
         If flagCod = 0 Then
            Call geraCodigo
            flagCod = 1
         End If
         rs!codigo = codigo
         rs!statusEntrega = 0
      
      End If
      
      rs.Update
      
      rs.Close
      
      If tblProdutos.TextMatrix(i, 5) = 1 And tblProdutos.TextMatrix(i, 2) = 0 And Prod!Status = 1 And CInt(tblProdutos.TextMatrix(i, 3)) > 0 Then
      
         rs.Open "SELECT * FROM suprequisicaocompra WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "') and idRequisicao=('" & cmbIdReq & "')", db, 3, 3
         
            If rs.EOF Then
            
               rs.AddNew
            
            End If
         
            rs!nomeProd = tblProdutos.TextMatrix(i, 0)
            rs!idRequisicao = cmbIdReq
            rs!qtdRequisitada = tblProdutos.TextMatrix(i, 1)
            rs!qtdEmEstoque = tblProdutos.TextMatrix(i, 2)
            rs!qtdPendente = tblProdutos.TextMatrix(i, 3)
            rs!estoqueMaximo = Prod!estoqueMaximo
            rs!qtdComprar = rs!estoqueMaximo + rs!qtdPendente
            rs.Update
            
         rs.Close
      
      End If
      pes.Close
      i = i + 1
      Prod.Close
   Loop
   db.CommitTrans
   
   If flagCod = 1 Then
      MsgBox ("Código para retirada de produtos: " & codigo), vbInformation
   End If
   
   cmdGeraRequisicao.Enabled = False
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao gerar requisição ao estoque: " & Err.Description), vbInformation
End Sub

Private Sub cmdJogaNaLista_Click()
   Call Rotina_AbrirBanco
   rs.Open "SELECT chPessoa FROM suprequisicao WHERE id = ('" & cmbIdReq & "')", db, 3, 3
   
   If Not rs.EOF Then
      
      If rs!chPessoa = glbUsuario Then
         
         tblProdutos.AddItem cmbProduto & vbTab & txtQtdProd
      
      Else
      
         MsgBox ("Acesso permitido somente ao requisitante."), vbInformation

      
      End If
   
   Else
      If tblProdutos.Rows > 1 Then
         If tblProdutos.TextMatrix(tblProdutos.Row, 0) = cmbProduto Then
            tblProdutos.TextMatrix(tblProdutos.Row, 1) = txtQtdProd
         Else
            tblProdutos.AddItem cmbProduto & vbTab & txtQtdProd & vbTab & 0 & vbTab & 0
         End If
      Else
         tblProdutos.AddItem cmbProduto & vbTab & txtQtdProd & vbTab & 0 & vbTab & 0
      End If
   End If
   
   rs.Close
   FechaDB
End Sub

Private Sub cmdRetiraDaLista_Click()
   Call Rotina_AbrirBanco
   rs.Open "SELECT chPessoa FROM suprequisicao WHERE id = ('" & cmbIdReq & "')", db, 3, 3
   
   If Not rs.EOF Then
      
      If rs!chPessoa = glbUsuario Then
         
         tblProdutos.RemoveItem tblProdutos.Row
      
      Else
      
         MsgBox ("Acesso permitido somente ao requisitante."), vbInformation

      
      End If
   
   Else
      If tblProdutos.Rows = 2 Then
         tblProdutos.Rows = 1
      Else
         tblProdutos.RemoveItem tblProdutos.Row
      End If
   End If
   
   rs.Close
   FechaDB
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdVerificaEstoque_Click()
   
   If tblProdutos.Rows < 2 Then
      MsgBox ("Não há produtos na requisição, adicione produtos antes de verificar o estoque"), vbInformation
      Exit Sub
   End If
   
   Call Rotina_AbrirBanco
   Dim i As Integer
   
   i = 1
   
   Do While i < tblProdutos.Rows
      Prod.Open "SELECT grupo,classe,codProd FROM supproduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
      rs.Open "SELECT * FROM supestoque WHERE grupo = ('" & Prod!Grupo & "') and classe = ('" & Prod!Classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
      If tblProdutos.TextMatrix(i, 5) <> "1" And tblProdutos.TextMatrix(i, 4) = Empty Then
      
         If rs.EOF Then
         
            MsgBox ("Produto não existente no estoque"), vbInformation
            tblProdutos.TextMatrix(i, 2) = 0
            tblProdutos.TextMatrix(i, 3) = tblProdutos.TextMatrix(i, 1)
         
         Else
         
            If (rs!qtdEmEstoque - rs!qtdReservado) >= CInt(tblProdutos.TextMatrix(i, 1) - (tblProdutos.TextMatrix(i, 2))) Then
               
               tblProdutos.TextMatrix(i, 2) = tblProdutos.TextMatrix(i, 1) - (tblProdutos.TextMatrix(i, 2))
               tblProdutos.TextMatrix(i, 3) = 0
               
            Else
            
               tblProdutos.TextMatrix(i, 2) = (rs!qtdEmEstoque - rs!qtdReservado)
               tblProdutos.TextMatrix(i, 3) = CInt(tblProdutos.TextMatrix(i, 1) - tblProdutos.TextMatrix(i, 2))
               
            End If
            
         End If
      
         tblProdutos.TextMatrix(i, 5) = "1"
      
      End If
      
      Prod.Close
      rs.Close
      i = i + 1
   
   Loop
   
   cmdGeraRequisicao.Enabled = True
   
   FechaDB
End Sub

Private Sub Form_Load()
      
   txtHoje = Date
      
   Call Rotina_AbrirBanco
   Dim Id As Integer
   
   rs.Open "SELECT id from suprequisicao where status=0 and id>0 ", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não há requisições "), vbInformation
     
   End If

   Do While Not rs.EOF
      cmbIdReq.AddItem rs!Id
      rs.MoveNext
   Loop
      
   
   rs.Close
   
   rs.Open "SELECT MAX(id) as novoId FROM suprequisicao", db, 3, 3
   
   If IsNull(rs!novoId) Then
   
      Id = 1
   
   Else
       
      Id = CInt(rs!novoId) + 1
   
   End If
   
   cmbIdReq = Id
   
   rs.Close
   
   rs.Open "Select descricao from supgrupoclasse where classe = '000'", db, 3, 3

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

   rs.Open "SELECT chUnidadeOperacional FROM unidadeoperacional", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não existem unidades operacionais cadastradas"), vbInformation
      FechaDB
      Exit Sub
      
   End If
   
   rs.MoveFirst
   cmbUnidOper.Clear
   cmbUnidOper.AddItem "BASE"
   
   Do While Not rs.EOF
   
      cmbUnidOper.AddItem rs!chUnidadeOperacional
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   FechaDB
End Sub

Public Sub geraCodigo()
   Dim dt As Long
   Dim alfa(26) As String
   Dim result As Integer
   
   alfa(0) = "A"
   alfa(1) = "B"
   alfa(2) = "C"
   alfa(3) = "D"
   alfa(4) = "E"
   alfa(5) = "F"
   alfa(6) = "G"
   alfa(7) = "H"
   alfa(8) = "I"
   alfa(9) = "J"
   alfa(10) = "K"
   alfa(11) = "L"
   alfa(12) = "M"
   alfa(13) = "N"
   alfa(14) = "O"
   alfa(15) = "P"
   alfa(16) = "Q"
   alfa(17) = "R"
   alfa(18) = "S"
   alfa(19) = "T"
   alfa(20) = "U"
   alfa(21) = "V"
   alfa(22) = "W"
   alfa(23) = "X"
   alfa(24) = "Y"
   alfa(25) = "Z"
   
   dt = DateDiff("s", CDate("1/1/2022"), Now())
   Randomize (dt)
   result = Rnd() * 1000
   codigo = cmbIdReq & alfa(result Mod 26) & Format$(result, "000")
End Sub



Public Sub limpaCampos()
   cmbGrupo = Empty
   cmbClasse = Empty
   cmbProduto = Empty
   tblProdutos.Rows = 1
   txtQtdProd = Empty
   cmbUnidOper = Empty
End Sub

Private Sub tblProdutos_Click()
   Call Rotina_AbrirBanco
   
   cmbProduto = tblProdutos.TextMatrix(tblProdutos.Row, 0)
   txtQtdProd = tblProdutos.TextMatrix(tblProdutos.Row, 1)
   
   FechaDB
End Sub
