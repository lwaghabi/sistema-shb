VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRequisicao 
   Caption         =   "frmRequisicao"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   13230
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6720
      TabIndex        =   15
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   12000
      TabIndex        =   13
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdGeraRequisicao 
      Caption         =   "Gera Requisição"
      Height          =   855
      Left            =   12000
      TabIndex        =   12
      Top             =   4920
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid tblProdutos 
      Height          =   2415
      Left            =   480
      TabIndex        =   11
      Top             =   4440
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "Produto                                                |Qtd Solic|Qtd Atend|Saldo Req|Codigo"
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
      Left            =   6720
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
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
      TabIndex        =   8
      Top             =   3840
      Width           =   5850
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
      TabIndex        =   7
      Top             =   2880
      Width           =   3495
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
      TabIndex        =   6
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
      Height          =   2415
      Left            =   8880
      TabIndex        =   16
      Top             =   1920
      Width           =   1575
      Begin VB.CommandButton cmdVerificaEstoque 
         Caption         =   "Verifica Estoque"
         Height          =   615
         Left            =   480
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdRetiraDaLista 
         Caption         =   "Retira da lista"
         Height          =   615
         Left            =   480
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdJogaNaLista 
         Caption         =   "Jogar na lista"
         Height          =   615
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Unidade Operacional"
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
      Left            =   6720
      TabIndex        =   14
      Top             =   3240
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
      Left            =   6720
      TabIndex        =   9
      Top             =   2280
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      
         pes.Open "Select nomeProd from supproduto where grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') and classe = ('" & Format$((cmbClasse.ListIndex + 1), "000") & "') order by codProd", db, 3, 3

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

Private Sub cmbIdReq_LostFocus()
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT unidadeOperacional FROM suprequisicao WHERE id = ('" & cmbIdReq & "')", db, 3, 3
   
   If Not rs.EOF Then
   
      cmbUnidOper = rs!unidadeOperacional
   
   End If
   
   rs.Close
   
   rs.Open "SELECT * FROM supRequisicaoDetalhe inner join supRequisicao on supRequisicaoDetalhe.id = supRequisicao.id WHERE supRequisicao.id = ('" & cmbIdReq & "')", db, 3, 3
   
      If rs.EOF Then
      
         FechaDB
         Exit Sub
      
      End If
      
      tblProdutos.Rows = 1
      rs.MoveFirst
      
      Do While Not rs.EOF
         
         If rs!Status = 0 Then
            Prod.Open "SELECT nomeProd FROM supproduto WHERE grupo = ('" & rs!grupo & "') AND classe = ('" & rs!classe & "') AND codProd = ('" & rs!codProd & "')", db, 3, 3
            
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

Private Sub cmdGeraRequisicao_Click()
   Call Rotina_AbrirBanco
   
   Dim i As Integer
   
   
   db.BeginTrans
   
   rs.Open "SELECT * from supRequisicao WHERE id=('" & cmbIdReq & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!id = cmbIdReq
   rs!chPessoa = glbUsuario
   rs!maquina = glbMaquina
   rs!dataReq = Date
   'rs!justificativa = ""
   rs!unidadeOperacional = cmbUnidOper
   
   rs.Update
   
   rs.Close
   
   i = 1
   
   Call geraCodigo
   
   Do While i < tblProdutos.Rows
      
      Prod.Open "SELECT * FROM supProduto WHERE nomeProd=('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
      
      rs.Open "SELECT * FROM supRequisicaoDetalhe WHERE id=('" & cmbIdReq & "') and grupo=('" & Prod!grupo & "') and classe = ('" & Prod!classe & "') and codProd=('" & Prod!codProd & "')", db, 3, 3
      
      If rs.EOF Then
      
         rs.AddNew
      
      End If
      
      
      rs!id = cmbIdReq
      
      rs!grupo = Prod!grupo
      
      rs!classe = Prod!classe
      
      rs!codProd = Prod!codProd
      
      rs!quantidade = tblProdutos.TextMatrix(i, 1)
      
      rs!quantidadeAtendida = tblProdutos.TextMatrix(i, 2)
      
      rs!qtdEntregue = rs!qtdEntregue + tblProdutos.TextMatrix(i, 2)
      
      pes.Open "SELECT * FROM supEstoque WHERE grupo=('" & Prod!grupo & "') and classe = ('" & Prod!classe & "') and codProd=('" & Prod!codProd & "')", db, 3, 3
      
         pes!qtdReservado = pes!qtdReservado + tblProdutos.TextMatrix(i, 2)
         pes.Update
         
      pes.Close

      If CInt(tblProdutos.TextMatrix(i, 2)) > 0 Then
      
         rs!codigo = codigo
      
      End If
      
      rs.Update
      
      rs.Close
      
      Prod.Close
      
      i = i + 1
   
   Loop
   db.CommitTrans
   
   MsgBox ("Código para retirada de produtos: " & codigo), vbInformation
   
   FechaDB
End Sub

Private Sub cmdJogaNaLista_Click()
   Call Rotina_AbrirBanco
   rs.Open "SELECT chPessoa FROM supRequisicao WHERE id = ('" & cmbIdReq & "')", db, 3, 3
   
   If Not rs.EOF Then
      
      If rs!chPessoa = glbUsuario Then
         
         tblProdutos.AddItem cmbProduto & vbTab & txtQtdProd
      
      Else
      
         MsgBox ("Acesso permitido somente ao requisitante."), vbInformation

      
      End If
   
   Else
       
      tblProdutos.AddItem cmbProduto & vbTab & txtQtdProd
      
   End If
   
   rs.Close
   FechaDB
End Sub

Private Sub cmdRetiraDaLista_Click()
   Call Rotina_AbrirBanco
   rs.Open "SELECT chPessoa FROM supRequisicao WHERE id = ('" & cmbIdReq & "')", db, 3, 3
   If rs!chPessoa = glbUsuario Then
      Prod.Open "SELECT grupo,classe,codProd FROM supProduto WHERE nomeProd=('" & tblProdutos.TextMatrix(tblProdutos.Row, 0) & "')", db, 3, 3
      If CInt(tblProdutos.TextMatrix(tblProdutos.Row, 2)) > 0 Then
         db.Execute ("UPDATE supEstoque SET qtdReservado = (qtdReservado - tblProdutos.TextMatrix(tblProdutos.row,2) WHERE grupo = ('" & Prod!grupo & "') and classe = ('" & Prod!classe & "') and codProd = ('" & Prod!codProd & "'))")
         db.Close
      End If
      db.Execute ("DELETE FROM supRequisicaoDetalhe WHERE id = ('" & cmbIdReq & "') and grupo = ('" & Prod!grupo & "') and classe = ('" & Prod!classe & "') and codProd = ('" & Prod!codProd & "')")
      Prod.Close
      tblProdutos.RemoveItem (tblProdutos.Row)
   End If
   rs.Close
   FechaDB
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdVerificaEstoque_Click()
   Call Rotina_AbrirBanco
   Dim i As Integer
   
   i = 1
   
   Do While i < tblProdutos.Rows
      Prod.Open "SELECT grupo,classe,codProd FROM supProduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
      rs.Open "SELECT * FROM supEstoque WHERE grupo = ('" & Prod!grupo & "') and classe = ('" & Prod!classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
      
      
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
      
      Prod.Close
      rs.Close
      i = i + 1
   
   Loop
   
   
   FechaDB
End Sub

Private Sub Form_Load()
      
   Call Rotina_AbrirBanco
   Dim id As Integer
   
   rs.Open "SELECT id from suprequisicao", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não há requisições "), vbInformation
     
   End If

   Do While Not rs.EOF
      cmbIdReq.AddItem rs!id
      rs.MoveNext
   Loop
      
   
   rs.Close
   
   rs.Open "SELECT MAX(id) as novoId FROM suprequisicao", db, 3, 3
   
   If IsNull(rs!novoId) Then
   
      id = 1
   
   Else
       
      id = CInt(rs!novoId) + 1
   
   End If
   
   cmbIdReq = id
   
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

   rs.Open "SELECT chUnidadeOperacional FROM unidadeOperacional", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não existem unidades operacionais cadastradas"), vbInformation
      FechaDB
      Exit Sub
      
   End If
   
   rs.MoveFirst
   
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
