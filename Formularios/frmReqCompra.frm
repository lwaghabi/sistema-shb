VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReqCompra 
   Caption         =   "frmReqCompra     "
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCotacao 
      Caption         =   "Cotação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   19200
      TabIndex        =   20
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   19200
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdFiltro 
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
      Height          =   975
      Left            =   9480
      TabIndex        =   18
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   9135
      Begin VB.ListBox lstRequisicao 
         Height          =   1035
         Left            =   7320
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox lstClasse 
         Height          =   1035
         Left            =   4680
         TabIndex        =   12
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox lstGrupo 
         Height          =   1035
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox lstContrato 
         Height          =   1035
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "Requisição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Classe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
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
      Left            =   19200
      TabIndex        =   8
      Top             =   9240
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid tblAcordo 
      Height          =   1935
      Left            =   11280
      TabIndex        =   7
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "Fornecedor                                                                                                   |Preço             ||"
   End
   Begin VB.CommandButton cmdGeraPO 
      Caption         =   "Gera P.O."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   19200
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   18975
      Begin MSFlexGridLib.MSFlexGrid tblProdutos 
         Height          =   6615
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   11668
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         FormatString    =   $"frmReqCompra.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
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
      Left            =   18240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Acordo/Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18360
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Requisição de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmReqCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fornecedor As String
Private Sub cmdCotacao_Click()
   Dim i As Integer
   Dim cotacao As Integer
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT MAX(codCotacao) as novaCotacao FROM suprequisicaocompra", db, 3, 3
   If IsNull(rs!novaCotacao) Then
      cotacao = 0
   Else
      cotacao = rs!novaCotacao
   End If
   i = 1
   Do While i < tblProdutos.Rows
      If tblProdutos.TextMatrix(i, 11) = "OK" Then
         db.Execute ("UPDATE suprequisicaocompra SET codCotacao='" & cotacao + 1 & "' WHERE idRequisicao='" & tblProdutos.TextMatrix(i, 5) & "' AND nomeProd = '" & tblProdutos.TextMatrix(i, 0) & "'")
         tblProdutos.TextMatrix(i, 11) = ""
      End If
      i = i + 1
   Loop
   rs.Close
   Call geraTabela("SELECT * FROM suprequisicaocompra WHERE status = 0 ORDER BY nomeProd")
   Call FechaDB
   Call geraPDF
Exit Sub
Erro: MsgBox ("Erro ao gerar cotação: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub cmdExcluir_Click()
   
   Dim i As Integer
   
   On Error GoTo Erro:
   
   Call Rotina_AbrirBanco
   
   i = 1
   
   Do While i < tblProdutos.Rows
      
      If tblProdutos.TextMatrix(i, 11) = "OK" Then
         db.Execute ("DELETE FROM suprequisicaocompra WHERE nomeProd= '" & tblProdutos.TextMatrix(i, 0) & "' AND idRequisicao = '" & tblProdutos.TextMatrix(i, 5) & "' ")
         tblProdutos.RemoveItem (i)
         i = i - 1
      End If
      
      i = i + 1
   
   Loop
   
   FechaDB
   
   MsgBox ("Excluido com sucesso!"), vbInformation
   
Exit Sub
Erro:     MsgBox ("Erro ao excluir: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub cmdGeraPO_Click()
   Call Rotina_AbrirBanco
   Dim i As Integer
   Dim Id As Integer
   Dim query As String
   Dim acumulado As Integer
   
   On Error GoTo Erro:
   
   db.BeginTrans
   
   rs.Open "SELECT * FROM suppedidodecompra WHERE id=('" & -1 & "')", db, 3, 3
   
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!DataPedido = Date
   rs!Status = 0
   rs!formaDePagamento = Empty
   rs!metodoPagamento = Empty
   rs!moeda = Empty
   rs!fornecedor = fornecedor
   rs.Update
   
   rs.Close
      
   Prod.Open "SELECT MAX(id) as idNovo FROM suppedidodecompra", db, 3, 3
   Id = Prod!idNovo
   Prod.Close
      
   i = 1
   acumulado = tblProdutos.TextMatrix(i, 3)
   Do While i < tblProdutos.Rows - 1
      If tblProdutos.TextMatrix(i, 11) = "OK" And tblProdutos.TextMatrix(i, 0) <> tblProdutos.TextMatrix(i + 1, 0) Then
         Prod.Open "SELECT * FROM supproduto INNER JOIN unidadeembalagem ON indice=unidadeProd WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
         rs.Open "SELECT * FROM suppedidodetalhe WHERE id = ('" & Id & "') and grupo = ('" & Prod!Grupo & "') and classe = ('" & Prod!Classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
            
            If rs.EOF Then
               rs.AddNew
            End If
            
            db.Execute ("UPDATE suprequisicaocompra SET PO='" & Id & "' WHERE idRequisicao='" & tblProdutos.TextMatrix(i, 5) & "' AND nomeProd = '" & tblProdutos.TextMatrix(i, 0) & "'")
      
            rs!Id = Id
            rs!Grupo = Prod!Grupo
            rs!Classe = Prod!Classe
            rs!codProd = Prod!codProd
            rs!qtdPedida = acumulado + tblProdutos.TextMatrix(i, 4) - tblProdutos.TextMatrix(i, 2)
            If tblProdutos.TextMatrix(i, 9) <> "S/Acordo" And tblProdutos.TextMatrix(i, 9) <> Empty Then
               rs!valorUnitario = tblProdutos.TextMatrix(i, 10)
               rs!acordo = tblProdutos.TextMatrix(i, 13)
               pes.Open "SELECT * FROM suppedidodecompra WHERE id=('" & Id & "')", db, 3, 3
               pes!fornecedor = tblProdutos.TextMatrix(i, 9)
               pes.Update
               pes.Close
            End If
            rs!unidade = Prod!AbreviaturaUnidadeEmbalagem
            rs.Update
            
         rs.Close
         Prod.Close
         acumulado = tblProdutos.TextMatrix(i + 1, 3)
         Call validaLinha(i)
      ElseIf tblProdutos.TextMatrix(i, 11) = "OK" And tblProdutos.TextMatrix(i, 0) = tblProdutos.TextMatrix(i + 1, 0) Then
         acumulado = acumulado + tblProdutos.TextMatrix(i, 3)
         Call validaLinha(i)
      Else
         acumulado = tblProdutos.TextMatrix(i + 1, 3)
      End If
      i = i + 1
   Loop
   
   If tblProdutos.TextMatrix(i, 11) = "OK" Then
      Prod.Open "SELECT grupo,classe,codProd,AbreviaturaUnidadeEmbalagem FROM supproduto INNER JOIN unidadeembalagem ON indice=unidadeProd WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
      rs.Open "SELECT * FROM suppedidodetalhe WHERE id = ('" & Id & "') and grupo = ('" & Prod!Grupo & "') and classe = ('" & Prod!Classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
         
         If rs.EOF Then
            rs.AddNew
         End If
         
         db.Execute ("UPDATE suprequisicaocompra SET PO='" & Id & "' WHERE idRequisicao='" & tblProdutos.TextMatrix(i, 5) & "' AND nomeProd = '" & tblProdutos.TextMatrix(i, 0) & "'")
         
         rs!Id = Id
         rs!Grupo = Prod!Grupo
         rs!Classe = Prod!Classe
         rs!codProd = Prod!codProd
         rs!qtdPedida = acumulado + tblProdutos.TextMatrix(i, 4) - tblProdutos.TextMatrix(i, 2)
         If tblProdutos.TextMatrix(i, 9) <> "S/Acordo" And tblProdutos.TextMatrix(i, 9) <> Empty Then
            rs!valorUnitario = tblProdutos.TextMatrix(i, 10)
            rs!acordo = tblProdutos.TextMatrix(i, 13)
            pes.Open "SELECT * FROM suppedidodecompra WHERE id=('" & Id & "')", db, 3, 3
            pes!fornecedor = tblProdutos.TextMatrix(i, 9)
            pes.Update
            pes.Close
         End If
         rs!unidade = Prod!AbreviaturaUnidadeEmbalagem
         rs.Update
         
      rs.Close
      Prod.Close
      Call validaLinha(i)
   End If
   db.CommitTrans
   
   FechaDB
   
   fornecedor = Empty
   
   MsgBox ("Ordem de compra gerada com sucesso"), vbInformation
Exit Sub

Erro: MsgBox ("Erro ao gerar PO: " & Err.Description), vbCritical
db.RollbackTrans
FechaDB
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdFiltro_Click()
   Call Rotina_AbrirBanco
   
   On Error GoTo Erro:
   
   If lstContrato = "Com Contrato" Then
   
      Call geraTabela(montaQuery("SELECT src.nomeProd,src.idRequisicao,src.qtdRequisitada,se.qtdEmEstoque,src.qtdPendente,sp.estoqueMaximo,src.qtdComprar,src.status,src.codCotacao FROM suprequisicaocompra src INNER JOIN supproduto sp ON src.nomeProd = sp.nomeProd INNER JOIN supestoque se ON sp.grupo=se.grupo AND sp.classe=se.classe AND sp.codProd=se.codProd INNER JOIN supacordocomercial sac ON sac.grupo=sp.grupo AND sac.classe = sp.classe INNER JOIN supacordocomercialdetalhe sacd ON sacd.codProd = sp.codProd WHERE src.status = 0"))
      
   ElseIf lstContrato = "Sem Contrato" Then
      
      Call geraTabela(montaQuery("SELECT src.nomeProd,src.idRequisicao,src.qtdRequisitada,se.qtdEmEstoque,src.qtdPendente,sp.estoqueMaximo,src.qtdComprar,src.status,src.codCotacao FROM suprequisicaocompra src INNER JOIN supproduto sp ON src.nomeProd = sp.nomeProd INNER JOIN supestoque se ON sp.grupo=se.grupo AND sp.classe=se.classe AND sp.codProd=se.codProd WHERE (src.nomeProd,idRequisicao) NOT IN (SELECT src.nomeProd,src.idRequisicao FROM suprequisicaocompra src INNER JOIN supproduto sp ON src.nomeProd = sp.nomeProd INNER JOIN supestoque se ON sp.grupo=se.grupo AND sp.classe=se.classe AND sp.codProd=se.codProd INNER JOIN supacordocomercial sac ON sac.grupo=sp.grupo AND sac.classe = sp.classe INNER JOIN supacordocomercialdetalhe sacd ON sacd.codProd = sp.codProd AND sac.id=sacd.id WHERE src.status = 0) AND src.status = 0 "))
      
   Else
      
      Call geraTabela(montaQuery("SELECT src.nomeProd,src.idRequisicao,src.qtdRequisitada,se.qtdEmEstoque,src.qtdPendente,sp.estoqueMaximo,src.qtdComprar,src.status,src.codCotacao FROM suprequisicaocompra src INNER JOIN supproduto sp ON src.nomeProd = sp.nomeProd LEFT JOIN supestoque se ON sp.grupo=se.grupo AND sp.classe=se.classe AND sp.codProd=se.codProd WHERE src.status = 0"))
   
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao filtrar: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub Form_Load()
   Call Rotina_AbrirBanco
   Dim agregado As Integer
   Dim nomeAnterior As String
   
   If glbUsuario = "pablo" Or glbUsuario = "lwaghabi" Or glbUsuario = "raphael" Then
      cmdExcluir.Enabled = True
      cmdGeraPO.Enabled = True
   Else
      cmdExcluir.Enabled = False
      cmdGeraPO.Enabled = False
   End If
   
   txtHoje = Date
   
   lstContrato.AddItem "Geral"
   lstContrato.AddItem "Sem Contrato"
   lstContrato.AddItem "Com Contrato"
   
   Prod.Open "Select * from supgrupoclasse where classe = '000'", db, 3, 3
      If Prod.EOF Then
         MsgBox ("ERRO: Arquivo vazio."), vbCritical
         Call FechaDB
         Exit Sub
      End If
      
      Prod.MoveFirst
      
      lstGrupo.AddItem "Geral"
      
      Do While Not Prod.EOF
         lstGrupo.AddItem Prod!Descricao
         Prod.MoveNext
      Loop
   Prod.Close
   
   lstGrupo.ListIndex = 0
   
   fornecedor = Empty
   
   Call geraTabela("SELECT src.nomeProd,src.idRequisicao,src.qtdRequisitada,se.qtdEmEstoque,src.qtdPendente,sp.estoqueMaximo,src.qtdComprar,src.status,src.codCotacao FROM suprequisicaocompra src INNER JOIN supproduto sp ON src.nomeProd = sp.nomeProd LEFT JOIN supestoque se ON sp.grupo=se.grupo AND sp.classe=se.classe AND sp.codProd=se.codProd WHERE src.status = 0 ORDER BY nomeProd")
   
   FechaDB
End Sub

Private Sub lstClasse_Click()
   
   lstRequisicao.Clear
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   If lstGrupo <> "Geral" And lstClasse <> "Geral" Then
   rs.Open "SELECT DISTINCT src.idRequisicao FROM suprequisicaocompra src INNER JOIN supproduto sp ON sp.nomeProd=src.nomeProd WHERE sp.grupo = ('" & Format(lstGrupo.ListIndex, "00") & "') and sp.classe = ('" & Format(lstClasse.ListIndex, "000") & "') ORDER BY src.idRequisicao", db, 3, 3
      If rs.EOF Then
         MsgBox ("Classe não possui requisição de compra."), vbInformation
         FechaDB
         Exit Sub
      End If
      
      rs.MoveFirst
      
      lstRequisicao.AddItem "Geral"
      
      Do While Not rs.EOF
      
         lstRequisicao.AddItem rs!idRequisicao
         rs.MoveNext
      
      Loop
      
   rs.Close
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao filtrar classe: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub lstGrupo_Click()
   
   lstClasse.Clear
   
   Call Rotina_AbrirBanco
   If lstGrupo <> "Geral" Then
      Prod.Open "Select * from supgrupoclasse where grupo = ('" & Format(lstGrupo.ListIndex, "00") & "') and classe != '000'", db, 3, 3
         If Prod.EOF Then
            MsgBox ("ERRO: Arquivo vazio."), vbCritical
            Call FechaDB
            Exit Sub
         End If
         
         Prod.MoveFirst
         lstClasse.AddItem "Geral"
         
         Do While Not Prod.EOF
            lstClasse.AddItem Prod!Descricao
            Prod.MoveNext
         Loop
      Prod.Close
      lstClasse.ListIndex = 0
   End If
   FechaDB
End Sub

Private Sub tblAcordo_Click()
   If tblAcordo.TextMatrix(tblAcordo.Row, 3) < tblProdutos.TextMatrix(tblProdutos.Row, 3) And tblAcordo.TextMatrix(tblAcordo.Row, 0) <> "S/Acordo" Then
      MsgBox ("Quantidade solicitada maior do que o restante no contrato!"), vbInformation
   Else
      tblProdutos.TextMatrix(tblProdutos.Row, 9) = tblAcordo.TextMatrix(tblAcordo.Row, 0)
      tblProdutos.TextMatrix(tblProdutos.Row, 10) = tblAcordo.TextMatrix(tblAcordo.Row, 1)
      tblProdutos.TextMatrix(tblProdutos.Row, 13) = tblAcordo.TextMatrix(tblAcordo.Row, 2)
   End If
   If tblProdutos.TextMatrix(tblProdutos.Row, 11) = "OK" Then
      tblProdutos.TextMatrix(tblProdutos.Row, 11) = Empty
   End If
End Sub

Private Sub tblProdutos_Click()
   Call Rotina_AbrirBanco
   If tblProdutos.Rows > 1 Then
      If tblProdutos.Col = 1 Then
         
         rs.Open "SELECT * FROM supproduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(tblProdutos.Row, 1) & "')", db, 3, 3
         If Not rs.EOF Then
            frmEspecTec.txtEspecificacaoTecnica = rs!especificacaoTecnica
            frmEspecTec.txtDescricao = rs!nomeProd
         End If
         rs.Close
         frmEspecTec.Show
      
      Else
      
         Prod.Open "SELECT grupo,classe,codProd FROM supproduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(tblProdutos.Row, 0) & "')", db, 3, 3
         If Not Prod.EOF Then
            rs.Open "SELECT * FROM supacordocomercial INNER JOIN supacordocomercialdetalhe ON supacordocomercialdetalhe.id = supacordocomercial.id WHERE supacordocomercial.grupo=('" & Prod!Grupo & "') AND supacordocomercial.classe=('" & Prod!Classe & "') AND supacordocomercialdetalhe.codProd=('" & Prod!codProd & "')", db, 3, 3
            
            tblAcordo.Rows = 1
            
               If rs.EOF Then
               
                  tblAcordo.AddItem "S/Acordo"
                  FechaDB
                  Exit Sub
               
               End If
               
               tblAcordo.AddItem "S/Acordo"
               
               rs.MoveFirst
               
               Do While Not rs.EOF
               
                  tblAcordo.AddItem rs!fornecedor & vbTab & Format$(rs!precoUnit, "##,##0.00") & vbTab & rs!Id & vbTab & rs!qtdTotal - rs!QtdEntregue
                  rs.MoveNext
               
               Loop
            rs.Close
         End If
         Prod.Close
         
         
      
      End If
   End If
   FechaDB
End Sub

Private Sub tblProdutos_DblClick()

   If fornecedor = Empty Then
      fornecedor = tblProdutos.TextMatrix(tblProdutos.Row, 9)
   End If
   
   If tblProdutos.TextMatrix(tblProdutos.Row, 9) <> Empty And tblProdutos.TextMatrix(tblProdutos.Row, 9) <> "S/Acordo" And tblProdutos.TextMatrix(tblProdutos.Row, 9) <> fornecedor Then
   
      MsgBox ("Requisição de compra pode ser feita somente para um fornecedor por vez"), vbInformation
   
   Else
      
      If tblProdutos.TextMatrix(tblProdutos.Row, 11) = "OK" Then
         tblProdutos.TextMatrix(tblProdutos.Row, 11) = Empty
      Else
         tblProdutos.TextMatrix(tblProdutos.Row, 11) = "OK"
      End If
   
   End If
   
End Sub

Public Sub geraTabela(query As String)
   Dim agregado As Integer
   Dim nomeAnterior As String
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open query, db, 3, 3
   
      If rs.EOF Then
      
         MsgBox ("Não existem requisições pendentes " & lstContrato), vbInformation
         FechaDB
         tblProdutos.Rows = 1
         Exit Sub
      
      End If
      
      rs.MoveFirst
      tblProdutos.Rows = 1
      
      Do While Not rs.EOF
      
      Prod.Open "SELECT chPessoa FROM suprequisicao WHERE id = ('" & rs!idRequisicao & "')", db, 3, 3
         
      If rs!nomeProd = nomeAnterior Then
         agregado = agregado + rs!qtdPendente
         tblProdutos.AddItem rs!nomeProd & vbTab & "" & vbTab & rs!qtdEmEstoque & vbTab & rs!qtdPendente & vbTab & rs!estoqueMaximo & vbTab & rs!idRequisicao & vbTab & rs!qtdRequisitada & vbTab & Prod!chPessoa & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & rs!codCotacao
      
      Else
         nomeAnterior = rs!nomeProd
         If tblProdutos.Rows > 1 Then
            tblProdutos.TextMatrix(tblProdutos.Rows - 1, 8) = agregado
         End If
         agregado = rs!estoqueMaximo + rs!qtdPendente - rs!qtdEmEstoque
         tblProdutos.AddItem rs!nomeProd & vbTab & rs!nomeProd & vbTab & rs!qtdEmEstoque & vbTab & rs!qtdPendente & vbTab & rs!estoqueMaximo & vbTab & rs!idRequisicao & vbTab & rs!qtdRequisitada & vbTab & Prod!chPessoa & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & rs!codCotacao
      
      End If
      
      Prod.Close
      
      rs.MoveNext
   
      Loop
   
      tblProdutos.TextMatrix(tblProdutos.Rows - 1, 8) = agregado
   
      
      rs.Close
Exit Sub
Erro: MsgBox ("Erro ao gerar tabela: " & Err.Description), vbInformation
FechaDB
End Sub

Public Function montaQuery(query As String) As String
   
   Dim bloco1 As String
   Dim bloco2 As String
   Dim bloco3 As String
   Dim Resp As String
   
   If lstGrupo = "Geral" Or lstGrupo = Empty Then
      bloco1 = ""
   Else
      bloco1 = "AND sp.grupo = " & Format(lstGrupo.ListIndex, "00")
   End If
   
   If lstClasse = "Geral" Or lstClasse = Empty Then
      bloco2 = ""
   Else
      bloco2 = "AND sp.classe = " & Format(lstClasse.ListIndex, "000")
   End If
   
   If lstRequisicao = "Geral" Or lstRequisicao = Empty Then
      bloco3 = ""
   Else
      bloco3 = "AND src.idRequisicao = " & lstRequisicao
   End If
   
   Resp = query & " " & bloco1 & " " & bloco2 & " " & bloco3 & " " & "ORDER BY nomeProd"
   
   montaQuery = Resp
   
End Function


Public Sub validaLinha(i As Integer)
      db.Execute ("UPDATE suprequisicaocompra SET status = 1 WHERE nomeProd =  '" & tblProdutos.TextMatrix(i, 0) & "'  AND idRequisicao =  '" & tblProdutos.TextMatrix(i, 5) & "'")
End Sub

Public Sub geraPDF()
On Error GoTo Erro
Call Rotina_AbrirBanco
FechaDB
Exit Sub
Erro: MsgBox ("Erro ao gerar PDF de cotação"), vbInformation
FechaDB
End Sub
