VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReqServCompra 
   Caption         =   "frmReqServCompra"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   18885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   17640
      TabIndex        =   12
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   17640
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdGerarPO 
      Caption         =   "Gerar PO"
      Height          =   735
      Left            =   17640
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCotacao 
      Caption         =   "Cotação"
      Height          =   735
      Left            =   17640
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid tblAcordo 
      Height          =   1935
      Left            =   12480
      TabIndex        =   8
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "Fornecedor                                                       |Valor             ||"
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9720
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9495
      Begin VB.ListBox lstRequisicao 
         Height          =   1035
         Left            =   7080
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox lstClasse 
         Height          =   1035
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox lstGrupo 
         Height          =   1035
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox lstContrato 
         Height          =   1035
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid tblServicos 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FixedCols       =   0
      FormatString    =   $"frmReqServCompra.frx":0000
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
   Begin VB.Label Label1 
      Caption         =   "Requisição de Compra de Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmReqServCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colunaConf As Integer
Dim colunaCotacao As Integer
Dim colunaGrupo As Integer
Dim colunaClasse As Integer
Dim colunaCodServ As Integer
Dim colunaId As Integer
Dim colunaAcordo As Integer
Dim colunaValAcordo As Integer
Dim colunaNumAcordo As Integer
Dim fornecedor As String

Private Sub cmdCotacao_Click()
   Call geraCotacao
End Sub
Private Sub cmdExcluir_Click()
Call excluirServicos
End Sub

Private Sub cmdFiltrar_Click()
   Call filtrarServicos
End Sub

Private Sub cmdGerarPO_Click()
Call geraPO
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call carregaTabela("SELECT ss.descricao,sr.chPessoa,src.idReq,sr.dataReq,src.acordo,src.cotacao,src.grupo,src.classe,src.codServ FROM servrequisicaocompra src INNER JOIN servservico ss ON src.grupo=ss.grupo AND src.classe=ss.classe AND src.codServ=ss.codServ INNER JOIN servrequisicao sr ON src.idReq = sr.id WHERE src.status=0")
Call carregaFiltros
colunaAcordo = 4
colunaValAcordo = 5
colunaId = 2
colunaCotacao = 6
colunaConf = 7
colunaGrupo = 8
colunaClasse = 9
colunaCodServ = 10
colunaNumAcordo = 11
End Sub

Public Sub carregaTabela(query As String)
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   rs.Open query, db, 3, 3
   
   If Not rs.EOF Then
   
      Do While Not rs.EOF
      
         tblServicos.AddItem rs!Descricao & vbTab & rs!chPessoa & vbTab & rs!idReq & vbTab & rs!dataReq & vbTab & rs!acordo & vbTab & " " & vbTab & rs!cotacao & vbTab & Empty & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!codServ
         rs.MoveNext
         
      Loop
   
   End If
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar a tabela: " & Err.Description), vbInformation
rs.Close
End Sub

Private Sub lstClasse_Click()
Call carregaRequisicao
End Sub

Private Sub lstGrupo_Click()
   Call carregaClasse
End Sub

Private Sub tblAcordo_Click()
   If tblAcordo.Row > 0 And tblAcordo.TextMatrix(tblAcordo.Row, 0) <> Empty Then
      If tblAcordo.TextMatrix(tblAcordo.Row, 3) < 1 And tblAcordo.TextMatrix(tblAcordo.Row, 0) <> "S/Acordo" Then
         MsgBox ("Quantidade solicitada maior do que o restante no contrato!"), vbInformation
      Else
         tblServicos.TextMatrix(tblServicos.Row, colunaAcordo) = tblAcordo.TextMatrix(tblAcordo.Row, 0)
         tblServicos.TextMatrix(tblServicos.Row, colunaValAcordo) = tblAcordo.TextMatrix(tblAcordo.Row, 1)
         tblServicos.TextMatrix(tblServicos.Row, colunaNumAcordo) = tblAcordo.TextMatrix(tblAcordo.Row, 2)
      End If
      If tblServicos.TextMatrix(tblServicos.Row, colunaConf) = "OK" Then
         tblServicos.TextMatrix(tblServicos.Row, colunaConf) = Empty
      End If
   End If
End Sub

Private Sub tblServicos_Click()
Call Rotina_AbrirBanco
Call carregaAcordo(tblServicos.Row)
FechaDB
End Sub

Private Sub tblServicos_DblClick()
   Call confirmaCampo
End Sub

Public Sub confirmaCampo()
   If tblServicos.Col = colunaConf And tblServicos.TextMatrix(tblServicos.Row, colunaConf) = Empty Then
      tblServicos.TextMatrix(tblServicos.Row, colunaConf) = "OK"
   ElseIf tblServicos.Col = colunaConf And tblServicos.TextMatrix(tblServicos.Row, colunaConf) = "OK" Then
      tblServicos.TextMatrix(tblServicos.Row, colunaConf) = Empty
   End If
End Sub

Public Sub geraCotacao()
   Dim i As Integer
   Dim cotacao As Integer
   cotacao = geraNumCotacao
   i = 1
   Do While i < tblServicos.Rows
      If tblServicos.TextMatrix(i, colunaConf) = "OK" Then
         tblServicos.TextMatrix(i, colunaCotacao) = cotacao
         Call atualizaCotacao(tblServicos.TextMatrix(i, colunaId), tblServicos.TextMatrix(i, colunaGrupo), tblServicos.TextMatrix(i, colunaClasse), tblServicos.TextMatrix(i, colunaCodServ), cotacao)
      End If
      i = i + 1
   Loop
End Sub

Public Function geraNumCotacao()
   Dim cotacao As Integer
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT MAX(cotacao) as novaCotacao FROM servrequisicaocompra", db, 3, 3
   If IsNull(rs!novaCotacao) Then
      cotacao = 1
   Else
      cotacao = rs!novaCotacao + 1
   End If
   
   geraNumCotacao = cotacao
   
   rs.Close
   
Exit Function
Erro: MsgBox ("Erro ao gerar número de cotação: " & Err.Description), vbInformation
rs.Close
End Function

Public Sub atualizaCotacao(Id As Integer, Grupo As String, Classe As String, codServ As String, cotacao As Integer)
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   db.BeginTrans
   db.Execute ("UPDATE servrequisicaocompra SET cotacao =" & cotacao & " WHERE grupo = " & Grupo & " AND classe = " & Classe & " AND codServ = " & codServ & " AND idReq = " & Id)
   db.CommitTrans
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao atualizar a cotação: " & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Function carregaAcordo(i As Integer)
   If tblServicos.Rows = 1 Then
      Exit Function
   End If
   rs.Open "SELECT * FROM servacordocomercial INNER JOIN servacordocomercialdetalhe ON servacordocomercialdetalhe.id = servacordocomercial.id WHERE servacordocomercial.grupo=('" & tblServicos.TextMatrix(i, colunaGrupo) & "') AND servacordocomercial.classe=('" & tblServicos.TextMatrix(i, colunaClasse) & "') AND servacordocomercialdetalhe.codServ=('" & tblServicos.TextMatrix(i, colunaCodServ) & "')", db, 3, 3
            
      tblAcordo.Rows = 1
      
         If rs.EOF Then
         
            tblAcordo.AddItem "S/Acordo"
            FechaDB
            Exit Function
         
         End If
         
         tblAcordo.AddItem "S/Acordo"
         
         rs.MoveFirst
         
         Do While Not rs.EOF
         
            tblAcordo.AddItem rs!fornecedor & vbTab & Format$(rs!precoUnit, "##,##0.00") & vbTab & rs!Id & vbTab & rs!qtdTotal - rs!QtdEntregue
            rs.MoveNext
         
         Loop
      rs.Close
End Function

Public Sub excluiRequisicao(i As Integer)
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   db.BeginTrans
   db.Execute ("DELETE FROM servrequisicaocompra WHERE grupo = " & tblServicos.TextMatrix(i, colunaGrupo) & " AND classe = " & tblServicos.TextMatrix(i, colunaClasse) & " AND codServ = " & tblServicos.TextMatrix(i, colunaCodServ) & " AND idReq = " & tblServicos.TextMatrix(i, colunaId))
   db.CommitTrans
Exit Sub
Erro: MsgBox ("Erro ao excluir requisição: " & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Sub excluirServicos()
   Dim i As Integer
   i = 1
   Do While i < tblServicos.Rows
      If tblServicos.TextMatrix(i, colunaConf) = "OK" Then
         Call excluiRequisicao(i)
         tblServicos.RemoveItem (i)
      End If
      i = i + 1
   Loop
End Sub

Public Sub carregaFiltros()
   Call carregaContrato
   Call carregaGrupo
End Sub

Public Sub carregaContrato()
   lstContrato.AddItem "Geral"
   lstContrato.AddItem "Sem Contrato"
   lstContrato.AddItem "Com Contrato"
End Sub
Public Sub carregaGrupo()
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select * from servgrupoclasse where classe = '000'", db, 3, 3
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
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar grupos" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaClasse()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   neg.Open "Select * from servgrupoclasse where grupo = ('" & Format$(lstGrupo.ListIndex, "00") & "') and classe != '000' ", db, 3, 3
   If neg.EOF Then
      MsgBox "Erro: Não existem classes nesse grupo", vbCritical
      FechaDB
      Exit Sub
   End If
   neg.MoveFirst
   
   lstClasse.Clear
   lstClasse.AddItem "Geral"
   
   Do While Not neg.EOF
      lstClasse.AddItem neg!Descricao
      neg.MoveNext
   Loop
   
   neg.Close
   
Exit Sub
Erro: MsgBox ("Erro ao carregar classes" & Err.Description), vbInformation
FechaDB
End Sub


Public Sub carregaRequisicao()
   
   lstRequisicao.Clear
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   If lstGrupo <> "Geral" And lstClasse <> "Geral" Then
   rs.Open "SELECT DISTINCT idReq FROM servrequisicaocompra WHERE grupo = ('" & Format(lstGrupo.ListIndex, "00") & "') and classe = ('" & Format(lstClasse.ListIndex, "000") & "') ORDER BY idReq", db, 3, 3
      If rs.EOF Then
         MsgBox ("Classe não possui requisição de compra."), vbInformation
         FechaDB
         Exit Sub
      End If
      
      rs.MoveFirst
      
      lstRequisicao.AddItem "Geral"
      
      Do While Not rs.EOF
      
         lstRequisicao.AddItem rs!idReq
         rs.MoveNext
      
      Loop
      
   rs.Close
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao filtrar classe: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub filtrarServicos()
   Call Rotina_AbrirBanco
   
   On Error GoTo Erro:
   
   If lstContrato = "Com Contrato" Then
   
      Call carregaTabela(montaQuery("SELECT sp.descricao,sr.chPessoa,src.idReq,sr.dataReq,src.acordo,src.cotacao,src.grupo,src.classe,src.codServ FROM suprequisicaocompra src INNER JOIN servservico ss ON ss.grupo=src.grupo AND ss.classe=src.classe AND ss.codServ=src.codServ INNER JOIN servacordocomercial sac ON sac.grupo=src.grupo AND sac.classe = src.classe INNER JOIN servacordocomercialdetalhe sacd ON sacd.codServ = src.codServ INNER JOIN servrequisicao sr ON src.idReq=sr.id WHERE src.status = 0"))
      
   ElseIf lstContrato = "Sem Contrato" Then
      
      Call carregaTabela(montaQuery("SELECT ss.descricao,sr.chPessoa,src.idReq,sr.dataReq,src.acordo,src.cotacao,src.grupo,src.classe,src.codServ FROM servrequisicaocompra src INNER JOIN servservico ss ON src.grupo=ss.grupo AND src.classe=ss.classe AND src.codServ=ss.codServ INNER JOIN servrequisicao sr ON src.idReq = sr.id WHERE (src.grupo,src.classe,src.codServ,src.idReq) NOT IN (SELECT src.grupo,src.classe,src.codServ,src.idReq FROM suprequisicaocompra src INNER JOIN servservico ss ON ss.grupo=src.grupo AND ss.classe=src.classe AND ss.codServ=src.codServ INNER JOIN servacordocomercial sac ON sac.grupo=src.grupo AND sac.classe = src.classe INNER JOIN servacordocomercialdetalhe sacd ON sacd.codServ = src.codServ INNER JOIN servrequisicao sr ON src.idReq=sr.id WHERE src.status = 0) AND src.status = 0 "))
      
   Else
      
      Call carregaTabela(montaQuery("SELECT ss.descricao,sr.chPessoa,src.idReq,sr.dataReq,src.acordo,src.cotacao,src.grupo,src.classe,src.codServ FROM servrequisicaocompra src INNER JOIN servservico ss ON src.grupo=ss.grupo AND src.classe=ss.classe AND src.codServ=ss.codServ INNER JOIN servrequisicao sr ON src.idReq = sr.id WHERE src.status=0"))
   
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao filtrar: " & Err.Description), vbInformation
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
      bloco1 = "AND src.grupo = " & Format(lstGrupo.ListIndex, "00")
   End If
   
   If lstClasse = "Geral" Or lstClasse = Empty Then
      bloco2 = ""
   Else
      bloco2 = "AND src.classe = " & Format(lstClasse.ListIndex, "000")
   End If
   
   If lstRequisicao = "Geral" Or lstRequisicao = Empty Then
      bloco3 = ""
   Else
      bloco3 = "AND src.idReq = " & lstRequisicao
   End If
   
   Resp = query & " " & bloco1 & " " & bloco2 & " " & bloco3 & " " & "ORDER BY ss.descricao"
   
   montaQuery = Resp
   
End Function


Public Sub geraPO()
   Dim i As Integer
   Dim numPO As Integer
   i = 1
   
   Call Rotina_AbrirBanco
   db.BeginTrans
   
   If verificaValidade = False Then
      MsgBox ("Fornecedores diferente para a mesma PO"), vbInformation
      db.RollbackTrans
      FechaDB
      Exit Sub
   End If
   
   If existeProdutoConf Then
   
      numPO = geraNumPO
      Call criaPO(numPO, fornecedor)
   
   End If
   
   
   Do While i < tblServicos.Rows
   
      If tblServicos.TextMatrix(i, colunaConf) = "OK" Then
         If tblServicos.TextMatrix(i, colunaAcordo) = Empty Then
            tblServicos.TextMatrix(i, colunaAcordo) = 0
            tblServicos.TextMatrix(i, colunaValAcordo) = 0
            tblServicos.TextMatrix(i, colunaNumAcordo) = 0
            
         End If
            Call criaServicoPO(numPO, tblServicos.TextMatrix(i, colunaGrupo), tblServicos.TextMatrix(i, colunaClasse), tblServicos.TextMatrix(i, colunaCodServ), CInt(tblServicos.TextMatrix(i, colunaNumAcordo)), tblServicos.TextMatrix(i, colunaValAcordo), pegaUnidadeServ(tblServicos.TextMatrix(i, colunaGrupo), tblServicos.TextMatrix(i, colunaClasse), tblServicos.TextMatrix(i, colunaCodServ)))
            Call atualizaLinha(i)
      End If
      
      i = i + 1
   Loop
   
   db.CommitTrans
   FechaDB
   MsgBox ("PO de serviços foi gerada com sucesso!"), vbInformation
Exit Sub
Erro: MsgBox ("Erro ao gerar PO: " & Err.Description), vbInformation
db.RollbackTrans
End Sub

Public Function existeProdutoConf() As Boolean
   Dim i As Integer
   Dim ret As Boolean
   ret = False
   
   Do While i < tblServicos.Rows
   
      If tblServicos.TextMatrix(i, colunaConf) = "OK" Then
         ret = True
      End If
   
      i = i + 1
   
   Loop

   existeProdutoConf = ret

End Function

Public Function geraNumPO() As Integer
   Dim numPO As Integer
   On Error GoTo Erro
   
   rs.Open "SELECT MAX(id) as codigo FROM servpo", db, 3, 3
   
   If IsNull(rs!codigo) Then
   
      numPO = 1
   
   Else
   
      numPO = rs!codigo + 1
      
   End If
   
   rs.Close
   
   geraNumPO = numPO
   
Exit Function
Erro: MsgBox ("Erro ao gerar número da PO: " & Err.Description), vbInformation
rs.Close
End Function

Public Sub criaPO(numPO As Integer, forn As String)
   On Error GoTo Erro
   db.Execute ("INSERT INTO servpo(id,dataPedido,fornecedor) VALUES ('" & numPO & "','" & Format$(Date, "yyyy-MM-dd") & "','" & forn & "')")
Exit Sub
Erro: MsgBox ("Erro ao criar PO: " & Err.Description), vbInformation
End Sub

Public Sub criaServicoPO(numPO As Integer, Grupo As String, Classe As String, codServ As String, acordo As Integer, valorAcordo As String, unidade As Integer)
   On Error GoTo Erro
   db.Execute ("INSERT INTO servpodetalhe (id,grupo,classe,codServ,acordo,valorServ,unid) VALUES ('" & numPO & "', '" & Grupo & "' , '" & Classe & "', '" & codServ & "','" & acordo & "','" & Replace(valorAcordo, ",", ".") & "','" & unidade & "')")
Exit Sub
Erro: MsgBox ("Erro ao criar serviço da PO: " & Err.Description), vbInformation
End Sub

Public Function pegaUnidadeServ(Grupo As String, Classe As String, codServ As String) As Integer
   Dim unid As Integer
   On Error GoTo Erro
   rs.Open "SELECT * FROM servservico WHERE grupo = ('" & Grupo & "') AND classe = ('" & Classe & "') AND codServ = ('" & codServ & "')", db, 3, 3
   unid = rs!unidade
   rs.Close
   pegaUnidadeServ = unid
Exit Function
Erro: MsgBox ("Erro ao acessar unidade de serviço: " & Err.Description), vbInformation
rs.Close
End Function

Public Sub atualizaLinha(i As Integer)
   On Error GoTo Erro:
   db.Execute ("UPDATE servrequisicaocompra SET status = 1 WHERE idReq = '" & tblServicos.TextMatrix(i, colunaId) & "' AND grupo = '" & tblServicos.TextMatrix(i, colunaGrupo) & "' AND classe = '" & tblServicos.TextMatrix(i, colunaClasse) & "' AND codServ = '" & tblServicos.TextMatrix(i, colunaCodServ) & "'")
Exit Sub
Erro: MsgBox ("Erro ao validar linha: " & Err.Description), vbInformation
FechaDB
End Sub

Public Function verificaValidade() As Boolean
   On Error GoTo Erro
   Dim i As Integer
   Dim j As Integer
   Dim listaAcordo(100) As Integer
   Dim ponteiro As String
   Dim resultado As Boolean
      
   resultado = True
      
   i = 1
   j = 0
   Do While i < tblServicos.Rows
      If tblServicos.TextMatrix(i, colunaAcordo) <> Empty And tblServicos.TextMatrix(i, colunaConf) = "OK" Then
         listaAcordo(j) = tblServicos.TextMatrix(i, colunaNumAcordo)
         j = j + 1
      End If
      i = i + 1
   Loop
   
   i = 0
   
   If j > 0 Then
   
      rs.Open "SELECT fornecedor FROM servacordocomercial WHERE id = '" & listaAcordo(0) & "'", db, 3, 3
      ponteiro = rs!fornecedor
      fornecedor = ponteiro
      rs.Close
   
      Do While i < j
         rs.Open "SELECT fornecedor FROM servacordocomercial WHERE id = '" & listaAcordo(i) & "'", db, 3, 3
         If ponteiro <> rs!fornecedor Then
            resultado = False
            i = j + 1
         Else
            i = i + 1
         End If
         rs.Close
      Loop
   End If
   
   verificaValidade = resultado
Exit Function
Erro: MsgBox ("Erro ao verificar validade das compras: " & Err.Description), vbInformation
End Function
