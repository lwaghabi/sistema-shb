VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReqServ 
   Caption         =   "frmReqServ"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   8
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemoveDaLista 
      Caption         =   "Retirar da Lista"
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
      Left            =   9360
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdInserirNaLista 
      Caption         =   "Inserir na Lista"
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
      Left            =   9360
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelaReq 
      Caption         =   "Cancelar Requisiçao"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   7
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdGerarReq 
      Caption         =   "Gerar Requisição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10680
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid tblServicos 
      Height          =   3495
      Left            =   240
      TabIndex        =   14
      Top             =   5280
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "Serviços                                                                                            |||"
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
   Begin VB.ComboBox cmbReq 
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
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbServ 
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
      Top             =   4320
      Width           =   8775
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
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   3615
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
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "ID"
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
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Requisição de Serviços"
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
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmReqServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbClasse_LostFocus()
   Call carregaServico
End Sub

Private Sub cmbGrupo_LostFocus()
   Call carregaClasse
End Sub

Private Sub cmbReq_LostFocus()
   Call carregaTela
End Sub

Private Sub cmdGerarReq_Click()
   Call salvaRequisicao
End Sub

Private Sub cmdInserirNaLista_Click()
   Call adicionaNaLista
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdRemoveDaLista_Click()
   Call removeDaLista
End Sub

Private Sub cmdCancelaReq_Click()
   Call cancelaReq
End Sub

Private Sub Form_Load()
   Call carregaGrupo
   Call carregaReq
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
      
      Do While Not Prod.EOF
         cmbGrupo.AddItem Prod!Descricao
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
   neg.Open "Select * from servgrupoclasse where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe != '000' ", db, 3, 3
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
   
Exit Sub
Erro: MsgBox ("Erro ao carregar classes" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaServico()
   cmbServ.Clear
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select descricao from servservico where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') order by codServ", db, 3, 3
   
   If Prod.EOF Then
   
      MsgBox ("Não há serviços cadastrados nessa categoria"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbServ.AddItem Prod!Descricao
      Prod.MoveNext
   
   Loop
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar serviços: " & Err.Description), vbInformation
FechaDB
End Sub


Public Sub carregaReq()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM servrequisicao where status = 0", db, 3, 3
      
      If rs.EOF Then
      
         cmbReq.AddItem 1
         cmbReq.ListIndex = 0
      
      Else
      
      rs.MoveFirst
      
         Do While Not rs.EOF
         
            cmbReq.AddItem rs!Id
            rs.MoveNext
         
         Loop
         
         pes.Open "SELECT MAX(id) as id FROM servrequisicao", db, 3, 3
         
         cmbReq.AddItem pes!Id + 1
         
         pes.Close
         
         cmbReq.ListIndex = cmbReq.ListCount - 1
         
      End If
      
      rs.Close
      
Exit Sub
Erro: MsgBox ("Erro ao carregar requisições: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub adicionaNaLista()
   If Not estaNaTabela(cmbServ) Then
      tblServicos.AddItem cmbServ & vbTab & Format(cmbGrupo.ListIndex + 1, "00") & vbTab & Format(cmbClasse.ListIndex + 1, "000") & vbTab & Format(cmbServ.ListIndex + 1, "00000")
   Else
      MsgBox ("Serviço já está listado na requisição"), vbInformation
   End If
End Sub

Public Sub removeDaLista()
   If tblServicos.Row <> 0 Then
      tblServicos.RemoveItem (tblServicos.Row)
   End If
End Sub

Public Function estaNaTabela(Item As String) As Boolean
   Dim i As Integer
   
   i = 1
   
   Do While i < tblServicos.Rows
      
      If tblServicos.TextMatrix(i, 0) = Item Then
         estaNaTabela = True
         Exit Function
      End If
      
      i = i + 1
      
   Loop
   
   estaNaTabela = False
   
End Function

Public Sub salvaRequisicao()
   On Error GoTo Erro
   Dim i As Integer
   Call Rotina_AbrirBanco
   
   If cmbReq <> Empty And tblServicos.Rows > 1 Then
      
      db.BeginTrans
   
      Call salvaReq
      
      i = 1
      
      Do While i < tblServicos.Rows
         Call salvaDetalhe(i)
         Call criaReqCompraServ(i)
         i = i + 1
      Loop
      
      db.CommitTrans
      FechaDB
      MsgBox ("Salvo com sucesso!"), vbInformation
   Else
      MsgBox ("Informação incompleta para gerar requisição!"), vbInformation
   End If
Exit Sub
Erro: MsgBox ("Erro ao salvar requisição: " & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Sub salvaReq()
   rs.Open "SELECT * FROM servrequisicao WHERE id = ('" & cmbReq & "')", db, 3, 3
   
   If rs.EOF Then
      
      rs.AddNew
      
   End If
   
   rs!Id = cmbReq
   rs!chPessoa = glbUsuario
   rs!maquina = glbMaquina
   rs!dataReq = Date
   
   rs.Update
   rs.Close
End Sub

Public Sub salvaDetalhe(indice As Integer)
   rs.Open "SELECT * FROM servrequisicaodetalhe WHERE id = ('" & cmbReq & "') AND grupo = ('" & tblServicos.TextMatrix(indice, 1) & "') AND classe = ('" & tblServicos.TextMatrix(indice, 2) & "') AND codServ = ('" & tblServicos.TextMatrix(indice, 3) & "')", db, 3, 3
   
   If rs.EOF Then
      
      rs.AddNew
      
   End If
   
   rs!Id = cmbReq
   rs!Grupo = tblServicos.TextMatrix(indice, 1)
   rs!Classe = tblServicos.TextMatrix(indice, 2)
   rs!codServ = tblServicos.TextMatrix(indice, 3)
   
   Call RegistraMovSv(rs!Grupo, rs!Classe, rs!codServ)
   
   rs.Update
   rs.Close
End Sub

Public Sub cancelaReq()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   db.BeginTrans
   db.Execute ("DELETE FROM servrequisicao WHERE id = ('" & cmbReq & "')")
   db.CommitTrans
   FechaDB
   MsgBox ("Exluido com sucesso!"), vbInformation
Exit Sub
Erro: MsgBox ("Erro ao cancelar requisição: " & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Sub carregaTela()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM servrequisicaodetalhe WHERE id = ('" & cmbReq & "')", db, 3, 3
   
   If Not rs.EOF Then
   
      Do While Not rs.EOF
         
         tblServicos.AddItem pegaNomeServ(rs!Grupo, rs!Classe, rs!codServ) & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!codServ
         rs.MoveNext
         
      Loop
   
   End If
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar tela: " & Err.Description), vbInformation
FechaDB
End Sub

Public Function pegaNomeServ(Grupo As String, Classe As String, codServ As String)
   On Error GoTo Erro
   
   Prod.Open "SELECT descricao FROM servservico WHERE grupo = ('" & Grupo & "') AND classe = ('" & Classe & "') AND codServ = ('" & codServ & "')", db, 3, 3
   
      pegaNomeServ = Prod!Descricao
   
   Prod.Close
   
Exit Function
Erro: MsgBox ("Erro ao pegar nome: " & Err.Description), vbInformation
Prod.Close
End Function

Public Sub limpaTela()
   cmbGrupo = Empty
   cmbClasse.Clear
   cmbServ.Clear
   tblServicos.Rows = 1
   cmbReq.Clear
   carregaReq
End Sub

Public Sub criaReqCompraServ(i As Integer)
   On Error GoTo Erro
   db.Execute ("INSERT INTO servrequisicaocompra(grupo,classe,codServ,idReq) VALUES ('" & tblServicos.TextMatrix(i, 1) & "','" & tblServicos.TextMatrix(i, 2) & "','" & tblServicos.TextMatrix(i, 3) & "','" & cmbReq & "')")
Exit Sub
Erro: MsgBox ("Erro ao criar requisição de compra: " & Err.Description), vbInformation
End Sub
