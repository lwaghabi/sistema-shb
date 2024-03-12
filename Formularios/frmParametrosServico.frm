VERSION 5.00
Begin VB.Form frmParametrosServico 
   Caption         =   "frmParametrosServico"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
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
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
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
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
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
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Grupos"
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
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "frmParametrosServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbGrupo_LostFocus()
Call carregaClasse
End Sub

Private Sub cmdIncluir_Click()
   Call salvaClasse
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Call carregaGrupo
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

Public Sub salvaClasse()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   db.BeginTrans
   pes.Open "SELECT * FROM servgrupoclasse WHERE grupo = '" & Format(cmbGrupo.ListIndex + 1, "00") & "' AND descricao = '" & cmbClasse & "'", db, 3, 3
   If pes.EOF Then
      rs.Open "SELECT COUNT(classe) as classe FROM servgrupoclasse WHERE grupo = '" & Format(cmbGrupo.ListIndex + 1, "00") & "'", db, 3, 3
      db.Execute ("INSERT INTO servgrupoclasse(grupo,classe,descricao) VALUES ('" & Format(cmbGrupo.ListIndex + 1, "00") & "','" & Format(rs!Classe, "000") & "','" & cmbClasse & "')")
      rs.Close
   End If
   pes.Close
   db.CommitTrans
   FechaDB
   MsgBox ("Criado com sucesso!"), vbInformation
Exit Sub
Erro: MsgBox ("Erro ao salvar classe: " & Err.Description), vbInformation
db.RollbackTrans
End Sub
