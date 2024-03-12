VERSION 5.00
Begin VB.Form frmParmServico 
   Caption         =   "frmParmServico"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   15600
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1095
      Left            =   11520
      TabIndex        =   20
      Top             =   8040
      Width           =   1335
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
      Height          =   1095
      Left            =   13080
      TabIndex        =   19
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      TabIndex        =   18
      Top             =   8040
      Width           =   1455
   End
   Begin VB.ComboBox cmbStatus 
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
      Left            =   7440
      TabIndex        =   17
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Classificação de Centro de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   840
      TabIndex        =   11
      Top             =   6840
      Width           =   6375
      Begin VB.ComboBox cmbSubGrupoCentroCusto 
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
         Left            =   1680
         TabIndex        =   15
         Top             =   1320
         Width           =   4455
      End
      Begin VB.ComboBox cmbGrupoCentroCusto 
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
         Left            =   1680
         TabIndex        =   14
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Sub-Grupo"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox txtEspecTec 
      Height          =   2535
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4080
      Width           =   13575
   End
   Begin VB.ComboBox cmbUnidServ 
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
      Left            =   840
      TabIndex        =   8
      Top             =   2760
      Width           =   2655
   End
   Begin VB.ComboBox cmbServico 
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
      Left            =   7920
      TabIndex        =   5
      Top             =   1560
      Width           =   6495
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
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
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
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Status"
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
      Left            =   7440
      TabIndex        =   16
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Especificação Técnica"
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
      Left            =   840
      TabIndex        =   9
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Unidade Serviço"
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
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Serviço"
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
      Left            =   7920
      TabIndex        =   6
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Classe 
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Grupo 
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
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Registro e Atualização de Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frmParmServico"
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

Private Sub cmbGrupoCentroCusto_LostFocus()
   Call Rotina_AbrirBanco
   Call carregaSubGrupoCentroDeCusto
End Sub

Private Sub cmbServico_LostFocus()
   Call carregaInfo
End Sub

Private Sub cmdExcluir_Click()
   Call excluiServico
   Call limpaTela
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
   Call salvarServico
   Call limpaTela
End Sub

Private Sub Form_Load()
   Call carregaUnidade
   Call carregaStatus
   Call carregaGrupoCentroDeCusto
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
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar classes" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaServico()
   Dim NomeTab As String
   Dim coluna As String

   If cmbClasse = "ASO" Then
   
      NomeTab = "asoexame -- "
      coluna = "chNomeExame"
   
   ElseIf cmbClasse = "TREINAMENTO" Then
      
      NomeTab = "treinamento  -- "
      coluna = "chNomeCurso"
      
   Else
   
      NomeTab = "servservico"
      coluna = "descricao"
   
   End If
   
   cmbServico.Clear
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select " & coluna & " from " & NomeTab & " where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') order by codServ", db, 3, 3
   
   If Prod.EOF Then
   
      MsgBox ("Não há serviços cadastrados nessa categoria"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbServico.AddItem Prod(0)
      Prod.MoveNext
   
   Loop
   
   FechaDB

Exit Sub
Erro: MsgBox ("Erro ao carregar serviços: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaInfo()
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   rs.Open "Select * from servservico where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and descricao = ('" & cmbServico & "')", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Inclusão de novo serviço."), vbYesNo
      FechaDB
      Exit Sub
   End If
   
   cmbUnidServ.ListIndex = rs!unidade - 1
   txtEspecTec = rs!especTec
   cmbGrupoCentroCusto.ListIndex = rs!GrupoCentroDeCusto - 1
   Call carregaSubGrupoCentroDeCusto
   cmbSubGrupoCentroCusto.ListIndex = rs!SubGrupoCentroDeCusto - 1
   cmbStatus.ListIndex = rs!Status
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar informações do serviço" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaUnidade()
   On Erro GoTo Erro
   Call Rotina_AbrirBanco
   Prod.Open "Select * from unidadedeservicos", db, 3, 3
   If Prod.EOF Then
      MsgBox "Erro: Unidades de serviços não cadastradas", vbCritical
      FechaDB
      Exit Sub
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbUnidServ.AddItem Prod!abrevunidserv
      Prod.MoveNext
   
   Loop
   Prod.Close
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar unidades de serviços: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaStatus()
   cmbStatus.AddItem "Ativo"
   cmbStatus.AddItem "Inativo"
Exit Sub
Erro: MsgBox ("Erro ao carregar status"), vbInformation
End Sub

Public Sub carregaGrupoCentroDeCusto()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   rs.Open "Select DescricaoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and chGrupoCentroDeCusto>'00' and chSubGrupoCentroDeCusto='00' ", db, 3, 3
   
   rs.MoveFirst
   
   Do While Not rs.EOF
   
      cmbGrupoCentroCusto.AddItem rs!DescricaoCentroDeCusto
      rs.MoveNext
      
   Loop
   
   
   rs.Close
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar grupo de centro de custo: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaSubGrupoCentroDeCusto()
   On Error GoTo Erro
   
   Dim grupodecusto As String
   
   Prod.Open "Select chGrupoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and DescricaoCentroDeCusto=('" & cmbGrupoCentroCusto & "') and chSubGrupoCentroDeCusto='00'", db, 3, 3
   
   grupodecusto = Prod!chGrupoCentroDeCusto
   
   Prod.Close
   
   pes.Open "Select DescricaoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and chGrupoCentroDeCusto=('" & grupodecusto & "') and chSubGrupoCentroDeCusto>'00' ", db, 3, 3
   
   pes.MoveFirst
   
   Do While Not pes.EOF
   
      cmbSubGrupoCentroCusto.AddItem pes!DescricaoCentroDeCusto
      pes.MoveNext
      
   Loop
   
   
   pes.Close
   
Exit Sub
Erro: MsgBox ("Erro ao carregar sub-grupo de centro de custo " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub salvarServico()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   db.BeginTrans
   
   rs.Open "SELECT * FROM servservico WHERE grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') AND classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') AND descricao = ('" & cmbServico & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
      rs!Grupo = Format$(cmbGrupo.ListIndex + 1, "00")
      rs!Classe = Format$(cmbClasse.ListIndex + 1, "000")
      rs!codServ = ultimoRegistro
   Else
      rs!Grupo = rs!Grupo
      rs!Classe = rs!Classe
      rs!codServ = rs!codServ
   End If
   
   rs!Descricao = cmbServico
   rs!unidade = cmbUnidServ.ListIndex + 1
   rs!especTec = txtEspecTec
   rs!GrupoCentroDeCusto = cmbGrupoCentroCusto.ListIndex + 1
   rs!SubGrupoCentroDeCusto = cmbSubGrupoCentroCusto.ListIndex + 1
   rs!Status = cmbStatus.ListIndex
   
   rs.Update
   rs.Close
   
   db.CommitTrans
   
   FechaDB
   
   MsgBox ("Salvo com sucesso!"), vbInformation
Exit Sub
Erro: MsgBox ("Erro ao salvar serviço" & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Function ultimoRegistro() As String
   On Error GoTo Erro
   Dim retorno As String
   Dim codigo As Integer
   Prod.Open "SELECT MAX(codServ) codigo FROM servservico WHERE grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') AND classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "')", db, 3, 3
   If IsNull(Prod!codigo) Then
      codigo = 0
   Else
      codigo = Prod!codigo
   End If
   retorno = Format$(codigo + 1, "00000")
   Prod.Close
   ultimoRegistro = retorno
Exit Function
Erro: MsgBox ("Erro ao gerar novo código" & Err.Description), vbInformation
End Function

Public Sub excluiServico()
   On Error GoTo Erro
   db.BeginTrans
   Call Rotina_AbrirBanco
   
   db.Execute ("DELETE FROM servservico WHERE grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') AND classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') AND descricao = ('" & cmbServico & "')")
   
   db.CommitTrans
   FechaDB
   MsgBox ("Excluído com sucesso!"), vbInformation
Exit Sub
Erro: MsgBox ("Erro ao excluir serviço: " & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Sub limpaTela()
   cmbServico = Empty
   cmbUnidServ.ListIndex = 0
   txtEspecTec = Empty
   cmbGrupoCentroCusto = Empty
   cmbSubGrupoCentroCusto = Empty
   cmbStatus.ListIndex = 0
End Sub
