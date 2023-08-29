VERSION 5.00
Begin VB.Form frmUsuario 
   Caption         =   "frmUsuario  (Usuário Senha)"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6270
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbPessoa 
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
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Frame frmTipoAcesso 
      Caption         =   "Tipo de Acesso"
      Height          =   1215
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   2775
      Begin VB.OptionButton optConsultas 
         Caption         =   "Consultas"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton optLançamentos 
         Caption         =   "Lançamentos e Atualizações"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optAdministrador 
         Caption         =   "Administrador"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdNavega 
      Caption         =   "&Último"
      Height          =   495
      Index           =   3
      Left            =   2520
      TabIndex        =   15
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdNavega 
      Caption         =   "&Ant."
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   14
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdNavega 
      Caption         =   "&Próx."
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdNavega 
      Caption         =   "&Início"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtSenha 
      DataField       =   "txtSenha"
      Height          =   475
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Código no Pessoa"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   240
      Top             =   0
      Width           =   2820
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Area As Workspace
'Dim dbLartMerco As Database
'Dim usu As Recordset
'Dim TabIndex As Index
Dim Incluir As Byte
Dim baseaberta As Byte
Dim msg As String
Dim Nome As String
Dim rotinicial As Byte
Dim Administrador As Byte



Private Sub cmdAlterar_Click()
    On Error Resume Next
   
If Not (Administrador = 1) Then
   MsgBox "Função somente permitida a administradores"
   Call FechaDB
   Unload Me
   Exit Sub
End If

If Not glbUsuario = txtNome Then
   If Not (Administrador = 1) Then
      MsgBox "Função somente permitida a administradores"
      Exit Sub
   End If
End If

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where chNome = ('" & txtNome & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Usuario não permitido"), vbCritical
   Call FechaDB
   End
End If


db.BeginTrans

   usu!chNome = txtNome
   
   usu!usuSenha = txtSenha
   
   If Administrador = 1 Then
      If optAdministrador = True Then
         usu!usuTipoAcesso = 1
      Else
         If optLançamentos = True Then
            usu!usuTipoAcesso = 2
         Else
            usu!usuTipoAcesso = 3
         End If
      End If
   End If
   
   If usu!usuTipoAcesso = 1 Then
      optAdministrador = True
      optLançamentos = False
      optConsultas = False
   Else
      If usu!usuTipoAcesso = 2 Then
         optAdministrador = False
         optLançamentos = True
         optConsultas = False
      Else
         optAdministrador = False
         optLançamentos = False
         optConsultas = True
      End If
    End If
   
   usu!usuSenha = txtSenha
   
   usu.Update
db.CommitTrans

MsgBox ("Alteração efetuada com sucesso"), vbInformation

Call FechaDB

End Sub

Private Sub cmdExcluir_Click()
Dim Resp As String

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where chNome = ('" & txtNome & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Exclusão de usuário não cadastrado"), vbCritical
   Call FechaDB
   Exit Sub
End If

If Not (glbUsuario = txtNome) Then
   If Not Administrador = 1 Then
      MsgBox ("Função reservada a Administradores")
      Call FechaDB
      cmdFechar.SetFocus
      Exit Sub
   End If
Else
   MsgBox ("Não permitido deletar a sí proprio"), vbCritical
   Call FechaDB
   cmdFechar.SetFocus
   Exit Sub
End If

Resp = MsgBox("Voce esta prestes a deletar este usuário. Confirma???", vbYesNo)
If Resp = vbYes Then

   usu.Delete
   MsgBox ("Usuário deletado."), vbInformation

End If

Call FechaDB

Unload Me

End Sub

Private Sub cmdFechar_Click()
       Unload Me
       'MenuParametro.Show vbModal
       baseaberta = 0
End Sub

Private Sub cmdIncluir_Click()
On Error Resume Next

If Not Administrador = 1 Then
   MsgBox "Função exclusiva para Administradores"
   Exit Sub
End If


If cmbPessoa = Empty Then
   MsgBox ("ERRO: Codigo pessoa não informado. Deve ser informado o codigo para cadastramento em Pessoa."), vbInformation
   Exit Sub
End If
   
Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where chNome = ('" & txtNome & "')", db, 3, 3
If usu.EOF Then
   usu.AddNew
End If

db.BeginTrans

usu!chNome = txtNome
usu!chPessoa = cmbPessoa
 
usu!usuSenha = txtSenha
   
If Administrador = 1 Then
   If optAdministrador = True Then
      usu!usuTipoAcesso = 1
   Else
      If optLançamentos = True Then
         usu!usuTipoAcesso = 2
      Else
         usu!usuTipoAcesso = 3
      End If
   End If
End If

usu!usustatus = 0
usu!usuordem = 99

cmdNovo.Enabled = True
cmdIncluir.Enabled = True
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True


usu.Update

db.CommitTrans

'Limpa campos do doc. de entrada

txtNome = Empty
txtSenha = Empty

optAdministrador = False
optLançamentos = False
optConsultas = False

cmdAlterar.Enabled = True
cmdExcluir.Enabled = True
cmdIncluir.Enabled = True
cmdFechar.Enabled = True

cmdFechar.SetFocus

MsgBox ("Usuário icluído com sucesso"), vbInformation

Call FechaDB

End Sub

Private Sub cmdNavega_Click(Index As Integer)
Select Case Index

   Case 0
        usu.MoveFirst
   Case 1
        usu.MoveNext
   Case 2
        usu.MovePrevious
   Case 3
        usu.MoveLast
        
End Select

   If usu.BOF = True Then
      usu.MoveFirst
   End If
   
   If usu.EOF = True Then
      usu.MoveLast
   End If
   
   txtNome = usu("chnome")
   txtSenha = usu("usuSenha")
   
  If usu("usuTipoAcesso") = True Then
     optAdministrador.SetFocus
   Else
       If usu("usuTipoAcesso") = True Then
          optLançamentos.SetFocus
      Else
          optConsultas.SetFocus
      End If
   End If
   
   cmdNovo.Enabled = True
   cmdIncluir.Enabled = False
   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   cmdFechar.Enabled = True
   
   txtNome.Enabled = False
   txtSenha.SetFocus
   
End Sub

Private Sub cmdNovo_Click()

      txtNome.Enabled = True
      
      txtNome.SetFocus
      
      txtNome = Empty
      
      optLançamentos = False
      optConsultas = False
      
      txtSenha = Empty
      
      cmdNovo.Enabled = True
      cmdAlterar.Enabled = True
      cmdExcluir.Enabled = True
      cmdIncluir.Enabled = True
      cmdFechar.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
  
Administrador = 0

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox "Usuario não permitido"
   Call FechaDB
   Exit Sub
Else
   If usu!usuTipoAcesso = 1 Then
      Administrador = 1
   Else
      Administrador = 0
 End If
End If

usu.MoveFirst
rotinicial = 0

pes.Open "Select chPessoa from Pessoa where pesTipoPessoa = 7 and pesStatusPessoa = 0", db, 3, 3
If pes.EOF Then
   MsgBox ("ERRO: Não encontrado Colaboradores no Pessoa."), vbCritical
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst
Do While Not pes.EOF
   cmbPessoa.AddItem pes!chPessoa
   pes.MoveNext
Loop



Call FechaDB
 
End Sub





Private Sub txtNome_GotFocus()
    txtNome.Enabled = True
    
End Sub


Private Sub txtNome_LostFocus()
On Error Resume Next

If Not (glbUsuario = txtNome) Then
   If Not Administrador = 1 Then
      MsgBox "Função reservada para administradores. LostFocus"
      txtNome = Empty
      txtSenha = Empty
      cmdFechar.SetFocus
      Exit Sub
   End If
End If

If rotinicial = 1 Then
   rotinicial = 0
   txtNome = usu("chNome")
   Nome = txtNome
   optAdministrador.SetFocus
   txtSenha.SetFocus
Else
   If txtNome = Empty Or txtNome = " " Then
      txtNome = Nome
      MsgBox "Campo obrigatório."
      txtNome.SetFocus
   End If
End If

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where chNome = ('" & txtNome & "')", db, 3, 3
If usu.EOF Then
   Incluir = 1
Else
   Incluir = 0
   txtSenha = usu!usuSenha
    
    
   If usu!usuTipoAcesso = 1 Then
      optAdministrador = True
      optLançamentos = False
      optConsultas = False
   Else
      If usu!usuTipoAcesso = 2 Then
         optAdministrador = False
         optLançamentos = True
         optConsultas = False
      Else
         optAdministrador = False
         optLançamentos = False
         optConsultas.SetFocus
      End If
   End If
   
   cmdNovo.Enabled = True
   cmdIncluir.Enabled = False
   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   cmdFechar.Enabled = True
   
End If

Call FechaDB

End Sub


