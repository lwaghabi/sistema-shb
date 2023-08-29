VERSION 5.00
Begin VB.Form frmUsuarioSenha 
   BackColor       =   &H8000000D&
   Caption         =   "Semi Hermatics do Brasil"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "UsuarioSenha.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   4455
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H000000FF&
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
         Height          =   615
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdEntra 
         BackColor       =   &H0000FF00&
         Caption         =   "Entra"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         DataField       =   "txtSenha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   2  'Center
         DataField       =   "txtUsuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtStatusSistema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Senha"
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
         Left            =   840
         TabIndex        =   8
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Usuário"
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
         Left            =   960
         TabIndex        =   7
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Status do Sistema"
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
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Sistema Integrado Semi Hermatics do Brasil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmUsuarioSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data_Login As Date
Dim Resp As String
Dim Computador As String
Dim Estacao As String
Dim StatusAbertura As Byte
Dim ErroAbertura As Byte
Dim DataLogin As Double
Dim DataInvertida As String

Dim Ano As Integer
Dim Mes As Integer
Dim Dia As Integer


Private Sub cmdEntra_Click()

Ano = Year(Date)
Mes = Month(Date)
Dia = Day(Date)

Call CriticaAbertura

If ErroAbertura = 1 Then
   ErroAbertura = 0
   Exit Sub
End If

Call Rotina_AbrirBanco

DataInvertida = Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
   
glb.Open "select * from Global where chDataAbertura = " & DataInvertida, db, 3, 3
acGlb = acGlb + 1
If glb.EOF Then
   If glbTipoAcesso = 3 Then
      If glbStusSistema = 0 Then
         MsgBox ("Aguarde a abertura do sistema"), vbCritical
         End
      End If
   End If
End If

If ErroAbertura = 0 Then
   usu.Open "select * from Usuario where chNome = ('" & txtUsuario & "')", db, 3, 3
   If usu.EOF Then
      MsgBox ("Usuário inválido. "), vbCritical
      Call FechaDB
      Exit Sub
   Else
      Call StatusAcesso
      usu!usustatus = 1
      usu!usumaquina = glbMaquina
      usu.Update
      FechaDB
      Unload Me
   End If
Else
   MsgBox ("Erro na abertura"), vbCritical
   'Unload Me
   End
End If
End Sub

Public Sub CriticaAbertura()


Call Rotina_AbrirBanco
ErroAbertura = 0


usu.Open "select * from Usuario where chNome = ('" & txtUsuario & "')", db, 3, 3
acUsu = acUsu + 1
If usu.EOF Then
   MsgBox ("Usuario/Senha inválido"), vbCritical
   ErroAbertura = 1
   cmdSair.SetFocus
   Exit Sub
Else
   If usu!usuSenha = txtSenha Then
      glbUsuario = usu!chNome
      glbTipoAcesso = usu!usuTipoAcesso
      glbStatusUsuario = usu!usustatus
   Else
      MsgBox ("Senha inválida."), vbCritical
      cmdSair.SetFocus
      ErroAbertura = 1
   End If
End If

Call FechaDB

End Sub


Public Sub StatusAcesso()

'txtMaquina.Text = GetIPHostName()
'txtEnderecoIP.Text = GetIPAddress()

If glbTipoAcesso = 1 Then
   mdiSHB.mdiPessoa.Enabled = True
   mdiSHB.mdiConsultsEspeciais.Enabled = True
   mdiSHB.mdiCadCli.Enabled = True
   mdiSHB.mdiNeg.Enabled = True
   mdiSHB.mdiProcesEControles.Enabled = True
   mdiSHB.mdiColaboradores = True
   mdiSHB.mdiParametros.Enabled = True
   mdiSHB.mdiFinanceiro.Enabled = True
   mdiSHB.mdiControleFinanceiro = True
   mdiSHB.mdiRecebimentos = True
   mdiSHB.mdiPagamentos = True
   mdiSHB.mdiReprogFinanc = True
   mdiSHB.mdiConsultaFinanc = True
   mdiSHB.mdiSupervisor.Enabled = True
   mdiSHB.mdiControleFinanceiro = True
   mdiSHB.mdiConsultaFinanc = True
   mdiSHB.mdiProducao.Enabled = True
   mdiSHB.mdiMovProducao.Enabled = True
   mdiSHB.mdiMoveEspecial.Enabled = True
   mdiSHB.mdiMateriaisEst.Enabled = True
   mdiSHB.mdiRelatorios.Enabled = True
   mdiSHB.mdiHabilitacao.Enabled = True
   mdiSHB.mdiSupervisor.Enabled = True
Else
   If glbTipoAcesso = 3 And glbStusSistema = 1 Then
      mdiSHB.mdiPessoa.Enabled = True
      mdiSHB.mdiCadCli = True
      mdiSHB.mdiConsultsEspeciais.Enabled = False
      mdiSHB.mdiProcesEControles.Enabled = False
      mdiSHB.mdiColaboradores = False
      mdiSHB.mdiParametros.Enabled = False
      mdiSHB.mdiFinanceiro.Enabled = False
      mdiSHB.mdiProducao.Enabled = False
      mdiSHB.mdiMateriaisEst.Enabled = False
      mdiSHB.mdiRelatorios.Enabled = False
      mdiSHB.mdiHabilitacao.Enabled = False
      mdiSHB.mdiSupervisor.Enabled = False
   Else
      If Not usu!usuTipoAcesso = 0 Then
         mdiSHB.mdiPessoa.Enabled = True
         mdiSHB.mdiConsultsEspeciais.Enabled = False
         mdiSHB.mdiNeg.Enabled = True
         mdiSHB.mdiProcesEControles.Enabled = True
         mdiSHB.mdiColaboradores = True
         mdiSHB.mdiParametros.Enabled = True
         mdiSHB.mdiFinanceiro.Enabled = False
         mdiSHB.mdiProducao.Enabled = True
         mdiSHB.mdiMateriaisEst.Enabled = True
         mdiSHB.mdiHabilitacao.Enabled = True
         mdiSHB.mdiRelatorios.Enabled = True
         If glbTipoAcesso = 3 Then
            mdiSHB.mdiSupervisor.Enabled = False
            mdiSHB.mdiNeg.Enabled = True
            mdiSHB.mdiProcesEControles.Enabled = False
            mdiSHB.mdiMovProducao.Enabled = False
            mdiSHB.mdiParametros.Enabled = False
            mdiSHB.mdiMoveEspecial.Enabled = False
            mdiSHB.mdiCadCli = True
         Else
            mdiSHB.mdiCadCli = True
            mdiSHB.mdiFinanceiro.Enabled = True
            mdiSHB.mdiControleFinanceiro = True
            mdiSHB.mdiRecebimentos = False
            mdiSHB.mdiPagamentos = False
            mdiSHB.mdiReprogFinanc = False
            mdiSHB.mdiConsultaFinanc = False
            mdiSHB.mdiMovProducao.Enabled = True
            mdiSHB.mdiMoveEspecial.Enabled = True
            mdiSHB.mdiSupervisor.Enabled = True
         End If
      Else
         mdiSHB.mdiHabilitacao.Enabled = False
      End If
   End If
End If

glbMaquina = GetIPHostName()

glbEnderecoIP = GetIPAddress()


End Sub


Private Sub cmdSair_Click()
MsgBox ("O Sistema será encerrado"), vbCritical
End
End Sub

Private Sub Form_Load()

Ano = Year(Date)
Mes = Month(Date)
Dia = Day(Date)
txtUsuario = Empty
txtSenha = Empty

DataLogin = Date
    
'If Compilando Then
'   txtUsuario = "lwaghabi"
'   txtSenha = "morena"
'End If

Call Rotina_AbrirBanco

DataHojeInvertida = Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

glb.Open "Select * from Global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3

If glb.EOF Then
   'MsgBox ("Atenção: Sistema encontra-se fechado."), vbInformation
   txtStatusSistema = "Fechado"
Else
   txtStatusSistema = "Aberto"
End If

End Sub



