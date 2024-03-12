VERSION 5.00
Begin VB.Form frmUnidadeOperacional 
   Caption         =   "frmUnidadeOperacional"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   16845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   6375
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   13095
      Begin VB.ComboBox cmbCliente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
         Width           =   4335
      End
      Begin VB.ComboBox cmbUnidadeOperacional 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7440
         TabIndex        =   2
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtOperadora 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   3
         Top             =   2520
         Width           =   4335
      End
      Begin VB.TextBox txtSiglaUnidadeOperadora 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7920
         TabIndex        =   4
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox txtContato 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   5
         Top             =   3480
         Width           =   4455
      End
      Begin VB.TextBox txtCelTel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6240
         TabIndex        =   6
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9120
         TabIndex        =   7
         Top             =   3480
         Width           =   2895
      End
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H0000FF00&
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H000000FF&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lblUnidadeOperacional 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade Operacional"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   16
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblOperadora 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operadora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label lblSiglaDa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla da Unid. Operacional"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   14
         Top             =   2160
         Width           =   3210
      End
      Begin VB.Label lblContatoNa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato na Unid Operacional"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   3120
         Width           =   3510
      End
      Begin VB.Label lblCelularTelefone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Celular/Telefone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   12
         Top             =   3120
         Width           =   2085
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         TabIndex        =   11
         Top             =   3120
         Width           =   675
      End
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastramento e atualização de Unidades Operacionais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   19
      Top             =   960
      Width           =   7935
   End
   Begin VB.Label lblCadastramentoE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastramento e atualização de Unidades Operacionais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   18
      Top             =   -600
      Width           =   7935
   End
End
Attribute VB_Name = "frmUnidadeOperacional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Erro As Byte


Private Sub cmbCliente_LostFocus()

Call Rotina_AbrirBanco

uoper.Open "Select * from unidadeoperacional where chPessoa = ('" & cmbCliente & "')", db, 3, 3
If Not uoper.EOF Then
   cmbUnidadeOperacional.Clear
   uoper.MoveFirst
   Do While Not uoper.EOF
      cmbUnidadeOperacional.AddItem uoper!chUnidadeOperacional
      uoper.MoveNext
   Loop
Else
   cmbUnidadeOperacional.Clear
End If

Call FechaDB

End Sub



Private Sub cmbUnidadeOperacional_LostFocus()

If cmbUnidadeOperacional = Empty Then
   MsgBox ("Não informado a Unidade Operacional para esta operação."), vbInformation
   Call FechaDB
   Exit Sub
End If

Call Rotina_AbrirBanco

uoper.Open "Select * from unidadeoperacional where chPessoa = ('" & cmbCliente & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "')", db, 3, 3
If Not uoper.EOF Then
   txtOperadora = uoper!uopOperadora
   txtSiglaUnidadeOperadora = uoper!uopSiglaOperadora
   txtContato = uoper!uopContatoUnidOper
   txtCelTel = uoper!uopTelefone
   txtEmail = uoper!iopemail
Else
   txtOperadora = Empty
   txtSiglaUnidadeOperadora = Empty
   txtContato = Empty
   txtCelTel = Empty
   txtEmail = Empty
End If
 
 
End Sub

Private Sub cmdExcluir_Click()
Call Rotina_AbrirBanco
   
uoper.Open "Select * from unidadeoperacional where chPessoa = ('" & cmbCliente & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "')", db, 3, 3
If Not uoper.EOF Then

   uoper.Delete
   
   MsgBox ("Unidade Operacional atualizado com sucesso."), vbInformation
   txtOperadora = Empty
   txtSiglaUnidadeOperadora = Empty
   txtContato = Empty
   txtCelTel = Empty
   txtEmail = Empty
   cmbUnidadeOperacional = Empty
Else
   MsgBox ("Solicitação inválida para deleção"), vbCritical
End If

Call FechaDB

cmbCliente.SetFocus

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()

Call CriticarCampos

If Erro = 1 Then
   MsgBox ("Verificar informações digitadas. Foram encontrados erros."), vbCritical
   Call FechaDB
   Exit Sub
End If
   
Call Rotina_AbrirBanco
   
uoper.Open "Select * from unidadeoperacional where chPessoa = ('" & cmbCliente & "') and chUnidadeOperacional = ('" & cmbUnidadeOperacional & "')", db, 3, 3
If uoper.EOF Then
   uoper.AddNew
End If

uoper!chPessoa = cmbCliente
uoper!chUnidadeOperacional = cmbUnidadeOperacional
uoper!uopOperadora = txtOperadora
uoper!uopSiglaOperadora = txtSiglaUnidadeOperadora
uoper!uopContatoUnidOper = txtContato
uoper!uopTelefone = txtCelTel
uoper!iopemail = txtEmail

uoper.Update

MsgBox ("Unidade Operacional atualizado com sucesso."), vbInformation
txtOperadora = Empty
txtSiglaUnidadeOperadora = Empty
txtContato = Empty
txtCelTel = Empty
txtEmail = Empty

cmbUnidadeOperacional = Empty

cmbCliente.SetFocus

Call FechaDB

End Sub

Private Sub Form_Load()

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where pesTipoPessoa = ('" & 0 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Tabela de Clientes vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If
   
pes.MoveFirst

Do While Not pes.EOF
   cmbCliente.AddItem pes!chPessoa
   pes.MoveNext
Loop

End Sub


Public Sub CriticarCampos()
Erro = 0
If txtOperadora = Empty Then
   Erro = 1
End If
If txtSiglaUnidadeOperadora = Empty Then
   Erro = 1
End If
If txtContato = Empty Then
   Erro = 1
End If
If txtCelTel = Empty Then
   Erro = 1
End If
If txtEmail = Empty Then
   Erro = 1
End If
End Sub
Private Sub cmbUnidadeOperacional_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtOperadora_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtSiglaUnidadeOperadora_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtContato_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtceltel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


