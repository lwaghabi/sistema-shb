VERSION 5.00
Begin VB.Form frmSenhaSupervisor 
   Caption         =   "Senha do Supervisor"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtSenhaSupervisor 
      BackColor       =   &H0000C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtNomeSupervisor 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Supervisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   0
      Top             =   1200
      Width           =   2430
   End
End
Attribute VB_Name = "frmSenhaSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()

'Set TabUsuario = dbSHB.OpenRecordset("Usuario")
'    TabUsuario.Index = "IndUsuario"
    
TabUsuario.Seek "=", txtNomeSupervisor

If TabUsuario.NoMatch Then
   MsgBox ("Usuario não habilitado como Supervisor")
  'TabUsuario.Close
   Unload Me
Else
   If Not (TabUsuario("ususenha") = txtSenhaSupervisor) Then
      MsgBox ("Senha Incorreta")
     'TabUsuario.Close
      Unload Me
   Else
      If Not (TabUsuario("usutipoacesso") = 1) Then
         MsgBox ("Usuario não é Supervisor")
         'TabUsuario.Close
         Unload Me
      Else
         'TabUsuario.Close
         frmAbre_Fecha.Show
      End If
   End If
End If
End Sub
