VERSION 5.00
Begin VB.Form frmEspecTec 
   Caption         =   "frmEspecTec"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEspecificacaoTecnica 
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   7455
   End
   Begin VB.TextBox txtDescricao 
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   9015
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   975
      Left            =   8040
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "frmEspecTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
   Unload Me
End Sub

