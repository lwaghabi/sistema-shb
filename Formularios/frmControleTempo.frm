VERSION 5.00
Begin VB.Form frmControleTempo 
   Caption         =   "frmControleTempo"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   840
   End
End
Attribute VB_Name = "frmControleTempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Timer1.Interval = 10000
   Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
   If Date <> Data_Hoje Then
      MsgBox "Você está logado com data diferente de hoje, o sistema será encerrado."
      End
   End If
End Sub
