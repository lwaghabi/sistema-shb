VERSION 5.00
Begin VB.Form frmAtualizaFinanceiro 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAtualizarFinanceiro 
      Caption         =   "Atualizar Lançamentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Entrada contacorrente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmAtualizaFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualizarFinanceiro_Click()
    On Error GoTo Erro:
    Call Rotina_AbrirBanco
    FechaDB
Exit Sub
Erro: MsgBox "", vbInformation
FechaDB
End Sub

Private Sub Form_Load()

End Sub
