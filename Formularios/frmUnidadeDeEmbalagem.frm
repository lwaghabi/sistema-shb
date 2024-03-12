VERSION 5.00
Begin VB.Form frmUnidadeDeEmbalagem 
   Caption         =   "frmUnidadeDeEmbalagem"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
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
      Height          =   735
      Left            =   3840
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
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
      Height          =   735
      Left            =   2160
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdCadastrar 
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
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtDescricao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   9495
   End
   Begin VB.TextBox txtAbreviatura 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.ComboBox cmbUnidadeDeEmbalagem 
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
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   9495
   End
   Begin VB.Label Label3 
      Caption         =   "Abreviatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Unidade de Embalagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Cadastro de Umidade de Embalagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   7695
   End
End
Attribute VB_Name = "frmUnidadeDeEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbUnidadeDeEmbalagem_LostFocus()
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM unidadeembalagem WHERE unidadeembalagem = ('" & cmbUnidadeDeEmbalagem & "')", db, 3, 3
  
   If rs.EOF Then
   
      MsgBox ("Não existem unidades de embalagens cadastradas"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   txtAbreviatura = rs!AbreviaturaUnidadeEmbalagem
   txtDescricao = rs!DescricaoUnidadeEmbalagem
   
   rs.Close
   
   FechaDB
End Sub

Private Sub cmdCadastrar_Click()
Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM unidadeembalagem WHERE unidadeembalagem = ('" & cmbUnidadeDeEmbalagem & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
      rs!indice = cmbUnidadeDeEmbalagem.ListCount
   
   Else
   
      rs!indice = rs!indice
   
   End If
   
   rs!UnidadeEmbalagem = cmbUnidadeDeEmbalagem
   rs!AbreviaturaUnidadeEmbalagem = txtAbreviatura
   rs!DescricaoUnidadeEmbalagem = txtDescricao
   
   
   rs.Update
   rs.Close
   
   MsgBox ("Salvo com sucesso!"), vbInformation
   
   FechaDB
End Sub

Private Sub cmdExcluir_Click()
   Dim indice As Integer
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT indice FROM unidadeembalagem WHERE unidadeembalagem = ('" & cmbUnidadeDeEmbalagem & "')", db, 3, 3
   If Not rs.EOF Then
   db.BeginTrans
   
   indice = rs!indice
   
   db.Execute ("DELETE FROM unidadeembalagem WHERE indice = '" & indice & "'")
   
   db.Execute ("UPDATE unidadeembalagem SET indice = indice - 1 WHERE indice > '" & indice & "' ")
   
   db.CommitTrans
   
   MsgBox ("Unidade de Medida excluida com sucesso!"), vbInformation
   
   Else
   
   MsgBox ("Unidade de Medida não cadastrada!"), vbInformation
   
   End If
   rs.Close
   FechaDB
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM unidadeembalagem", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não existem unidades de embalagens cadastradas"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   rs.MoveFirst
   
   Do While Not rs.EOF
      
      cmbUnidadeDeEmbalagem.AddItem rs!UnidadeEmbalagem
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   FechaDB
End Sub
