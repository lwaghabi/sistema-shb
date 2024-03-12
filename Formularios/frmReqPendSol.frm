VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReqPendSol 
   Caption         =   "frmReqPendSol"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid tblProdutos 
      Height          =   1815
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Produto                                          |Quantidade"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstRequisicao 
      Height          =   1815
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Requisições Pendentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Requisições Pendentes de Separação"
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
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "frmReqPendSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT id FROM suprequisicaodetalhe WHERE codigo !="" and status=0 GROUP BY id", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não existem requisições pendentes."), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   rs.MoveFirst
   
   Do While Not rs.EOF
   
      lstRequisicao.AddItem rs!Id
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   Call FechaDB
End Sub

Private Sub lstRequisicao_Click()
   Call Rotina_AbrirBanco
   
   Call FechaDB
End Sub
