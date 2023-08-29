VERSION 5.00
Begin VB.Form frmProdutosTerceiros 
   Caption         =   "Produtos de Teceiros"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Navegação"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   3495
      Begin VB.CommandButton cmdNavega 
         Caption         =   "Último"
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNavega 
         Caption         =   "Próximo"
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNavega 
         Caption         =   "Anterior"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNavega 
         Caption         =   "Primeiro"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operações"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   3495
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtTipoProduto 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmProdutosTerceiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Novo As Byte
Dim Incluir As Byte
Dim Resp As String

Private Sub cmdExcluir_Click()
Dim Resp As String
Resp = MsgBox("Exclusão de Registro. Confirma?", vbYesNo)
If Resp = vbYes Then

   Call Rotina_AbrirBanco
   
   ProdTerc.Open "Select * from ProdutoTerceiros where chTipoProduto = ('" & txtTipoProduto & "')", db, 3, 3
   If ProdTerc.EOF Then
      MsgBox ("Produto Terceiros solicitado para exclusão não encontrado."), vbInformation
      Call FechaDB
      Exit Sub
    End If


   db.BeginTrans
         
   ProdTerc.Delete
   
   txtTipoProduto = Empty
   
  db.CommitTrans
End If
      
Call FechaDB
      
End Sub

Private Sub cmdIncluir_Click()
   Novo = 0
   
   Call Rotina_AbrirBanco
   
   ProdTerc.Open "Select * from ProdutoTerceiros where chTipoProduto = ('" & txtTipoProduto & "')", db, 3, 3
   If Not (ProdTerc.EOF) Then
      MsgBox ("Produto Terceiros solicitado para inclusão já existe."), vbInformation
      Call FechaDB
      Exit Sub
    End If
     
   db.BeginTrans
      
   ProdTerc.AddNew
   
   ProdTerc!chTipoProduto = txtTipoProduto
   
   ProdTerc.Update
  db.CommitTrans
   cmdNovo.Enabled = True
   cmdExcluir.Enabled = True
   cmdSair.Enabled = True
   
   txtTipoProduto = Empty
   
   Call FechaDB
   
   Unload Me
End Sub

Private Sub cmdNavega_Click(Index As Integer)

Call Rotina_AbrirBanco

ProdTerc.Open "Select * from Produtoterceiros", db, 3, 3
If ProdTerc.EOF Then
   MsgBox ("Erro no acesso a Produto Terceiros em Navegação"), vbCritical
   Call FechaDB
   Exit Sub
End If
   

Select Case Index

   Case 0
        ProdTerc.MoveFirst
   Case 1
        ProdTerc.MoveNext
   Case 2
        ProdTerc.MovePrevious
   Case 3
        ProdTerc.MoveLast
        
End Select

   If ProdTerc.BOF = True Then
      ProdTerc.MoveFirst
   End If
   
   If ProdTerc.EOF = True Then
      ProdTerc.MoveLast
   End If
   
   txtTipoProduto = ProdTerc!chTipoProduto
   txtTipoProduto.Enabled = False
   
   cmdNovo.Enabled = True
   cmdIncluir.Enabled = False
   cmdExcluir.Enabled = True
   cmdSair.Enabled = True
   
   Call FechaDB
   
End Sub

Private Sub cmdNovo_Click()
   Incluir = 0
   Novo = 1
   txtTipoProduto = Empty
   txtTipoProduto.Enabled = True
   txtTipoProduto.SetFocus
   cmdIncluir.Enabled = True
   cmdExcluir.Enabled = False
   cmdSair.Enabled = True
   
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmProdutosTerceiros.txtTipoProduto = frmProdutosDeEntrada.cmbTipoProduto
End Sub

Private Sub txtTipoProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtTipoProduto_LostFocus()

If txtTipoProduto = Empty Then
   MsgBox ("Informe o tipo de produto")
   txtTipoProduto.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

ProdTerc.Open "Select * from ProdutoTerceiros where chTipoProduto = ('" & txtTipoProduto & "')", db, 3, 3
If ProdTerc.EOF Then
   Resp = MsgBox("Inclusão de Produto. Confirma???", vbYesNo)
   If Resp = vbYes Then
      Incluir = 1
      cmdIncluir.Enabled = True
      cmdExcluir.Enabled = False
      cmdSair.Enabled = True
      Exit Sub
   Else
      txtTipoProduto.SetFocus
      Exit Sub
   End If
Else
    Incluir = 0
    txtTipoProduto = ProdTerc!chTipoProduto
    txtTipoProduto.Enabled = False
         
    cmdNovo.Enabled = True
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = True
    cmdSair.Enabled = True
  
End If

Call FechaDB

End Sub
