VERSION 5.00
Begin VB.Form frmFornecFabrica 
   Caption         =   "Produtos de Terceiros - Código Fábrica"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form3"
   ScaleHeight     =   3210
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Navegação"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   3975
      Begin VB.CommandButton cmdNavega 
         Caption         =   "Úlrimo"
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNavega 
         Caption         =   "Próximo"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNavega 
         Caption         =   "Anterior"
         Height          =   375
         Index           =   1
         Left            =   1080
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
      Height          =   2175
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   855
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.TextBox txtTipoProduto 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtCodigoFabrica 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código Fábrica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmFornecFabrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Resp As String
Dim Incluir As Byte

Private Sub txtTipoProduto_LostFocus()

'Call Rotina_AbrirBanco

ProdTerc.Open "Select * from ProdutoTerceiros where chTipoProduto = ('" & txtTipoProduto & "')", db, 3, 3
If ProdTerc.EOF Then
   MsgBox ("Verifique a tabela ProdutoTerceiros")
   Call FechaDB
   Exit Sub
End If

If ProdFornec.State = 1 Then
   ProdFornec.Close: Set ProdFornec = Nothing
End If

End Sub

Private Sub cmdExcluir_Click()
Dim Resp As String
Resp = MsgBox("Exclusão de Registro. Confirma?", vbYesNo)
If Resp = vbYes Then
   
   'Call Rotina_AbrirBanco

   ProdFornec.Open "Select * from ProdutoFornecedor wher chTipoProduto = ('" & txtTipoProduto & "') and chProdutoFabrica = ('" & txtCodigoFabrica & "')", db, 3, 3
   If ProdFornec.EOF Then
      MsgBox ("Solicitação de exclusão inválida. O registro não existe."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   
   db.BeginTrans
         
   ProdFornec.Delete
   
  db.CommitTrans
   
   txtTipoProduto = Empty
   txtCodigoFabrica = Empty
End If


If ProdFornec.State = 1 Then
   ProdFornec.Close: Set ProdFornec = Nothing
End If

End Sub

Private Sub cmdIncluir_Click()

'Call Rotina_AbrirBanco

ProdFornec.Open "Select * from ProdutoFornecedor where chTipoProduto = ('" & txtTipoProduto & "') and chProdutoFabrica = ('" & txtCodigoFabrica & "')", db, 3, 3
If ProdFornec.EOF Then
   
   db.BeginTrans

   ProdFornec.AddNew
   ProdFornec!chTipoProduto = txtTipoProduto
   ProdFornec!chProdutoFabrica = txtCodigoFabrica
   ProdFornec.Update

  db.CommitTrans
End If
   
cmdExcluir.Enabled = True
cmdSair.Enabled = True

txtTipoProduto = Empty
txtCodigoFabrica = Empty

If ProdFornec.State = 1 Then
   ProdFornec.Close: Set ProdFornec = Nothing
End If

Unload Me
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub txtCodigoFabrica_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCodigoFabrica_LostFocus()

Call Rotina_AbrirBanco

ProdFornec.Open "Select * from ProdutoFornecedor where chTipoProduto = ('" & txtTipoProduto & "') and chProdutoFabrica = ('" & txtCodigoFabrica & "')", db, 3, 3
If ProdFornec.EOF Then
   Incluir = 1
   cmdIncluir.Enabled = True
   cmdExcluir.Enabled = False
   cmdSair.Enabled = True
Else
    Incluir = 0
    txtCodigoFabrica = ProdFornec!chProdutoFabrica
    txtCodigoFabrica.Enabled = False
    
    cmdIncluir.Enabled = False
    cmdExcluir.Enabled = True
    cmdSair.Enabled = True
End If


If ProdFornec.State = 1 Then
   ProdFornec.Close: Set ProdFornec = Nothing
End If

End Sub

