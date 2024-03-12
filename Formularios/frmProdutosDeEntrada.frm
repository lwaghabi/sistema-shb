VERSION 5.00
Begin VB.Form frmProdutosDeEntrada 
   Caption         =   "Cadastramento de Produtos e Serviços (Fornecedor/Despesa - frmProdutosEntrada"
   ClientHeight    =   7845
   ClientLeft      =   2415
   ClientTop       =   1980
   ClientWidth     =   12345
   LinkTopic       =   "Form4"
   ScaleHeight     =   7845
   ScaleWidth      =   12345
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbOrigemProd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Frame Frame4 
      Height          =   7815
      Left            =   1080
      TabIndex        =   20
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtDataHoje 
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
         Left            =   7680
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Caption         =   "Função de Atualização"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   14
         Top             =   6840
         Width           =   3855
         Begin VB.CommandButton cmdNovo 
            BackColor       =   &H00FFFF00&
            Caption         =   "Novo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdInclui 
            BackColor       =   &H0000FF00&
            Caption         =   "Salvar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdAltera 
            BackColor       =   &H0000FFFF&
            Caption         =   "Alterar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdExclui 
            BackColor       =   &H000000FF&
            Caption         =   "Excluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Navegação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4320
         TabIndex        =   30
         Top             =   6840
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdNavega 
            BackColor       =   &H000000FF&
            Caption         =   "Último"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdNavega 
            BackColor       =   &H00FFFF00&
            Caption         =   "Anter."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdNavega 
            BackColor       =   &H000080FF&
            Caption         =   "Próx."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdNavega 
            BackColor       =   &H0000FF00&
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   8520
         TabIndex        =   29
         Top             =   6840
         Width           =   1095
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H0000FFFF&
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
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   6135
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   9975
         Begin VB.ComboBox cmbSubGrupoCentroDeCusto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   360
            TabIndex        =   4
            Top             =   4440
            Width           =   5895
         End
         Begin VB.TextBox txtChaveEnvio 
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame Frame 
            BackColor       =   &H00E0E0E0&
            Height          =   735
            Left            =   360
            TabIndex        =   32
            Top             =   240
            Width           =   8415
            Begin VB.Label lblLabel8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cadastramento de Serviços e Produtos de Fornecedores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               TabIndex        =   33
               Top             =   240
               Width           =   7950
            End
         End
         Begin VB.ComboBox txtUnidade 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2760
            TabIndex        =   7
            Top             =   5520
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtPrazoEntrega 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   840
            TabIndex        =   6
            Top             =   5520
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cmbFornecedor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4800
            TabIndex        =   1
            Top             =   1680
            Width           =   3855
         End
         Begin VB.ComboBox cmbProdutoFabrica 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   360
            TabIndex        =   3
            Top             =   3600
            Width           =   5895
         End
         Begin VB.TextBox txtDescricaoProduto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   360
            MaxLength       =   50
            TabIndex        =   5
            Top             =   5400
            Width           =   8175
         End
         Begin VB.ComboBox cmbTipoProduto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   360
            TabIndex        =   2
            Top             =   2640
            Width           =   5895
         End
         Begin VB.TextBox txtQtdUnidade 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5160
            TabIndex        =   8
            Top             =   5520
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtPesoUnidade 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6960
            TabIndex        =   9
            Top             =   5520
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sub Grupo de Centro de Custo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            TabIndex        =   38
            Top             =   4080
            Width           =   3795
         End
         Begin VB.Label lblLabel10 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Origem do Produto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   330
            TabIndex        =   36
            Top             =   1200
            Width           =   2640
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Prazo Entrega (dias)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   31
            Top             =   5160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Serviço/Fornecedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   28
            Top             =   1200
            Width           =   3855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Grupo De Centro de Custo "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            TabIndex        =   27
            Top             =   3240
            Width           =   3360
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Descrição do Produto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            TabIndex        =   26
            Top             =   5040
            Width           =   8175
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Código do  Produto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            TabIndex        =   25
            Top             =   2280
            Width           =   2610
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Unidade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3000
            TabIndex        =   24
            Top             =   5160
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Qtd./Unid."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5160
            TabIndex        =   23
            Top             =   5160
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Peso/Unid."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6840
            TabIndex        =   22
            Top             =   5160
            Visible         =   0   'False
            Width           =   1380
         End
      End
      Begin VB.Label lblLabel9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hoje"
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
         Left            =   6960
         TabIndex        =   34
         Top             =   240
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmProdutosDeEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProdutoAnterior As String
Dim Resp As String
Dim TamanhoCampo As Integer
Dim centrodecusto As String
Dim GrupoCentroDeCusto As String
Dim SubGrupoCentroDeCusto As String
Dim Descricao As String


Private Sub cmbFornecedor_LostFocus()

If txtChaveEnvio = "NFE" Then
   txtChaveEnvio = Empty
   Exit Sub
End If

cmbTipoProduto.Clear

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where chPessoa = ('" & cmbFornecedor & "')", db, 3, 3
If pes.EOF Then
   If cmbOrigemProd = "FORNECEDOR" Then
      MsgBox ("Efetuar o Cadastro deste fornecedor em pessoa e somente após o cadastramento lançar Centro de Custo"), vbCritical
      Call FechaDB
      cmdSair.SetFocus
      Exit Sub
   End If
Else
   If cmbOrigemProd = "DESPESA" Then
      MsgBox ("Efetuar este lançamento como fornecedor."), vbCritical
      Call FechaDB
      cmdSair.SetFocus
      Exit Sub
   End If
End If

If cmbOrigemProd = "FORNECEDOR" Then
   ProdEntrada.Open "Select * from produtoentrada where chPessoa = ('" & cmbFornecedor & "')", db, 3, 3
   If ProdEntrada.EOF Then
      cmbTipoProduto.SetFocus
   Else
      ProdEntrada.MoveFirst
      
      Do While Not ProdEntrada.EOF
         cmbTipoProduto.AddItem ProdEntrada!chTipoProduto
         ProdEntrada.MoveNext
      Loop
   End If
Else
   ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbFornecedor & "')", db, 3, 3
   If Not ProdFornec.EOF Then
      ProdFornec.MoveFirst
      Do While Not ProdFornec.EOF
         If Not ProdFornec!chProdutoFabrica = ProdutoAnterior Then
            cmbTipoProduto.AddItem ProdFornec!chProdutoFabrica
            ProdutoAnterior = ProdFornec!chProdutoFabrica
         End If
         ProdFornec.MoveNext
      Loop
   End If
End If

Call FechaDB

End Sub

Private Sub cmbOrigemProd_LostFocus()

If Not txtChaveEnvio = "NFE" Then
   cmbFornecedor.Clear
   cmbTipoProduto.Clear
   If cmbOrigemProd = "DESPESA" Then
      Call CarregaDespesa
   Else
      Call carregaFornecedor
   End If
   cmbFornecedor.SetFocus
Else
   'txtChaveEnvio = Empty
   cmbProdutoFabrica.SetFocus
End If

End Sub

Private Sub cmbProdutoFabrica_LostFocus()

Verifica = Empty
Verifica = Mid$(cmbTipoProduto, 30, 5)
If Not Verifica = Empty Then
   MsgBox ("Código do Produto Informado ultrapassa 30 caracteres.")
   cmbTipoProduto.SetFocus
   Exit Sub
End If
   
If cmbProdutoFabrica = Empty Then
   MsgBox ("Informar um centro de custo. Esta informação é obrigatória."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If
   
Call Rotina_AbrirBanco
   
txtUnidade = "Un"
txtPesoUnidade = 0
txtPrazoEntrega = 0
txtQtdUnidade = 0

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and DescricaoCentroDeCusto = ('" & cmbProdutoFabrica & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

centrodecusto = "2"
GrupoCentroDeCusto = Format$(Prod!chGrupoCentroDeCusto, "00")
SubGrupoCentroDeCusto = "00"

If Not cmbSubGrupoCentroDeCusto = Empty Then
   Descricao = cmbSubGrupoCentroDeCusto
Else
   Descricao = Empty
End If
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   
   Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & GrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("Centro de Custo não registrado em produtoentrada"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   cmbSubGrupoCentroDeCusto.Clear
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
      cmbSubGrupoCentroDeCusto.AddItem Prod!DescricaoCentroDeCusto
      Prod.MoveNext
   Loop
If Not Descricao = Empty Then
   cmbSubGrupoCentroDeCusto = Descricao
End If

End Sub



Private Sub cmbTipoProduto_LostFocus()

If cmbOrigemProd = "FORNECEDOR" Then
   Call TratarFornecedor
Else
   Call TratarDespesa
End If
   
End Sub

Private Sub cmdAltera_Click()

Call Rotina_AbrirBanco

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
      
Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and DescricaoCentroDeCusto = ('" & cmbProdutoFabrica & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Não encontrei Centro de custo."), vbCritical
   Call FechaDB
   Exit Sub
End If

GrupoCentroDeCusto = Prod!chGrupoCentroDeCusto

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & GrupoCentroDeCusto & "') and DescricaoCentroDeCusto = ('" & cmbSubGrupoCentroDeCusto & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Não encontrei Centro de custo."), vbCritical
   Call FechaDB
   Exit Sub
End If

SubGrupoCentroDeCusto = Prod!chSubGrupoCentroDeCusto

If cmbOrigemProd = "FORNECEDOR" Then
   ProdEntrada.Open "Select * from produtoentrada where chPessoa = ('" & cmbFornecedor & "') and chTipoProduto = ('" & cmbTipoProduto & "')", db, 3, 3
   If ProdEntrada.EOF Then
      MsgBox ("Produto solicitado para alteração não encontrado."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   If ProdEntrada.State = 1 Then
      ProdEntrada.Close: Set ProdEntrada = Nothing
   End If
   Call Rotina_010_Form_DB
Else
   ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbFornecedor & "') and chProdutoFabrica = ('" & cmbTipoProduto & "')", db, 3, 3
   If ProdFornec.EOF Then
      MsgBox ("Produto solicitado para alteração não encontrado em Despesa."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   If ProdFornec.State = 1 Then
      ProdFornec.Close: Set ProdFornec = Nothing
   End If
   
   Call Rotina_015_Despesa

End If

Call Rotina_020_Limpa_Form
cmdNovo.SetFocus

End Sub

Private Sub cmdExclui_Click()

Resp = MsgBox("Exclusão de Registro. Cinfirma???", vbYesNo)

If Resp = vbNo Then
   Exit Sub
End If

Call Rotina_AbrirBanco

db.BeginTrans

If cmbOrigemProd = "FORNECEDOR" Then
   ProdEntrada.Open "Select * from produtoentrada where chPessoa = ('" & cmbFornecedor & "') and chTipoProduto = ('" & cmbTipoProduto & "') and chProdutoFabrica = ('" & cmbProdutoFabrica & "')", db, 3, 3
   If ProdEntrada.EOF Then
      MsgBox ("Produto solicitado para Exclusão não encontrado."), vbCritical
      Call FechaDB
      Exit Sub
   Else
      ProdEntrada.Delete
   End If
Else
   ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbFornecedor & "') and chProdutoFabrica = ('" & cmbTipoProduto & "') and chCentroDeCusto = ('" & cmbProdutoFabrica & "')", db, 3, 3
      If ProdFornec.EOF Then
         MsgBox ("Produto solicitado para Exclusão não encontrado em Despesa."), vbCritical
         Call FechaDB
         Exit Sub
      Else
         ProdFornec.Delete
      End If
End If

db.CommitTrans

MsgBox ("Registro Excluido"), vbInformation

Call Rotina_020_Limpa_Form

txtChaveEnvio = Empty
cmdSair.SetFocus
End Sub

Private Sub cmdInclui_Click()

TamanhoCampo = Len(cmbFornecedor)

If TamanhoCampo > 20 Then
   MsgBox ("Tamnho do campo ultrapassa a 20 caracteres."), vbCritical
   Call FechaDB
   Exit Sub
End If

If cmbFornecedor = Empty Then
   MsgBox ("Serviço/Fornecedor não informado. Inclusão será abortada"), vbCritical
   Exit Sub
End If
   
If cmbTipoProduto = Empty Then
   MsgBox ("Codigo do Produto não informado. A inclusão será abortada."), vbCritical
   Exit Sub
End If

If txtDescricaoProduto = Empty Then
   MsgBox ("Descrição do Produto não informado"), vbInformation
   txtDescricaoProduto.SetFocus
End If
 
If txtUnidade = Empty Then
   txtUnidade = "Un"
End If

If txtQtdUnidade = Empty Then
   txtQtdUnidade = 0
End If

If txtPesoUnidade = Empty Then
   txtPesoUnidade = 0
End If

Call Rotina_AbrirBanco

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
      
Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and DescricaoCentroDeCusto = ('" & cmbProdutoFabrica & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Não encontrei Centro de custo."), vbCritical
   Call FechaDB
   Exit Sub
End If

GrupoCentroDeCusto = Prod!chGrupoCentroDeCusto

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & GrupoCentroDeCusto & "') and DescricaoCentroDeCusto = ('" & cmbSubGrupoCentroDeCusto & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Não encontrei Centro de custo."), vbCritical
   Call FechaDB
   Exit Sub
End If

SubGrupoCentroDeCusto = Prod!chSubGrupoCentroDeCusto

If cmbOrigemProd = "FORNECEDOR" Then
   Call Rotina_010_Form_DB
Else
   Call Rotina_015_Despesa
End If

MsgBox ("Processamento efetuado com Sucesso"), vbInformation

Call Rotina_020_Limpa_Form


Unload Me

End Sub

Private Sub cmdNavega_Click(Index As Integer)
  
  Call Rotina_AbrirBanco
  
  ProdEntrada.Open "Select * from produtoentrada", db, 3, 3
  
  Select Case Index

   Case 0
        ProdEntrada.MoveFirst
   Case 1
        ProdEntrada.MoveNext
   Case 2
        ProdEntrada.MovePrevious
   Case 3
        ProdEntrada.MoveLast
        
End Select

   If ProdEntrada.BOF = True Then
      ProdEntrada.MoveFirst
   End If
   
   If ProdEntrada.EOF = True Then
      ProdEntrada.MoveLast
   End If

Call Rotina_020_Limpa_Form

Call Rotina_030_Enche_Tela

cmdAltera.Enabled = True
cmdExclui.Enabled = True
cmdInclui.Enabled = False

If ProdEntrada.State = 1 Then
   ProdEntrada.Close: Set ProdEntrada = Nothing
End If

End Sub

Private Sub cmdNovo_Click()
Call Rotina_020_Limpa_Form
'lblFornecedor.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

centrodecusto = "2"
GrupoCentroDeCusto = "00"
SubGrupoCentroDeCusto = "00"
'cmbControlarEstoque.AddItem "Não"
'cmbControlarEstoque.AddItem "Sim"

'cmbControlarEstoque.ListIndex = 0

txtUnidade.Clear
txtDataHoje = Date

Call Rotina_AbrirBanco

UnidEmb.Open "Select * from unidadeembalagem", db, 3, 3
If UnidEmb.EOF Then
   MsgBox ("Unidade de embalagem não cadastrada. "), vbCritical
   Call FechaDB
   Exit Sub
End If

UnidEmb.MoveFirst
Do While Not UnidEmb.EOF
   txtUnidade.AddItem UnidEmb!UnidadeEmbalagem
   UnidEmb.MoveNext
Loop
txtUnidade.ListIndex = 0

'ProdTerc.Open "Select * from produtoterceiros", db, 3, 3
'If ProdTerc.EOF Then
'   MsgBox ("Produto terceiros não cadastrado."), vbCritical
'   Call FechaDB
'   Exit Sub
'Else
'   ProdTerc.MoveFirst
'End If

'Do While Not ProdTerc.EOF
'   cmbProdutoFabrica.AddItem ProdTerc!chTipoProduto
'   ProdTerc.MoveNext
'Loop

cmbProdutoFabrica.Clear

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & centrodecusto & "') and chGrupoCentroDeCusto > ('" & GrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto = ('" & SubGrupoCentroDeCusto & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Erro. Tabela de Centro de Custo Vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

Prod.MoveFirst

Do While Not Prod.EOF
   cmbProdutoFabrica.AddItem Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop



If ProdTerc.State = 1 Then
   ProdTerc.Close: Set ProdTerc = Nothing
End If
If UnidEmb.State = 1 Then
   UnidEmb.Close: Set UnidEmb = Nothing
End If

cmbOrigemProd.AddItem "FORNECEDOR"
cmbOrigemProd.AddItem "DESPESA"

End Sub

Public Sub Rotina_010_Form_DB()

db.BeginTrans

ProdEntrada.Open "Select * from produtoentrada where chPessoa = ('" & cmbFornecedor & "') and chTipoProduto = ('" & cmbTipoProduto & "')", db, 3, 3
If ProdEntrada.EOF Then
   ProdEntrada.AddNew
End If

ProdEntrada!chPessoa = cmbFornecedor
ProdEntrada!chTipoProduto = cmbTipoProduto
ProdEntrada!chProdutoFabrica = cmbProdutoFabrica
ProdEntrada!pinDescricao = txtDescricaoProduto
ProdEntrada!chCodProduto = cmbFornecedor
ProdEntrada!pinUnidade = txtUnidade
ProdEntrada!pinPesoLiquidoUnidade = txtPesoUnidade
ProdEntrada!pinQtdUnidade = txtQtdUnidade

ProdEntrada!pinCentroDeCusto = centrodecusto
ProdEntrada!pinGrupoCentroDeCusto = GrupoCentroDeCusto
ProdEntrada!pinSubGrupoCentroDeCusto = SubGrupoCentroDeCusto

ProdEntrada!pinaliquotaicms = 0
ProdEntrada!pinaliquotaipi = 0
ProdEntrada!pinClassificacao = "S/C"
ProdEntrada!pinCntrlEstoque = 0
ProdEntrada!pinPrazoEntrega = txtPrazoEntrega


ProdEntrada.Update

db.CommitTrans

'MsgBox ("Alteração efetuada com sucesso"), vbInformation

End Sub
Public Sub Rotina_015_Despesa()

Call Rotina_AbrirBanco

db.BeginTrans
 
   ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbFornecedor & "') and chProdutoFabrica = ('" & cmbTipoProduto & "')", db, 3, 3
   If ProdFornec.EOF Then
      ProdFornec.AddNew
   End If
   
   ProdFornec!chTipoProduto = cmbFornecedor
   ProdFornec!chProdutoFabrica = cmbTipoProduto
   ProdFornec!chCentroDeCusto = cmbProdutoFabrica
   ProdFornec!pinCentroDeCusto = centrodecusto
   ProdFornec!pinGrupoCentroDeCusto = GrupoCentroDeCusto
   ProdFornec!pinSubGrupoCentroDeCusto = SubGrupoCentroDeCusto

   ProdFornec.Update

db.CommitTrans

'MsgBox ("Inclusão/Alteração em Despesa efetuada com sucesso"), vbInformation

End Sub
Public Sub Rotina_020_Limpa_Form()

cmbFornecedor = Empty
cmbTipoProduto = Empty
cmbProdutoFabrica = Empty
txtDescricaoProduto = Empty
'lblCodProduto = Empty
txtUnidade = Empty
txtQtdUnidade = Empty
txtPesoUnidade = Empty
cmbProdutoFabrica.Clear
cmbSubGrupoCentroDeCusto.Clear
'txtICMS = Empty
'txtIPI = Empty
'txtClassificacao = Empty
End Sub

Public Sub Rotina_030_Enche_Tela()
cmbFornecedor = ProdEntrada!chPessoa
cmbTipoProduto.Text = ProdEntrada!chTipoProduto
cmbProdutoFabrica.Text = ProdEntrada!chProdutoFabrica
txtDescricaoProduto = ProdEntrada!pinDescricao
'lblCodProduto = ProdEntrada!chCodProduto
txtUnidade = ProdEntrada!pinUnidade
txtPesoUnidade = ProdEntrada!pinPesoLiquidoUnidade
txtQtdUnidade = ProdEntrada!pinQtdUnidade

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ProdEntrada!pinGrupoCentroDeCusto & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Erro. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
Else
   cmbProdutoFabrica = Prod!DescricaoCentroDeCusto
End If

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
   
Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ProdEntrada!pinGrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto = ('" & ProdEntrada!pinGrupoCentroDeCusto & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Erro. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
Else
   cmbSubGrupoCentroDeCusto = Prod!DescricaoCentroDeCusto
End If


'txtICMS = ProdEntrada!pinaliquotaicms
'txtIPI = ProdEntrada!pinaliquotaipi
'txtClassificacao = ProdEntrada!pinclassificacao
'cmbControlarEstoque.ListIndex = ProdEntrada!pinCntrlEstoque
txtPrazoEntrega = ProdEntrada!pinPrazoEntrega
End Sub

Private Sub txtClassificacao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCodProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub



Private Sub txtDescricaoProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUnidade_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Public Sub CarregaDespesa()

ProdutoAnterior = Empty

Call Rotina_AbrirBanco

ProdFornec.Open "Select * from produtofornecedor", db, 3, 3
   If ProdFornec.EOF Then
      MsgBox ("Erro. Tabela de Produto fornecedor vazia. Comunicar ao analista responsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   ProdFornec.MoveFirst
   
   Do While Not ProdFornec.EOF
      If Not (ProdFornec!chTipoProduto = ProdutoAnterior) Then
         cmbFornecedor.AddItem ProdFornec!chTipoProduto
         ProdutoAnterior = ProdFornec!chTipoProduto
      End If
         
      ProdFornec.MoveNext

   Loop
Call FechaDB

End Sub

Public Sub carregaFornecedor()

ProdutoAnterior = Empty

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where pesTipoPessoa = ('" & 2 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro na carga pessoa"), vbCritical
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

Do While Not pes.EOF
   cmbFornecedor.AddItem pes!chPessoa
   pes.MoveNext
Loop

Call FechaDB
End Sub

Public Sub TratarFornecedor()
Call Rotina_AbrirBanco

ProdFornec.Open "Select * from produtoentrada where chPessoa = ('" & cmbFornecedor & "') and chTipoProduto = ('" & cmbTipoProduto & "')", db, 3, 3
If Not ProdFornec.EOF Then
   Call PreparaAlteracaoFornecedor
Else
   Call PreparaInclusao
End If
End Sub

Public Sub TratarDespesa()

Call Rotina_AbrirBanco

ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbFornecedor & "') and chProdutoFabrica = ('" & cmbTipoProduto & "')", db, 3, 3
If Not ProdFornec.EOF Then
   Call PreparaAlteracao
Else
   Call PreparaInclusao
End If
   
End Sub

Public Sub PreparaAlteracao()

cmbProdutoFabrica = ProdFornec!chCentroDeCusto
txtDescricaoProduto = ProdFornec!chProdutoFabrica
txtUnidade = "Unidade"
txtPesoUnidade = 0
txtPrazoEntrega = 0
txtQtdUnidade = 0
'cmdAltera.Enabled = True
'cmdAltera.SetFocus

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto > ('" & "00" & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbProdutoFabrica.Clear

Prod.MoveFirst

Do While Not Prod.EOF
   cmbProdutoFabrica.AddItem Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
   
Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ProdFornec!pinGrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Grupo Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbProdutoFabrica = Prod!DescricaoCentroDeCusto

Call CarregaSubGrupo

cmbProdutoFabrica.SetFocus
      
End Sub

Public Sub PreparaInclusao()

Call Rotina_AbrirBanco

cmbProdutoFabrica.Clear
cmbSubGrupoCentroDeCusto.Clear

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto > ('" & "00" & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

Prod.MoveFirst

Do While Not Prod.EOF
   cmbProdutoFabrica.AddItem Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop

cmbProdutoFabrica.SetFocus

cmdAltera.Enabled = False
cmdInclui.Enabled = True

End Sub


Public Sub CarregaSubGrupo()

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto > ('" & ProdFornec!pinGrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbSubGrupoCentroDeCusto.Clear

Prod.MoveFirst

Do While Not Prod.EOF
   cmbSubGrupoCentroDeCusto.AddItem Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
   
Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ProdFornec!pinGrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto = ('" & ProdFornec!pinSubGrupoCentroDeCusto & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Grupo Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbSubGrupoCentroDeCusto = Prod!DescricaoCentroDeCusto

End Sub

Public Sub PreparaAlteracaoFornecedor()

cmbProdutoFabrica = ProdFornec!chTipoProduto
txtDescricaoProduto = ProdFornec!pinDescricao
txtUnidade = "Unidade"
txtPesoUnidade = 0
txtPrazoEntrega = 0
txtQtdUnidade = 0
cmdAltera.Enabled = True
'cmdAltera.SetFocus

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto > ('" & "00" & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbProdutoFabrica.Clear

Prod.MoveFirst

Do While Not Prod.EOF
   cmbProdutoFabrica.AddItem Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
   
Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ProdFornec!pinGrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Grupo Centro de Custo não registrado em produtoentrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbProdutoFabrica = Prod!DescricaoCentroDeCusto

Call CarregaSubGrupo

cmbProdutoFabrica.SetFocus
      
End Sub

