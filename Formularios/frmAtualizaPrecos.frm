VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAtualizaPrecoProd 
   Caption         =   "frmAtualizaPrecoProd Atualiza Preço da Tabela de Produtos"
   ClientHeight    =   8520
   ClientLeft      =   1335
   ClientTop       =   1530
   ClientWidth     =   12495
   LinkTopic       =   "Form3"
   ScaleHeight     =   8520
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtHoje 
      Height          =   495
      Left            =   9960
      TabIndex        =   21
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   600
      TabIndex        =   15
      Top             =   600
      Width           =   11055
      Begin VB.TextBox txtDesconto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4080
         TabIndex        =   4
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox cmbAtividade 
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
         Left            =   4080
         TabIndex        =   3
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   6240
         Width           =   8895
         Begin VB.Frame Frame5 
            Height          =   975
            Left            =   6960
            TabIndex        =   27
            Top             =   240
            Width           =   1695
            Begin VB.CommandButton cmdSair 
               BackColor       =   &H0000FF00&
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
               Height          =   525
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Processa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   6615
            Begin VB.CommandButton cmdExcluir 
               BackColor       =   &H008080FF&
               Caption         =   "Excluir"
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
               Left            =   4920
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H0000FFFF&
               Caption         =   "Cancela"
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
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   240
               Width           =   1495
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00FFFF00&
               Caption         =   "Confirma"
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
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   240
               Width           =   1495
            End
         End
      End
      Begin VB.ComboBox cmbTabPreco 
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
         Left            =   4080
         TabIndex        =   1
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtAjuste 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4080
         TabIndex        =   6
         Top             =   4920
         Width           =   1815
      End
      Begin VB.ComboBox cmbTipoAjuste 
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
         Left            =   4080
         TabIndex        =   5
         Top             =   3720
         Width           =   4935
      End
      Begin VB.ComboBox cmbProduto 
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
         Left            =   4080
         TabIndex        =   2
         Top             =   1320
         Width           =   3735
      End
      Begin VB.ComboBox cmbTipoTabPreco 
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
         Left            =   4080
         TabIndex        =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label lblLabel6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Perc. Desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   3525
      End
      Begin VB.Label Label 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atividade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label txtPrecoAtual 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4080
         TabIndex        =   13
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Preço Corrigido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   24
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Label txtPrecoAnterior 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4080
         TabIndex        =   12
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pco anter. ao ajuste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   23
         Top             =   4320
         Width           =   3615
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ajuste ou Pco Atual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   20
         Top             =   4920
         Width           =   3615
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Forma de Ajuste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   3495
      End
      Begin VB.Label txtDescProduto 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4080
         TabIndex        =   11
         Top             =   2520
         Width           =   4695
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contrato/Equipamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Tab de Preços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atualização de Preços de Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmAtualizaPrecoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Resp As String
Dim CidadeAnterior As String
Dim GrupoAnterior As String
Dim chaveatualiza As Integer
Dim StatusAtual As Integer
Dim PessoaAnterior As String
Dim fim As Integer

Private Sub cmbAtividade_LostFocus()

Call Rotina_AbrirBanco

ProdPreco.Open "Select * from ProdutoPreco where chPessoa = ('" & cmbTabPreco & "') and chProduto = ('" & cmbProduto & "') and chAtividade = ('" & cmbAtividade & "') and pdpStatus = ('" & 0 & "')", db, 3, 3
If ProdPreco.EOF Then
   txtPrecoAnterior = Format$(0, "#0.00")
Else
   txtPrecoAnterior = Format$(ProdPreco!pdpPrecoDoProduto, "#0.00")
End If
txtPrecoAtual = Empty
txtAjuste = Empty

Call FechaDB

End Sub


Private Sub cmbProduto_LostFocus()

If cmbAtividade = Empty Then
   MsgBox ("Informar a atividade"), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

If cmbProduto = " Todos" Then
   txtDescProduto = "Todos os Produtos"
   cmbTipoAjuste.ListIndex = 0
   txtAjuste = Empty
Else
   Call Rotina_AbrirBanco

   Prod.Open "Select * from Produto where chProduto = ('" & cmbProduto & "')", db, 3, 3
   
   If Prod.EOF Then
      MsgBox ("Entre com um código da lista"), vbInformation
      cmbTabPreco.SetFocus
      Exit Sub
   End If
   txtDescProduto = Prod!prdNomeProd

End If
'ProdPreco.Open "Select * from ProdutoPreco where chPessoa = ('" & cmbTabPreco & "') and chProduto = ('" & cmbProduto & "') and chAtividade = ('" & cmbAtividade & "') and pdpStatus = ('" & 0 & "')", db, 3, 3
'If ProdPreco.EOF Then
'   txtPrecoAnterior = Format$(0, "#0.00")
'Else
'   txtPrecoAnterior = Format$(ProdPreco!pdpPrecoDoProduto, "#0.00")
'End If
'txtPrecoAtual = Empty

Call FechaDB

End Sub

Private Sub cmbTabPreco_LostFocus()
If cmbTabPreco = Empty Then
   MsgBox ("Esta informação é obrigatória")
End If

cmbProduto.Clear

Call Rotina_AbrirBanco

Prod.Open "Select * from Produto where prdLocadora = ('" & cmbTabPreco & "') and prdOrdemApresentacao = ('" & 0 & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Cliente sem contrato e sem Equipamentos cadastrados."), vbCritical
   Call FechaDB
   Exit Sub
End If

Prod.MoveFirst

Do While Not Prod.EOF
   cmbProduto.AddItem Prod!chProduto
   Prod.MoveNext
Loop
cmbProduto.AddItem "MOB/DESMOB"
Call FechaDB

End Sub









'Private Sub cmbTipoTabPreco_LostFocus()


'cmbTabPreco = Empty
'cmbTipoAjuste = Empty
'cmbTabPreco = Empty
'cmbProduto.ListIndex = 0
'txtDescProduto = Empty
'PessoaAnterior = Empty
''cmbTipoTabPreco.ListIndex = 1

'If cmbTipoTabPreco.ListIndex = 0 Then
'   cmbTabPreco.Clear
'   cmbTabPreco.AddItem "GERAL"
'   cmbTabPreco.ListIndex = 0
'Else
'   cmbTabPreco.Clear
'   GrupoAnterior = Empty
'   CidadeAnterior = Empty
'
'   Call Rotina_AbrirBanco
'
'   pes.Open "Select * from Pessoa", db, 3, 3
'
'
''   pes.MoveFirst
'   Do While Not pes.EOF
'      'If cmbTipoTabPreco.ListIndex = 1 Then
'         If pes!pestipopessoa = 0 And pes!pesStatusPessoa = 0 Then
'            If Not pes!chPessoa = PessoaAnterior Then
'                   cmbTabPreco.AddItem pes!chPessoa
'                   PessoaAnterior = pes!chPessoa
'                   pes.MoveNext
'            Else
'               pes.MoveNext
'            End If
'         Else
'           pes.MoveNext
'         End If
'     ' End If
'   Loop
'      cmbTabPreco.SetFocus
''End If'

'End Sub


Private Sub cmdExcluir_Click()

Call Rotina_AbrirBanco

ProdPreco.Open "Select * from ProdutoPreco where chPessoa = ('" & cmbTabPreco & "') and chProduto = ('" & cmbProduto & "') and chAtividade = ('" & cmbAtividade & "')", db, 3, 3
If ProdPreco.EOF Then
   MsgBox ("Erro na deleção do Produto Preço. Comunicar ao analista responsável"), vbCritical
   Call FechaDB
   Exit Sub
End If

ProdPreco.Delete

MsgBox ("Registro excluído"), vbInformation

cmdSair.SetFocus

Call FechaDB

End Sub

Private Sub cmdSair_Click()

Unload Me

End Sub



Private Sub Form_Load()

txtHoje = Date
chaveatualiza = 0
StatusAtual = 0

Call Rotina_AbrirBanco

TabPreco.Open "Select * from TipoTabPreco", db, 3, 3

TabPreco.MoveFirst
Do While Not TabPreco.EOF
   cmbTipoTabPreco.AddItem TabPreco!ttpdescricaotipopreco
   TabPreco.MoveNext
Loop
cmbTipoTabPreco.ListIndex = 0

cmbProduto.AddItem "GERAL"

'Prod.Open "Select * from Produto", db, 3, 3


'Prod.MoveFirst

'Do While Not Prod.EOF
'   If Prod!prdtipo = 1 And cmbTabPreco = Prod!prdLocadora Then
'      cmbProduto.AddItem Prod!chProduto
'   End If
'   Prod.MoveNext
'Loop

Ativ.Open "Select * from Atividade", db, 3, 3
If Ativ.EOF Then
   MsgBox ("Erro na tabela Produto"), vbCritical
   Call FechaDB
   Exit Sub
End If

Ativ.MoveFirst
Do While Not Ativ.EOF
   cmbAtividade.AddItem Ativ!atvAtividade
   Ativ.MoveNext
Loop

cmbProduto.ListIndex = 0

cmbAtividade.ListIndex = 0

txtDescProduto = Empty

cmbTipoAjuste.AddItem " % Todos os Produtos"
cmbTipoAjuste.AddItem " % para o produto indicado"
cmbTipoAjuste.AddItem " Valor informado"

cmbTabPreco = Empty
cmbTipoAjuste = Empty
cmbTabPreco = Empty
cmbProduto.ListIndex = 0
txtDescProduto = Empty
PessoaAnterior = Empty
'cmbTipoTabPreco.ListIndex = 1

'If cmbTipoTabPreco.ListIndex = 0 Then
'   cmbTabPreco.Clear
'   cmbTabPreco.AddItem "GERAL"
'   cmbTabPreco.ListIndex = 0
'Else
   cmbTabPreco.Clear
   GrupoAnterior = Empty
   CidadeAnterior = Empty
   
   Call Rotina_AbrirBanco
   
   pes.Open "Select * from Pessoa", db, 3, 3
   
   
   pes.MoveFirst
   Do While Not pes.EOF
      'If cmbTipoTabPreco.ListIndex = 1 Then
         If pes!pestipopessoa = 0 And pes!pesStatusPessoa = 0 Then
            If Not pes!chPessoa = PessoaAnterior Then
                   cmbTabPreco.AddItem pes!chPessoa
                   PessoaAnterior = pes!chPessoa
                   pes.MoveNext
            Else
               pes.MoveNext
            End If
         Else
           pes.MoveNext
         End If
     ' End If
   Loop
'      cmbTabPreco.SetFocus
'End If
Call FechaDB

End Sub


Private Sub Command1_Click()

If txtAjuste = Empty Then
   MsgBox ("Este campo é obrigatorio")
   txtAjuste.SetFocus
Else
   If Not (IsNumeric(txtAjuste)) Then
      MsgBox ("Valor preenchido não é válido")
      txtAjuste.SetFocus
   End If
End If

If txtDesconto = Empty Then
   txtDesconto = 0
End If

'If cmbProduto.ListIndex = 0 Then
'   Resp = MsgBox("Correçao de Toda a Tabela de preços. Confirma???", vbYesNo)
'   If Resp = vbYes Then
'      cmbTipoAjuste.ListIndex = 0
'      MsgBox ("Chamar rotina de atualizaçao geral"), vbInformation
'   Else
'      cmbTabPreco.SetFocus
'   End If
'   Exit Sub
'Else
   Call Rotina_AbrirBanco

   ProdPreco.Open "Select * from ProdutoPreco where chpessoa = ('" & cmbTabPreco & "') and chProduto = ('" & cmbProduto & "') and chAtividade = ('" & cmbAtividade & "') and pdpstatus = ('" & 0 & "')", db, 3, 3
   If ProdPreco.EOF Then
      If cmbTipoAjuste.ListIndex < 2 Then
         MsgBox ("Não há preço nesta tabela para ser atualizado"), vbInformation
         cmbTipoTabPreco.SetFocus
      Else
         txtPrecoAnterior = Empty
         ProdPreco.Close: Set ProdPreco = Nothing
         Call Rotina_010_Atualiza_Preco_Prod
      End If
   Else
      ProdPreco.Close: Set ProdPreco = Nothing
      Call Rotina_010_Atualiza_Preco_Prod
   End If
'End If
cmbProduto.SetFocus
End Sub

Public Sub Rotina_010_Atualiza_Preco_Prod()

Call Rotina_AbrirBanco

ProdPreco.Open "Select * from ProdutoPreco where chPessoa = ('" & cmbTabPreco & "') and chproduto = ('" & cmbProduto & "') and chAtividade = ('" & cmbAtividade & "') and pdpStatus = ('" & 0 & "')", db, 3, 3
If ProdPreco.EOF Then
   ProdPreco.AddNew
   Call Rotina_014_Gera_Novo_Preco
Else
   chaveatualiza = 1
   ProdPreco.Close: Set ProdPreco = Nothing
   Call Rotina_012_Novo_Com_Anterior
   MsgBox ("Atualização concluída"), vbInformation
   cmbProduto.SetFocus
End If

Call FechaDB

End Sub

'--------------------------------------------------------------------------------
' Project    :       SHB
' Procedure  :       Rotina_012_Novo_Com_Anterior
' Description:       Select
' Created by :       Project Administrator
' Machine    :       DESKTOP-IR48S83
' Date-Time  :       14/02/2021-21:08:15
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub Rotina_012_Novo_Com_Anterior()

If chaveatualiza = 1 Then
   chaveatualiza = 0
   
   PessoaAnterior = cmbTabPreco
   fim = 0
   Call Rotina_15_Reposiciona_Ordem_Atualizacao
End If

Call Rotina_AbrirBanco

ProdPreco.Open "Select * from ProdutoPreco where chPessoa = ('" & PessoaAnterior & "') and chproduto = ('" & cmbProduto & "') and chAtividade =('" & cmbAtividade & "') and pdpStatus = ('" & 0 & "')", db, 3, 3
If ProdPreco.EOF Then
   MsgBox ("Erro acesso ProdutooPreço em Rotina 12 Novo com Anterior."), vbCritical
   Call FechaDB
   Exit Sub
End If

StatusAtual = ProdPreco!pdpstatus

db.BeginTrans


ProdPreco!pdpstatus = 1

ProdPreco!pdpDataFim = Date

ProdPreco.Update

db.CommitTrans

db.BeginTrans

ProdPreco.AddNew
ProdPreco!chPessoa = cmbTabPreco
ProdPreco!chProduto = cmbProduto
ProdPreco!chAtividade = cmbAtividade

If txtDesconto = "" Then
   ProdPreco!pdpDesconto = 0
Else
   ProdPreco!pdpDesconto = txtDesconto
End If

ProdPreco!pdpstatus = 0
ProdPreco!pdpDataInicio = Date
If cmbTipoAjuste.ListIndex = 2 Then
   ProdPreco!pdpPrecoDoProduto = Format$((txtAjuste), "#0.00")
Else
   ProdPreco!pdpPrecoDoProduto = Format$(((txtAjuste * ProdPreco!pdpPrecoDoProduto) / 100) + ProdPreco!pdpPrecoDoProduto, "#0.00")
End If

ProdPreco.Update
txtPrecoAtual = ProdPreco!pdpPrecoDoProduto
db.CommitTrans

Call FechaDB

End Sub

Public Sub Rotina_014_Gera_Novo_Preco()  'Qdo não existir preço anterior
db.BeginTrans

ProdPreco!chPessoa = cmbTabPreco
ProdPreco!chProduto = cmbProduto
ProdPreco!chAtividade = cmbAtividade
ProdPreco!pdpDesconto = txtDesconto
ProdPreco!pdpstatus = 0
ProdPreco!pdpDataInicio = Date

ProdPreco!pdpPrecoDoProduto = txtAjuste

ProdPreco.Update

txtPrecoAtual = txtAjuste

db.CommitTrans
End Sub


Public Sub Rotina_15_Reposiciona_Ordem_Atualizacao()

'Do While fim = 0

'   tabProdutoPreco.MoveNext
'   If tabProdutoPreco.EOF Then
'      fim = 1
'   Else
'      StatusAtual = tabProdutoPreco("pdpstatus")
'      If (tabProdutoPreco("chpessoa") = PessoaAnterior) Then
'         fim = 0
'      Else
'         fim = 1
'      End If
'   End If
'Loop

fim = 0

Call Rotina_AbrirBanco
   
ProdPreco.Open "Select * from ProdutoPreco where chPessoa = ('" & PessoaAnterior & "') and chproduto = ('" & cmbProduto & "') and chAtividade = ('" & cmbAtividade & "')", db, 3, 3
If ProdPreco.EOF Then
   MsgBox ("Erro acesso tabprodutopreco"), vbCritical
   Call FechaDB
   Exit Sub
End If

ProdPreco.MoveLast
'ProdPreco.MovePrevious

Do While fim = 0

   If ProdPreco!chPessoa = cmbTabPreco And ProdPreco!pdpstatus > 0 Then

      ProdPreco!pdpstatus = ProdPreco!pdpstatus + 1
      ProdPreco!pdpDataFim = Date
      ProdPreco.Update
 '    db.CommitTrans
      
      ProdPreco.MovePrevious
      If ProdPreco.BOF Then
         fim = 1
      End If
    Else
      fim = 1
    End If
Loop
  
Call FechaDB
  
End Sub


