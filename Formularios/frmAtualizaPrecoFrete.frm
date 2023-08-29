VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAtualizaPrecoFrete 
   Caption         =   "frmAtualizaPrecoFrete"
   ClientHeight    =   5700
   ClientLeft      =   2205
   ClientTop       =   2040
   ClientWidth     =   8430
   LinkTopic       =   "Form3"
   ScaleHeight     =   5700
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   8175
      Begin VB.ComboBox cmbTipoTabPreco 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cmbTipoFrete 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cmbTipoAjuste 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
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
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox cmbTabPreco 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   7935
         Begin VB.Frame Frame3 
            Caption         =   "Navega"
            Height          =   615
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3015
            Begin VB.CommandButton cmbNavega 
               Caption         =   "Início"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmbNavega 
               Caption         =   "Prox."
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   17
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmbNavega 
               Caption         =   "Anter."
               Height          =   255
               Index           =   2
               Left            =   1560
               TabIndex        =   16
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmbNavega 
               Caption         =   "Último"
               Height          =   255
               Index           =   3
               Left            =   2280
               TabIndex        =   15
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Processa"
            Height          =   615
            Left            =   3240
            TabIndex        =   11
            Top             =   240
            Width           =   3255
            Begin VB.CommandButton Command1 
               Caption         =   "Confirma"
               Height          =   255
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   855
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Cancela"
               Height          =   255
               Left            =   1200
               TabIndex        =   13
               Top             =   240
               Width           =   855
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Excluir"
               Height          =   255
               Left            =   2280
               TabIndex        =   12
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   6600
            TabIndex        =   10
            Top             =   240
            Width           =   1215
            Begin VB.CommandButton cmdSair 
               Caption         =   "Sair"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   855
            End
         End
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Tab de Preços"
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
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo do Frete"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Forma de Ajuste"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ajuste ou Pco Atual"
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
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tabela de Preços"
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
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pco anter. ao ajuste"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   1935
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
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Preço Corrigido"
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
         Left            =   240
         TabIndex        =   20
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label txtPrecoAtual 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   3240
         Width           =   1815
      End
   End
   Begin MSMask.MaskEdBox txtHoje 
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atualização de Preços de Frete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   29
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   28
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmAtualizaPrecoFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As String
Dim GrupoAnterior As String
Dim CidadeAnterior As String
Private Sub cmbTabPreco_LostFocus()
If cmbTabPreco = Empty Then
   MsgBox ("Esta informação é obrigatória")
   cmbTipoTabPreco.SetFocus
End If
End Sub


Private Sub cmbTipoTabPreco_LostFocus()

If txtAjuste = "999" Then
   cmbTipoAjuste.ListIndex = 2
   txtAjuste = Empty
   txtAjuste.SetFocus
Else
   If cmbTipoTabPreco.ListIndex = 0 Then
      cmbTabPreco.Clear
      cmbTabPreco.AddItem "Geral"
      cmbTabPreco.ListIndex = 0
   Else
      If cmbTipoTabPreco.ListIndex = 1 Then
         cmbTabPreco.Clear
         cmbTabPreco.AddItem " Geral"
         TabICMS.MoveFirst
         Do While Not TabICMS.EOF
            cmbTabPreco.AddItem TabICMS("chuf")
            TabICMS.MoveNext
         Loop
         cmbTabPreco.ListIndex = 1
         cmbTabPreco.SetFocus
      Else
         cmbTabPreco.Clear
         GrupoAnterior = Empty
         CidadeAnterior = Empty
         Tabpessoa.MoveFirst
         Do While Not Tabpessoa.EOF
            If cmbTipoTabPreco.ListIndex = 2 Then
               If Tabpessoa("pestabprecofrete") = cmbTipoTabPreco.ListIndex Then
                  If Not Tabpessoa("pescidade") = CidadeAnterior Then
                         cmbTabPreco.AddItem Tabpessoa("pescidade")
                         GrupoAnterior = Tabpessoa("pescidade")
                         Tabpessoa.MoveNext
                  Else
                     Tabpessoa.MoveNext
                  End If
               Else
                  Tabpessoa.MoveNext
               End If
            Else
               If Tabpessoa("pestabprecofrete") = cmbTipoTabPreco.ListIndex Then
                  If Not Tabpessoa("pesgrupo") = GrupoAnterior Then
                         cmbTabPreco.AddItem Tabpessoa("chpessoa")
                         GrupoAnterior = Tabpessoa("chpessoa")
                         Tabpessoa.MoveNext
                  Else
                     Tabpessoa.MoveNext
                  End If
               Else
                  Tabpessoa.MoveNext
               End If
            End If
         Loop
         'cmbTabPreco.ListIndex = 0
      End If
   End If
End If


End Sub

Private Sub cmdSair_Click()
'frmPedido.txtFrete = txtAjuste
Unload Me
'frmPedido.txtUnidade.SetFocus
End Sub

Private Sub Form_Load()
Dim TEXTO As String
txtHoje = Date

tabTipoTabPreco.MoveFirst
Do While Not tabTipoTabPreco.EOF
   cmbTipoTabPreco.AddItem tabTipoTabPreco("ttpdescricaotipopreco")
   tabTipoTabPreco.MoveNext
Loop
cmbTipoTabPreco.ListIndex = 0
TEXTO = cmbTipoTabPreco
cmbTipoFrete.AddItem " Todos"

cmbTipoFrete.AddItem "1=Pesos Inf"
cmbTipoFrete.AddItem "2=Pesos Sup"

cmbTipoAjuste.AddItem " % Para Pesos Inf e Sup"
cmbTipoAjuste.AddItem " % para o peso indicado"
cmbTipoAjuste.AddItem " Valor informado"

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

If cmbTipoFrete.ListIndex = 0 Then
   resp = MsgBox("Correçao de Toda a Tabela de preços de frete. Confirma???", vbYesNo)
   If resp = vbYes Then
      cmbTipoAjuste.ListIndex = 0
      MsgBox ("Chamar rotina de atualizaçao geral")
   Else
      cmbTabPreco.SetFocus
   End If
   Exit Sub
Else
   tabFretePrecoAtu.Seek "=", cmbTabPreco, cmbTipoFrete.ListIndex, 0
   If tabFretePrecoAtu.NoMatch Then
      If cmbTipoAjuste.ListIndex < 2 Then
         MsgBox ("Não há frete nesta tabela para ser atualizado")
         cmbTipoTabPreco.SetFocus
      Else
         txtPrecoAnterior = Empty
         Call Rotina_010_Atualiza_Preco_Prod
      End If
   Else
      Call Rotina_010_Atualiza_Preco_Prod
   End If
End If
cmdSair.SetFocus
End Sub

Public Sub Rotina_010_Atualiza_Preco_Prod()

tabFretePrecoAtu.Seek "=", cmbTabPreco, cmbTipoFrete.ListIndex, 0
If tabFretePrecoAtu.NoMatch Then
   Call Rotina_014_Gera_Novo_Preco
Else
   Call Rotina_012_Novo_Com_Anterior
   MsgBox ("Atualização concluída")
   cmbTipoFrete.SetFocus
End If
End Sub

Public Sub Rotina_012_Novo_Com_Anterior()
BeginTrans
tabFretePreco.AddNew
tabFretePreco("chpessoa") = tabFretePrecoAtu("chpessoa")
tabFretePreco("chtipofrete") = tabFretePrecoAtu("chtipofrete")
tabFretePreco("pdfstatusfrete") = 1
tabFretePreco("pdfdatainicio") = Date
tabFretePreco("pdfdatafim") = Date
tabFretePreco("pdfprecodofrete") = tabFretePrecoAtu("pdfprecodofrete")


tabFretePreco.Update

tabFretePrecoAtu.Edit
tabFretePrecoAtu("pdfdatainicio") = Date
If cmbTipoAjuste.ListIndex = 2 Then
   tabFretePrecoAtu("pdfprecodofrete") = Format$((txtAjuste), "#0.00")
Else
   tabFretePrecoAtu("pdfprecodofrete") = Format$(((txtAjuste * tabFretePrecoAtu("pdfprecodofrete")) / 100) + tabFretePrecoAtu("pdfprecodoproduto"), "#0.00")
End If
tabFretePrecoAtu.Update
txtPrecoAtual = tabFretePrecoAtu("pdfprecodofrete")
CommitTrans
End Sub

Public Sub Rotina_014_Gera_Novo_Preco()  'Qdo não existir preço anterior
BeginTrans
tabFretePreco.AddNew
tabFretePreco("chpessoa") = cmbTabPreco
tabFretePreco("chtipofrete") = cmbTipoFrete.ListIndex
tabFretePreco("pdfstatusFRETE") = 0
tabFretePreco("pdfdatainicio") = Date

tabFretePreco("pdfprecodofrete") = txtAjuste

tabFretePreco.Update

txtPrecoAtual = txtAjuste

CommitTrans
End Sub

