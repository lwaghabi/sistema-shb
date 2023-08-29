VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAjusteComissao 
   Caption         =   "frmAjusteComissao"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form3"
   ScaleHeight     =   8115
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtHoje 
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   8415
      Begin VB.CommandButton cmdRecalcula 
         BackColor       =   &H0080FF80&
         Caption         =   "Recalcula Comissão"
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdTrenfereComis 
         BackColor       =   &H00FFFF80&
         Caption         =   "Transferir Comissão"
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelaComis 
         BackColor       =   &H008080FF&
         Caption         =   "Cancelar Comissão"
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
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00C0E0FF&
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3015
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid GridDestino 
         Height          =   2175
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         FormatString    =   "Faturamento |Nota Fiscal |Pedido  |Comp  |Cliente                        |Valor Neg.       |Comissão     |%Base|"
      End
      Begin VB.CommandButton cmdCancelaProc 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cancela  Procedimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtNovoPerc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtTotalValorDest 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTotalComisDest 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   32
         Top             =   480
         Width           =   850
      End
      Begin VB.ComboBox cmbColaboradorDest 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label16 
         Caption         =   "Ajuste Comis."
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
         TabIndex        =   39
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
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
         Left            =   5400
         TabIndex        =   34
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Transf. P/Colaborador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   31
         Top             =   120
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   8415
      Begin MSMask.MaskEdBox txtFaturamento 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "%"
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
         Left            =   7320
         TabIndex        =   36
         Top             =   120
         Width           =   495
      End
      Begin VB.Label txtPercComis 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7200
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label txtComissao 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6360
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.Label txtValor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5400
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.Label txtCliente 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3480
         TabIndex        =   27
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label txtComplemento 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2880
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.Label txtPedido 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2160
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.Label txtNotaFiscal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Comissão"
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
         Left            =   6360
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Valor Neg."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   18
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label9 
         Caption         =   "Comp"
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
         Left            =   2880
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "N.Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Faturamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid GridOrigem 
         Height          =   2055
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         FormatString    =   "Faturamento |Nota Fiscal |Pedido  |Comp  |Cliente                        |Valor Neg.       |Comissão     |%Base|"
      End
      Begin VB.TextBox txtTotalComis 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtTotalNeg 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   22
         Top             =   600
         Width           =   945
      End
      Begin VB.CommandButton cmdCarrega 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Carrega Negociações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbTipoColaborador 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbColaboradorOrigem 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
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
         Left            =   5400
         TabIndex        =   21
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Colaborador"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Colaborador"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   20
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ajuste de Comissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmAjusteComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Ind As Single
Dim IndGrid As Single
Dim ValorNegociado As Currency
Dim ValorComissao As Currency
Dim TotalNegociado As Currency
Dim TotalComissao As Currency
Dim ComissaoAtual As Currency
Dim AcumComisAnter As Currency
Dim AcumComisAtu As Currency
Dim DiaNeg As Integer
Dim MesNeg As Integer
Dim AnoNeg As Integer
Dim DataFinanc As Date
Dim Resp As String
Dim fim As Byte
Dim NaoEncontrei As Byte
Dim NovoPercComis As Single
Dim Pedido As String
Dim PedidoComp As String
Dim Produto As String
Dim Qtd As Currency


Private Sub cmbColaboradorDest_lostfocus()

Call Rotina_012_LimpaGridDestino

Call Rotina_045_Processa_Destino

End Sub

Private Sub cmbTipoColaborador_LostFocus()

Call Rotina_014_LimpaTela

If cmbTipoColaborador = Empty Then
   cmdSair.SetFocus
Else
   If cmbTipoColaborador.ListIndex = 0 Then
      TabCarteira_Rep.MoveFirst
      Do While Not TabCarteira_Rep.EOF
         cmbColaboradorOrigem.AddItem TabCarteira_Rep("chpessoa")
         cmbColaboradorDest.AddItem TabCarteira_Rep("chpessoa")
         TabCarteira_Rep.MoveNext
      Loop
   Else
      TabCarteira_Promot.MoveFirst
      Do While Not TabCarteira_Promot.EOF
         cmbColaboradorOrigem.AddItem TabCarteira_Promot("chpessoa")
         cmbColaboradorDest.AddItem TabCarteira_Promot("chpessoa")
         TabCarteira_Promot.MoveNext
      Loop
   End If
End If
End Sub

Private Sub cmdCancelaComis_Click()
Resp = MsgBox("Cancelamento de Comissão. Confirma???", vbYesNo)
If Resp = vbNo Then
   MsgBox ("Abortado Procedimento de cancelamento de comissão")
   Exit Sub
Else
   If txtNotaFiscal = Empty Then
      MsgBox ("Cancelamento inválido")
      Exit Sub
    End If
End If

BeginTrans

TabNegociacao.Seek "=", txtPedido, txtComplemento
If TabNegociacao.NoMatch Then
   MsgBox ("Erro no sistema")
   IndGrid = 1 / 0
End If

TabNegociacao.Edit

If cmbTipoColaborador.ListIndex = 0 Then
   TabNegociacao("chRepresentante") = "FABRICA"
Else
   TabNegociacao("chpromotor") = "NENHUM"
End If

MesNeg = Month(TabNegociacao("negdatanegociação"))
MesNeg = MesNeg + 1
AnoNeg = Year(TabNegociacao("negdatanegociação"))
If MesNeg = 13 Then
   MesNeg = 1
   AnoNeg = AnoNeg + 1
End If

If cmbTipoColaborador.ListIndex = 0 Then
   DataFinanc = 25 & "/" & MesNeg & "/" & AnoNeg
Else
   DataFinanc = "05" & "/" & MesNeg & "/" & AnoNeg
End If

'Data1.Recordset.FindFirst "chnumpedido='" & TabNegociacao("chNumpedido") & "' and chnumpedidocomp='" & TabNegociacao("chNumpedidocomp") & "'"
'Pedido = Data1.Recordset.Fields("chnumpedido")
'PedidoComp = Data1.Recordset.Fields("chnumpedidocomp")
'Produto = Data1.Recordset.Fields("chproduto")

TabDetalheNegociacao.Seek "=", Pedido, PedidoComp, Produto
If TabDetalheNegociacao.NoMatch Then
   MsgBox ("Tabdetalhenegociacao nao encontrado"), , TabDetNeg("chnumpedido") & TabDetNeg("chnumpedidocomp") & TabDetNeg("chproduto")
   cmdSair.SetFocus
   Exit Sub
End If
fim = 0
Do While fim = 0
      If TabDetalheNegociacao("chNumpedido") > TabNegociacao("chNumpedido") Then
         fim = 1
      Else
            If TabDetalheNegociacao("chNumpedidocomp") > TabNegociacao("chNumpedidocomp") Then
               fim = 1
            Else
               TabDetalheNegociacao.Edit
               TabDetalheNegociacao("pedcomissaorep") = 0
               TabDetalheNegociacao.Update
               TabDetalheNegociacao.MoveNext
               If TabDetalheNegociacao.EOF Then
                  fim = 1
               End If
            End If
      End If
Loop

TabNegociacao.Update

TabCtaPagar.Seek "=", 0, cmbColaboradorOrigem, cmbTipoColaborador, "Comissão", DataFinanc

If TabCtaPagar.NoMatch Then
   MsgBox ("Deu caquinha")
   IndGrid = 1 / 0
End If

TabCtaPagar.Edit

TabCtaPagar("ctpvalorlart") = TabCtaPagar("ctpvalorlart") - txtComissao

TabCtaPagar("ctpvalordaboleta") = TabCtaPagar("ctpvalordaboleta") - txtComissao

If TabCtaPagar("ctpvalordaboleta") = 0 And TabCtaPagar("ctpvalorlart") = 0 Then
   TabCtaPagar.Delete
Else
   TabCtaPagar.Update
End If

CommitTrans

Call Rotina_010_Limpa_GridOrigem

Call Rotina_012_LimpaGridDestino

Call Rotina_016_Limpa_Evidencia

Call Rotina_040_Processa_Carga

End Sub

Private Sub cmdCancelaProc_Click()
Rotina_016_Limpa_Evidencia
End Sub

Private Sub cmdCarrega_Click()

Call Rotina_010_Limpa_GridOrigem

Call Rotina_012_LimpaGridDestino

Call Rotina_016_Limpa_Evidencia

Call Rotina_040_Processa_Carga

End Sub

Private Sub cmdRecalcula_Click()
If txtNovoPerc = Empty Then
   MsgBox ("Novo desconto de percentual não Informado")
   cmdSair.SetFocus
   Exit Sub
End If

BeginTrans

TabNegociacao.Seek "=", txtPedido, txtComplemento
If TabNegociacao.NoMatch Then
   MsgBox ("Erro no sistema")
   IndGrid = 1 / 0
End If

TabNegociacao.Edit

Call Rotina_050_Recalcula_DetNeg

Call Rotina_060_Financeiro

If cmbTipoColaborador.ListIndex = 0 Then
   TabNegociacao("negdesccomissao") = txtNovoPerc
Else
   TabNegociacao("negdesccomispromot") = txtNovoPerc
End If

TabNegociacao.Update

CommitTrans

cmbColaboradorDest = cmbColaboradorOrigem

Call Rotina_045_Processa_Destino

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdTrenfereComis_Click()

If cmbColaboradorOrigem = cmbColaboradorDest Then
   MsgBox ("Transferência Inválida. Origem igual a Destino")
   cmdSair.SetFocus
   Exit Sub
End If

If cmbColaboradorDest = Empty Then
   MsgBox ("Solicitação de transferência sem informar o colaborador destinatário")
   cmdSair.SetFocus
   Exit Sub
End If
Resp = MsgBox("Transferência de Comissão. Confirma???", vbYesNo)
If Resp = vbNo Then
   MsgBox ("Abortado Procedimento de Transferência de comissão")
   Exit Sub
Else
   If txtNotaFiscal = Empty Then
      MsgBox ("Informações para Transferência não informadas. ")
      Exit Sub
    End If
End If

BeginTrans

TabNegociacao.Seek "=", txtPedido, txtComplemento
If TabNegociacao.NoMatch Then
   MsgBox ("Erro no sistema")
   IndGrid = 1 / 0
End If

TabNegociacao.Edit

If cmbTipoColaborador.ListIndex = 0 Then
   TabNegociacao("chRepresentante") = cmbColaboradorDest
Else
   TabNegociacao("chpromotor") = cmbColaboradorDest
End If

MesNeg = Month(TabNegociacao("negdatanegociação"))
MesNeg = MesNeg + 1
AnoNeg = Year(TabNegociacao("negdatanegociação"))
If MesNeg = 13 Then
   MesNeg = 1
   AnoNeg = AnoNeg + 1
End If
If cmbTipoColaborador.ListIndex = 0 Then
   DataFinanc = 25 & "/" & MesNeg & "/" & AnoNeg
Else
   DataFinanc = "05" & "/" & MesNeg & "/" & AnoNeg
End If

TabNegociacao.Update

TabCtaPagar.Seek "=", 0, cmbColaboradorOrigem, cmbTipoColaborador, "Comissão", DataFinanc

If Not TabCtaPagar.NoMatch Then

   TabCtaPagar.Edit

   TabCtaPagar("ctpvalorlart") = TabCtaPagar("ctpvalorlart") - txtComissao
   If TabCtaPagar("ctpvalordaboleta") = 0 Then
      TabCtaPagar.Delete
   Else
      TabCtaPagar.Update
   End If
End If

TabCtaPagar.Seek "=", 0, cmbColaboradorDest, cmbTipoColaborador, "Comissão", DataFinanc

If TabCtaPagar.NoMatch Then
   TabCtaPagar.AddNew
   TabCtaPagar("chfabricante") = 0
   TabCtaPagar("chpessoa") = cmbColaboradorDest
   'TabCtaPagar("chnotafiscal") = "Representante"
   TabCtaPagar("chnotafiscal") = cmbTipoColaborador
   TabCtaPagar("chfatura") = "Comissão"
   TabCtaPagar("chdatavencito") = DataFinanc
   TabCtaPagar("ctpdatavencOriginal") = DataFinanc
   TabCtaPagar("ctpdataemissao") = Date
   TabCtaPagar("ctpdatalanc") = Date
   TabCtaPagar("ctpDescricaoOperacao") = "Comissão por Transf."
   TabCtaPagar("ctpvalorlart") = txtComissao
   TabCtaPagar("ctpvalormerco") = 0
   TabCtaPagar("ctpvalordaboleta") = txtComissao
   TabCtaPagar("chano") = AnoNeg
   TabCtaPagar("chmes") = MesNeg
   TabCtaPagar("chdia") = Day(DataFinanc)
   TabCtaPagar("chcodbcolart") = "UNIBANCO"
   TabCtaPagar("ctpstatus") = 0
   TabCtaPagar("ctptipolancamento") = 5
Else
   TabCtaPagar.Edit
   TabCtaPagar("ctpvalorlart") = TabCtaPagar("ctpvalorlart") + txtComissao
   TabCtaPagar("ctpvalordaboleta") = TabCtaPagar("ctpvalordaboleta") + txtComissao
End If

TabCtaPagar.Update

CommitTrans

Call Rotina_010_Limpa_GridOrigem

Call Rotina_012_LimpaGridDestino

Call Rotina_016_Limpa_Evidencia

Call Rotina_040_Processa_Carga

Call Rotina_045_Processa_Destino

End Sub

Private Sub Form_Load()
txtHoje = Date

cmbTipoColaborador.AddItem "Representante"
cmbTipoColaborador.AddItem "Promotora"

Call Rotina_010_Limpa_GridOrigem

Call Rotina_012_LimpaGridDestino

Call Rotina_014_LimpaTela

Call Rotina_016_Limpa_Evidencia

End Sub

Public Sub Rotina_010_Limpa_GridOrigem()

GridOrigem.Rows = 2
Ind = 1
GridOrigem.TextMatrix(Ind, 0) = Empty
GridOrigem.TextMatrix(Ind, 1) = Empty
GridOrigem.TextMatrix(Ind, 2) = Empty
GridOrigem.TextMatrix(Ind, 3) = Empty
GridOrigem.TextMatrix(Ind, 4) = Empty
GridOrigem.TextMatrix(Ind, 5) = Empty
GridOrigem.TextMatrix(Ind, 6) = Empty
GridOrigem.TextMatrix(Ind, 7) = Empty
GridOrigem.TextMatrix(Ind, 8) = Empty

End Sub
Public Sub Rotina_012_LimpaGridDestino()

GridDestino.Rows = 2
Ind = 1
GridDestino.TextMatrix(Ind, 0) = Empty
GridDestino.TextMatrix(Ind, 1) = Empty
GridDestino.TextMatrix(Ind, 2) = Empty
GridDestino.TextMatrix(Ind, 3) = Empty
GridDestino.TextMatrix(Ind, 4) = Empty
GridDestino.TextMatrix(Ind, 5) = Empty
GridDestino.TextMatrix(Ind, 6) = Empty
GridDestino.TextMatrix(Ind, 7) = Empty
GridDestino.TextMatrix(Ind, 8) = Empty

End Sub

Public Sub Rotina_014_LimpaTela()
cmbColaboradorOrigem.Clear
cmbColaboradorDest.Clear
End Sub

Public Sub Rotina_016_Limpa_Evidencia()
txtFaturamento = "__/__/____"
txtNotaFiscal = Empty
txtPedido = Empty
txtComplemento = Empty
txtCliente = Empty
txtValor = Empty
txtComissao = Empty
End Sub

Public Sub Rotina_020_Acumula_Comissao()

fim = 0
ValorNegociado = 0
ValorComissao = 0

'Data1.Recordset.FindFirst "chnumpedido='" & TabNegociacao("chNumpedido") & "' and chnumpedidocomp='" & TabNegociacao("chNumpedidocomp") & "'"
'Pedido = Data1.Recordset.Fields("chnumpedido")
'PedidoComp = Data1.Recordset.Fields("chnumpedidocomp")
'Produto = Data1.Recordset.Fields("chproduto")

TabDetalheNegociacao.Seek "=", TabNegociacao("chNumpedido"), TabNegociacao("chNumpedidocomp")
If TabDetalheNegociacao.NoMatch Then
   MsgBox ("Tabdetalhenegociacao nao encontrado"), , TabDetNeg("chnumpedido") & TabDetNeg("chnumpedidocomp") & TabDetNeg("chproduto")
   cmdSair.SetFocus
   Exit Sub
End If

Do While fim = 0
      If TabDetalheNegociacao("chNumpedido") > TabNegociacao("chNumpedido") Then
         fim = 1
      Else
            If TabDetalheNegociacao("chNumpedidocomp") > TabNegociacao("chNumpedidocomp") Then
               fim = 1
            Else
               ValorNegociado = ValorNegociado + (TabDetalheNegociacao("pedvalordaoperacao")) '- TabDetalheNegociacao("pedvalorDesconto")) * TabDetalheNegociacao("pedQuantidadeMetro")
               If cmbTipoColaborador.ListIndex = 0 Then
                  ValorComissao = ValorComissao + TabDetalheNegociacao("pedcomissaorep")
               Else
                  ValorComissao = ValorComissao + TabDetalheNegociacao("pedcomissaopromot")
               End If
               TabDetalheNegociacao.MoveNext
               If TabDetalheNegociacao.EOF Then
                  fim = 1
               End If
            End If
      End If
Loop
End Sub

Public Sub Rotina_030_GridOrigem()

DiaNeg = Day(TabNegociacao("negdatanegociação"))
MesNeg = Month(TabNegociacao("negdatanegociação"))
AnoNeg = Year(TabNegociacao("negdatanegociação"))

GridOrigem.Rows = IndGrid + 1
GridOrigem.TextMatrix(IndGrid, 0) = TabNegociacao("negdatanegociação")
GridOrigem.TextMatrix(IndGrid, 1) = TabNegociacao("negNotafiscal")
GridOrigem.TextMatrix(IndGrid, 2) = TabNegociacao("chnumpedido")
GridOrigem.TextMatrix(IndGrid, 3) = TabNegociacao("chnumpedidocomp")
GridOrigem.TextMatrix(IndGrid, 4) = TabNegociacao("chpessoa")
GridOrigem.TextMatrix(IndGrid, 5) = Format$(ValorNegociado, "##,##0.00")
GridOrigem.TextMatrix(IndGrid, 6) = Format$(ValorComissao, "##,##0.00")
If cmbTipoColaborador.ListIndex = 0 Then
    If TabNegociacao("negdesccomissao") = 0 Then
       GridOrigem.TextMatrix(IndGrid, 7) = Format$(TabNegociacao("negdesccomissao"), "#0.00" & "%")
    Else
       GridOrigem.TextMatrix(IndGrid, 7) = Format$((TabNegociacao("negdesccomissao") * -1) / 100, "#0.00" & "%")
    End If
Else
    If TabNegociacao("negdesccomispromot") = 0 Then
       GridOrigem.TextMatrix(IndGrid, 7) = Format$(TabNegociacao("negdesccomispromot"), "#0.00" & "%")
    Else
       GridOrigem.TextMatrix(IndGrid, 7) = Format$((TabNegociacao("negdesccomispromot") * -1) / 100, "#0.00" & "%")
    End If
End If
GridOrigem.TextMatrix(IndGrid, 8) = AnoNeg & MesNeg & DiaNeg & TabNegociacao("negnotafiscal")
End Sub
Public Sub Rotina_035_GridDestino()

DiaNeg = Day(TabNegociacao("negdatanegociação"))
MesNeg = Month(TabNegociacao("negdatanegociação"))
AnoNeg = Year(TabNegociacao("negdatanegociação"))

GridDestino.Rows = IndGrid + 1
GridDestino.TextMatrix(IndGrid, 0) = TabNegociacao("negdatanegociação")
GridDestino.TextMatrix(IndGrid, 1) = TabNegociacao("negNotafiscal")
GridDestino.TextMatrix(IndGrid, 2) = TabNegociacao("chnumpedido")
GridDestino.TextMatrix(IndGrid, 3) = TabNegociacao("chnumpedidocomp")
GridDestino.TextMatrix(IndGrid, 4) = TabNegociacao("chpessoa")
GridDestino.TextMatrix(IndGrid, 5) = Format$(ValorNegociado, "##,##0.00")
GridDestino.TextMatrix(IndGrid, 6) = Format$(ValorComissao, "##,##0.00")
If cmbTipoColaborador.ListIndex = 0 Then
    If TabNegociacao("negdesccomissao") = 0 Then
       GridDestino.TextMatrix(IndGrid, 7) = Format$(TabNegociacao("negdesccomissao"), "#0.00" & "%")
    Else
       GridDestino.TextMatrix(IndGrid, 7) = Format$((TabNegociacao("negdesccomissao") * -1) / 100, "#0.00" & "%")
    End If
Else
    If TabNegociacao("negdesccomispromot") = 0 Then
       GridDestino.TextMatrix(IndGrid, 7) = Format$(TabNegociacao("negdesccomispromot"), "#0.00" & "%")
    Else
       GridDestino.TextMatrix(IndGrid, 7) = Format$((TabNegociacao("negdesccomispromot") * -1) / 100, "#0.00" & "%")
    End If
End If
GridDestino.TextMatrix(IndGrid, 8) = AnoNeg & MesNeg & DiaNeg & TabNegociacao("negnotafiscal")
End Sub

Public Sub Rotina_040_Processa_Carga()

TotalNegociado = 0
TotalComissao = 0
IndGrid = 0

TabNegociacao.MoveFirst
Do While Not TabNegociacao.EOF
   If TabNegociacao("negstatus") > 0 And TabNegociacao("negcondprocess") <> 4 Then
      If TabNegociacao("chRepresentante") = cmbColaboradorOrigem Or TabNegociacao("chPromotor") = cmbColaboradorOrigem Then
         Call Rotina_020_Acumula_Comissao
         IndGrid = IndGrid + 1
         Call Rotina_030_GridOrigem
         TotalComissao = TotalComissao + ValorComissao
         TotalNegociado = TotalNegociado + ValorNegociado
      End If
   End If
   TabNegociacao.MoveNext
Loop

txtTotalNeg = Format$(TotalNegociado, "##,##0.00")
txtTotalComis = Format$(TotalComissao, "##,##0.00")

TotalNegociado = 0
TotalComissao = 0

If IndGrid > 1 Then
   GridOrigem.Row = 1
   GridOrigem.Col = 0
   GridOrigem.RowSel = IndGrid
   GridOrigem.ColSel = 8

   GridOrigem.Sort = 1

   GridOrigem.Row = IndGrid
   GridOrigem.Col = 8
   GridOrigem.RowSel = IndGrid
   GridOrigem.ColSel = 8
End If

End Sub
Public Sub Rotina_045_Processa_Destino()
TotalNegociado = 0
TotalComissao = 0
IndGrid = 0

TabNegociacao.MoveFirst
Do While Not TabNegociacao.EOF
   If TabNegociacao("negstatus") = 1 Then
      If TabNegociacao("chRepresentante") = cmbColaboradorDest Or TabNegociacao("chPromotor") = cmbColaboradorDest Then
         Call Rotina_020_Acumula_Comissao
         IndGrid = IndGrid + 1
         Call Rotina_035_GridDestino
         TotalComissao = TotalComissao + ValorComissao
         TotalNegociado = TotalNegociado + ValorNegociado
      End If
   End If
   TabNegociacao.MoveNext
Loop

txtTotalValorDest = Format$(TotalNegociado, "##,##0.00")
txtTotalComisDest = Format$(TotalComissao, "##,##0.00")

TotalNegociado = 0
TotalComissao = 0

If IndGrid > 1 Then
   GridDestino.Row = 1
   GridDestino.Col = 0
   GridDestino.RowSel = IndGrid
   GridDestino.ColSel = 8

   GridDestino.Sort = 1

   GridDestino.Row = IndGrid
   GridDestino.Col = 8
   GridDestino.RowSel = IndGrid
   GridDestino.ColSel = 8
End If
End Sub

Private Sub GridOrigem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

IndGrid = GridOrigem.Row
If GridOrigem.TextMatrix(IndGrid, 0) = Empty Then
   MsgBox ("Para Procedimentos Especiais, clicar em linha com conteúdo")
   Exit Sub
End If

txtFaturamento = GridOrigem.TextMatrix(IndGrid, 0)
txtNotaFiscal = GridOrigem.TextMatrix(IndGrid, 1)
txtPedido = GridOrigem.TextMatrix(IndGrid, 2)
txtComplemento = GridOrigem.TextMatrix(IndGrid, 3)
txtCliente = GridOrigem.TextMatrix(IndGrid, 4)
txtValor = GridOrigem.TextMatrix(IndGrid, 5)
txtComissao = GridOrigem.TextMatrix(IndGrid, 6)
txtPercComis = GridOrigem.TextMatrix(IndGrid, 7)
End Sub

Public Sub Rotina_050_Recalcula_DetNeg()

NaoEncontrei = 0

'TabDetNeg.FindFirst "chnumpedido = txtPedido and chnumpedidocomp = txtComplemento"

'Data1.Recordset.FindFirst "chnumpedido='" & txtPedido & "' and chnumpedidocomp='" & txtComplemento & "'"
'Pedido = Data1.Recordset.Fields("chnumpedido")
'PedidoComp = Data1.Recordset.Fields("chnumpedidocomp")
'Produto = Data1.Recordset.Fields("chproduto")

TabDetalheNegociacao.Seek "=", Pedido, PedidoComp, Produto
If TabDetalheNegociacao.NoMatch Then
   MsgBox ("Tabdetalhenegociacao nao encontrado"), , TabDetNeg("chnumpedido") & TabDetNeg("chnumpedidocomp") & TabDetNeg("chproduto")
   cmdSair.SetFocus
   Exit Sub
End If

fim = 0
Do While fim = 0
   If TabDetalheNegociacao("chnumpedido") = txtPedido And TabDetalheNegociacao("chnumpedidocomp") = txtComplemento Then
      fim = 1
   Else
      TabDetalheNegociacao.MoveNext
      If TabDetalheNegociacao.EOF Then
         fim = 1
         NaoEncontrei = 1
      End If
   End If
Loop

If NaoEncontrei = 1 Then
   MsgBox ("Detalhe não encontrado sequencialmente")
   cmdSair.SetFocus
   Exit Sub
End If

fim = 0
AcumComisAnter = 0
AcumComisAtu = 0

Do While fim = 0

    TabDetalheNegociacao.Edit
    
    tabproduto.Seek "=", TabDetalheNegociacao("chproduto")
    If tabproduto.NoMatch Then
       MsgBox ("Erro no acesso a Produtos")
       cmdSair.SetFocus
       Exit Sub
    End If
    
    If TabDetalheNegociacao("pedquantidademetro") = 0 Then
       Qtd = 1
    Else
       Qtd = TabDetalheNegociacao("pedquantidademetro")
    End If
    
    If cmbTipoColaborador.ListIndex = 0 Then
       AcumComisAnter = AcumComisAnter + TabDetalheNegociacao("pedcomissaorep")
       ComissaoAtual = (TabDetalheNegociacao("pedprecometro") - TabDetalheNegociacao("pedvalordesconto")) * Qtd
       TabDetalheNegociacao("pedComissaoRep") = Format$(ComissaoAtual * ((tabproduto("prdComissao") - txtNovoPerc) / 100), "#.00")
       AcumComisAtu = AcumComisAtu + TabDetalheNegociacao("pedComissaoRep")
    Else
       Tabpessoa.Seek "=", TabNegociacao("chpessoa")
       TabCarteira_Promot.Seek "=", Tabpessoa("chcarteirapromot")
       AcumComisAnter = AcumComisAnter + TabDetalheNegociacao("pedcomissaopromot")
       ComissaoAtual = (TabDetalheNegociacao("pedprecometro") - TabDetalheNegociacao("pedvalordesconto")) * Qtd
       TabDetalheNegociacao("pedComissaoPromot") = Format$(ComissaoAtual * ((TabCarteira_Promot("proComissaopromot") - txtNovoPerc) / 100), "#.00")
       AcumComisAtu = AcumComisAtu + TabDetalheNegociacao("pedComissaoPromot")
    End If
    
    TabDetalheNegociacao.Update
    
    TabDetalheNegociacao.MoveNext
    
    If TabDetalheNegociacao.EOF Then
       fim = 1
    Else
       If TabDetalheNegociacao("chnumpedido") > txtPedido Then
          fim = 1
       Else
          If TabDetalheNegociacao("chnumpedidocomp") > txtComplemento Then
             fim = 1
          End If
       End If
    End If
Loop

End Sub

Public Sub Rotina_060_Financeiro()

If cmbTipoColaborador.ListIndex = 0 Then
   DiaNeg = 25
Else
   DiaNeg = 5
End If

MesNeg = Month(Date) + 1

If MesNeg = 13 Then
   MesNeg = 1
   AnoNeg = Year(Date) + 1
Else
   AnoNeg = Year(Date)
End If

DataFinanc = DiaNeg & "/" & MesNeg & "/" & AnoNeg

TabCtaPagar.Seek "=", 0, cmbColaboradorOrigem, cmbTipoColaborador, "Comissão", DataFinanc

If TabCtaPagar.NoMatch Then
   TabCtaPagar.AddNew
   TabCtaPagar("chfabricante") = 0
   TabCtaPagar("chpessoa") = cmbColaboradorOrigem
   TabCtaPagar("chnotafiscal") = cmbTipoColaborador
   TabCtaPagar("chfatura") = "Comissão"
   TabCtaPagar("chdatavencito") = DataFinanc
   TabCtaPagar("ctpdatavencoriginal") = DataFinanc
   TabCtaPagar("ctpdataemissao") = Date
   TabCtaPagar("ctpdatalanc") = Date
   TabCtaPagar("ctpDescricaoOperacao") = "Comis Origem Aut."
   TabCtaPagar("ctpvalorlart") = AcumComisAtu
   TabCtaPagar("ctpvalormerco") = 0
   TabCtaPagar("ctpvalordaboleta") = AcumComisAtu
   TabCtaPagar("chano") = AnoNeg
   TabCtaPagar("chmes") = MesNeg
   TabCtaPagar("chdia") = Day(DataFinanc)
   TabCtaPagar("chcodbcolart") = "UNIBANCO"
   TabCtaPagar("ctpstatus") = 0
   TabCtaPagar("ctptipolancamento") = 5
Else
   TabCtaPagar.Edit
   TabCtaPagar("ctpvalorlart") = (TabCtaPagar("ctpvalorlart") + AcumComisAtu) - AcumComisAnter
   TabCtaPagar("ctpvalordaboleta") = (TabCtaPagar("ctpvalordaboleta") + AcumComisAtu) - AcumComisAnter
End If

TabCtaPagar.Update
End Sub
