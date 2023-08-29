VERSION 5.00
Begin VB.Form frmEmpenho 
   Caption         =   "Previsão de Saldo no Período - frmEmpenho"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18255
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   18255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMesExtenso 
      Height          =   375
      Left            =   15600
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   13560
      TabIndex        =   65
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmbAtualizaExtrato 
         BackColor       =   &H00FFFF00&
         Caption         =   "Atualiza Extrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1095
      Left            =   13560
      TabIndex        =   64
      Top             =   960
      Width           =   1935
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Posição sem Projeção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   240
      TabIndex        =   39
      Top             =   2160
      Width           =   7455
      Begin VB.Frame Frame6 
         Caption         =   "Saldo Calculado e Informado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   52
         Top             =   4200
         Width           =   7215
         Begin VB.Label lblSaldoSistema 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5040
            TabIndex        =   63
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label txtSaldoAtual1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5040
            TabIndex        =   62
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblSaldoProcessado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5040
            TabIndex        =   61
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "Diferença Encontrada(Sist.)"
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
            Left            =   240
            TabIndex        =   56
            Top             =   1320
            Width           =   4695
         End
         Begin VB.Label Label27 
            Caption         =   "Saldo No Banco(Bco.)"
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
            TabIndex        =   55
            Top             =   720
            Width           =   4815
         End
         Begin VB.Label Label26 
            Caption         =   "Saldo Processado(Sist.)"
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
            TabIndex        =   54
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label25 
            Caption         =   "Label25"
            Height          =   495
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pagamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   49
         Top             =   3240
         Width           =   7215
         Begin VB.TextBox lblTotalPagoNoPeriodoSist 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   5040
            TabIndex        =   51
            Top             =   340
            Width           =   1935
         End
         Begin VB.Label Label24 
            Caption         =   "Total Pago no Período(Sist.)"
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
            Left            =   120
            TabIndex        =   50
            Top             =   340
            Width           =   4815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Recebimentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   7215
         Begin VB.TextBox lblTotalCreditoSist 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   5040
            TabIndex        =   48
            Top             =   1275
            Width           =   1935
         End
         Begin VB.TextBox lblAplicBancBco 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
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
            Left            =   5040
            TabIndex        =   47
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox lblTotalRecebidoSist 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
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
            Left            =   5040
            TabIndex        =   44
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label23 
            Caption         =   "Total Recebido(Sist.)"
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
            TabIndex        =   46
            Top             =   1275
            Width           =   4815
         End
         Begin VB.Label Label22 
            Caption         =   "Aplicações Bancárias(Bco.)"
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
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   4815
         End
         Begin VB.Label Label21 
            Caption         =   "Total  Recebido no Período(Sist.)"
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
            TabIndex        =   43
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Label txtSaldoAtual0 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5160
         TabIndex        =   60
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblSaldoInicial1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5160
         TabIndex        =   59
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label29 
         Caption         =   "Label29"
         Height          =   375
         Left            =   5160
         TabIndex        =   58
         Top             =   960
         Width           =   15
      End
      Begin VB.Label Label5 
         Caption         =   "Saldo Atual no Banco(Bco.)"
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
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label20 
         Caption         =   "Saldo no Início do Período(Bco.)"
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
         TabIndex        =   40
         Top             =   960
         Width           =   4695
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15840
      TabIndex        =   36
      Top             =   960
      Width           =   2175
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H008080FF&
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
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
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
      Left            =   15840
      TabIndex        =   25
      Top             =   480
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Extrato Referente ao Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   12855
      Begin VB.TextBox txtTerminoPeriodo 
         Alignment       =   2  'Center
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
         Left            =   7560
         TabIndex        =   24
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtInicioPeriodo 
         Alignment       =   2  'Center
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
         Left            =   4680
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Término"
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
         Left            =   7560
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame txtLocacaoRealizada 
      Caption         =   "Posição com Projeção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   7800
      TabIndex        =   2
      Top             =   2160
      Width           =   10215
      Begin VB.TextBox txtSaldoGeralProjetado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   8040
         TabIndex        =   35
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox txtLocacaoNaoProcessada 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8040
         TabIndex        =   33
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label lblSaldoInicial0 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   8040
         TabIndex        =   57
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Geral Projetado"
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
         Left            =   600
         TabIndex        =   34
         Top             =   5640
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "Locações e Serviços realizados e não processados"
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
         Left            =   600
         TabIndex        =   32
         Top             =   5160
         Width           =   7335
      End
      Begin VB.Label lblTaxasBancarias 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   31
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Taxas e Débitos Bancários"
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
         Left            =   600
         TabIndex        =   30
         Top             =   3480
         Width           =   6375
      End
      Begin VB.Label lblRecebAtraso 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   29
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label lblAplicacoesBancarias 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   28
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Rendimentos de Aplicações Bancárias"
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
         Left            =   600
         TabIndex        =   27
         Top             =   1320
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "Previsão de Recebimentos até o fim do mês"
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
         Left            =   600
         TabIndex        =   26
         Top             =   960
         Width           =   6615
      End
      Begin VB.Label lblTotalDebito 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   22
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Valores Pagos no Período"
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
         Left            =   600
         TabIndex        =   21
         Top             =   2400
         Width           =   5295
      End
      Begin VB.Label lblValorePagosNoPeriodo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   20
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblValorRecebido 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   19
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Valor Recebido no Período"
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
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label lblSaldoProjetado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   17
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Saldo Projetado para o Mês"
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
         Left            =   600
         TabIndex        =   16
         Top             =   4320
         Width           =   4815
      End
      Begin VB.Label lblDebAteFimPeriodo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   15
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Total a Débito Projetado para o Mês"
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
         Left            =   600
         TabIndex        =   14
         Top             =   3840
         Width           =   5775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Débitos Previstos até o final do Mês"
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
         Left            =   600
         TabIndex        =   13
         Top             =   3120
         Width           =   4965
      End
      Begin VB.Label lblDebitosAteData 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   12
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Pagamentos em atraso até a data"
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
         Left            =   600
         TabIndex        =   11
         Top             =   2760
         Width           =   4620
      End
      Begin VB.Label lblTotalCreditoProjetado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Total a Crédito Projetado para o Mês"
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
         Left            =   600
         TabIndex        =   9
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label lblValorAteFimPeriodo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Recebimentos em atraso até a data"
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
         Left            =   600
         TabIndex        =   7
         Top             =   4800
         Width           =   4905
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Saldo no Início do Período"
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
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   3720
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hoje"
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
      Left            =   16560
      TabIndex        =   1
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Projeção Financeira"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12855
   End
End
Attribute VB_Name = "frmEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dia As Integer
Dim Mes As Integer
Dim ano As Integer
Dim dataInicio As Date
Dim DataInicioDeb As Date
Dim dataFim As Date
Dim DatafimDeb As Date
Dim DataInicioPeriodo As Date
Dim DataHojeInvertida As String

Dim PosicaoPesquisa As String

Dim AcumRecPeriodo As Currency
Dim AcumAteFimPeriodo As Currency
Dim AcumAtrasados As Currency

Dim AcumDebPago As Currency
Dim AcumDebAteFimPeriodo As Currency
Dim AcumEmpenho As Currency

Dim UltimoSaldo As Currency
Dim SaldoInicial As Currency
Dim RendimentosBancarios As Currency
Dim TarifasBancarias As Currency

Dim DataInicioInvertida As String
Dim DataFinalInvertida As String
Dim DataHoje As Date
Dim DataHojeCSV As String
Dim MesAno As String
Dim MesExtenso As String
Dim EncontreiExtrato As Integer
   
Dim caminhoArquivo As String
Dim nomeArquivo As String
Dim nomeArquivoNovo As String
Dim excelApp As Excel.Application
Dim oWB As Excel.Workbook
   

Dim AcumulaNaoProcessado As Currency

Private Sub cmbAtualizaExtrato_Click()
Call converte_csv
If EncontreiExtrato = 1 Then
   Call carrega_tabela_extrato
End If
   
Call Carga_Load

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()

Call Rotina_Tratar_Credito
Call Rotina_Tratar_Debito

Call Rotina_Tratar_Nao_Processado

End Sub

Private Sub Form_Load()

Call Carga_Load

End Sub

Public Sub Rotina_Tratar_Credito()
AcumRecPeriodo = 0
AcumAteFimPeriodo = 0
AcumAtrasados = 0


Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber", db, 3, 3
If ctr.EOF Then
   MsgBox ("Sem Contas a receber registradas."), vbInformation
   Call FechaDB
   Exit Sub
End If
Do While Not ctr.EOF
  
      If ctr!ctrStatus = 1 Then
         If ctr!ctrDataRecebimento > (dataInicio - 1) And ctr!ctrDataRecebimento < (dataFim + 1) Then
            AcumRecPeriodo = AcumRecPeriodo + Format$(ctr!ctrValorDaBoleta, "##,##0.00")
         End If
      Else
         If ctr!ctrDataBanco < Date Then
            AcumAtrasados = AcumAtrasados + Format$(ctr!ctrValorDaBoleta, "##,##0.00")
         Else
            If ctr!ctrDataBanco < (dataFim + 1) Then
               AcumAteFimPeriodo = AcumAteFimPeriodo + Format$(ctr!ctrValorDaBoleta, "##,##0.00")
            End If
         End If
      End If
 
   ctr.MoveNext
   
Loop

lblValorRecebido = Format$(AcumRecPeriodo, "##,###,##0.00")
lblTotalRecebidoSist = Format$(AcumRecPeriodo, "##,###,##0.00")
lblRecebAtraso = Format$(AcumAtrasados, "###,##0.00")
lblAplicacoesBancarias = Format$(RendimentosBancarios, "##,##0.00")
lblAplicBancBco = Format$(RendimentosBancarios, "##,##0.00")
lblValorAteFimPeriodo = Format$(AcumAteFimPeriodo, "##,###,##0.00") ' + RendimentosBancarios Este valor estava sendo adicionado
lblTotalCreditoSist = Format$(AcumRecPeriodo + RendimentosBancarios, "##,###,##0.00")
lblTotalCreditoProjetado = Format$(SaldoInicial + AcumRecPeriodo + AcumAteFimPeriodo + RendimentosBancarios, "##,###,##0.00")
If lblTotalCreditoProjetado < 0 Then
   lblTotalCreditoProjetado.ForeColor = vbRed
Else
   lblTotalCreditoProjetado.ForeColor = vbBlue
End If

Call FechaDB

End Sub

Public Sub Rotina_Tratar_Debito()
AcumDebPago = 0
AcumDebAteFimPeriodo = 0
AcumEmpenho = 0

Call Rotina_AbrirBanco

ctp.Open "Select * from Contas_A_Pagar", db, 3, 3
If ctp.EOF Then
   MsgBox ("Sem contas a pagar registradas"), vbInformation
   Call FechaDB
   Exit Sub
End If

Do While Not ctp.EOF
   
      If ctp!ctpStatus = 1 Then
         If ctp!ctpDataPagamento > DataInicioDeb And ctp!ctpDataPagamento < (dataFim + 1) Then
            AcumDebPago = AcumDebPago + Format$(ctp!ctpValorDaBoleta, "###,##0.00")
         End If
      Else
         If ctp!chDataVencito > (Date - 1) And ctp!chDataVencito < (dataFim + 1) Then
            AcumDebAteFimPeriodo = AcumDebAteFimPeriodo + Format$(ctp!ctpValorDaBoleta, "###,##0.00")
         Else
            If ctp!chDataVencito < Date And ctp!ctpStatus = 0 Then
               AcumEmpenho = AcumEmpenho + Format$(ctp!ctpValorDaBoleta, "###,##0.00")
            End If
         End If
      End If
   ctp.MoveNext
   
Loop

TarifasBancarias = TarifasBancarias * -1

lblValorePagosNoPeriodo = Format$(AcumDebPago, "##,###,##0.00")
lblTotalPagoNoPeriodoSist = Format$(AcumDebPago, "##,###,##0.00")
lblDebitosAteData = Format$(AcumEmpenho, "##,###,##0.00")
lblTaxasBancarias = Format$(TarifasBancarias, "##,###,##0.00")
lblDebAteFimPeriodo = Format$(AcumDebAteFimPeriodo, "##,###,##0.00")
lblTotalDebito = Format$(AcumDebPago + AcumEmpenho + AcumDebAteFimPeriodo + TarifasBancarias, "##,###,##0.00")
lblTotalDebito.ForeColor = vbRed
lblSaldoProjetado = Format$(lblTotalCreditoProjetado - lblTotalDebito, "##,###,##0.00")
If lblSaldoProjetado < 0 Then
   lblSaldoProjetado.ForeColor = vbRed
Else
   lblSaldoProjetado.ForeColor = vbBlue
End If

lblSaldoProcessado = Format$((SaldoInicial + AcumRecPeriodo + RendimentosBancarios) - AcumDebPago, "##,###,##0.00")

lblSaldoSistema = Format$(lblSaldoProcessado - txtSaldoAtual1, "##,###,##0.00")

If lblSaldoSistema < 0 Then
   lblSaldoSistema.ForeColor = vbRed
Else
   lblSaldoSistema.ForeColor = vbBlue
End If

End Sub


Private Sub txtSaldoAtual0_LostFocus()
If IsNumeric(txtSaldoAtual0) Then
   txtSaldoAtual0 = Format$(txtSaldoAtual0, "##,##0.00")
End If

End Sub
Public Sub GeraDataInicioDataFim()
Dim MesProximo As Integer

Mes = Format$(Month(Date), "00")
ano = Year(Date)
Dia = Format$(1, "00")

DataInicioInvertida = Format$(ano & "-" & Mes & "-" & Dia, "yyyy-mm-dd")

MesProximo = Format$(Mes, "00")
DataHoje = Date
Do While Mes = MesProximo
   DataHoje = DataHoje + 1
   MesProximo = Format$(Month(DataHoje), "00")
Loop
DataHoje = DataHoje - 1
DataFinalInvertida = Format$(DataHoje, "yyyy-mm-dd")
DataHoje = Date
dataFim = DataFinalInvertida

End Sub

Public Sub Rotina_Tratar_Nao_Processado()

Call Rotina_AbrirBanco

AcumulaNaoProcessado = 0
   
neg.Open "Select * from Negociacao", db, 3, 3
If neg.EOF Then
   txtLocacaoNaoProcessada = Format$(0, "0.00")
   txtSaldoGeralProjetado = Format$(lblSaldoProjetado, "0.00")
   Exit Sub
End If

neg.MoveFirst

Do While Not neg.EOF
   If Not neg!negStatus = 1 Then
      If neg!negFinalMedicao < Date Then
         If dneg.State = 1 Then
            dneg.Close: Set dneg = Nothing
         End If
      
         dneg.Open "Select * from DetalheNegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
         If Not dneg.EOF Then
            dneg.MoveFirst
            Do While Not dneg.EOF
               AcumulaNaoProcessado = AcumulaNaoProcessado + dneg!pedValorDaOperacao
               dneg.MoveNext
            Loop
         End If
      End If
   End If
   
   neg.MoveNext
   
Loop
      
txtLocacaoNaoProcessada = Format$(AcumulaNaoProcessado, "0,000.00")
txtSaldoGeralProjetado = Format$((lblSaldoProjetado + AcumulaNaoProcessado + AcumAtrasados), "0,000.00")
      

'txtLocacaoNaoProcessada.ForeColor = vbBlue
txtSaldoGeralProjetado.ForeColor = vbBlue

End Sub

Public Sub Carga_Load()
txtHoje = Date
txtMesExtenso = UCase$(Format$(Date, "MMMM"))

RendimentosBancarios = 0

TarifasBancarias = 0

ano = Year(Date)
Mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = Format$(ano & "-" & Mes & "-" & Dia, "yyyy-mm-dd")
Call GeraDataInicioDataFim

Call Rotina_AbrirBanco

ext.Open "Select * from Extrato", db, 3, 3
If ext.EOF Then
   MsgBox ("Extrato incorreto. A função será descontinuada"), vbCritical
   Call FechaDB
   Exit Sub
End If
   
ext.MoveFirst

Do While Not ext.EOF
   If ext!F1 = "Periodo:" Then
      txtInicioPeriodo = Mid$(ext!F2, 1, 10)
      dataInicio = Mid$(ext!F2, 1, 10)
      txtTerminoPeriodo = Mid$(ext!F2, 16, 10)
     ' DataFim = Mid$(ext!F2, 16, 10)
   End If
   
   If ext!F2 = "SALDO ANTERIOR" Then
      SaldoInicial = Format$(ext!F5, "###,##0.00")
      lblSaldoInicial1 = Format$(ext!F5, "###,##0.00")
      lblSaldoInicial0 = Format$(ext!F5, "###,##0.00")
      'txtSaldoAtual(2) = Format$(ext!F5, "###,##0.00")
   End If
   
   If Not IsNull(ext!F2) Then
      PosicaoPesquisa = Mid$(ext!F2, 1, 8)
   End If

   If (PosicaoPesquisa = "REND PAG") Or (PosicaoPesquisa = "EST TAR ") Then
      RendimentosBancarios = RendimentosBancarios + Format$(ext!F4, "###,##0.00")
   End If
   
   'If PosicaoPesquisa = "TAR CONT" Or PosicaoPesquisa = "TAR CTA " Or PosicaoPesquisa = "INT PR" Or PosicaoPesquisa = "TAR ADAP" Then
   '   TarifasBancarias = TarifasBancarias + Format$(ext!F4, "###,##0.00")
   'End If
   
   TarifasBancarias = 0
   
   If PosicaoPesquisa = "SDO CTA/" Then
      If Not IsNull(ext!F5) Then
         UltimoSaldo = Format$(ext!F5, "###,##0.00")
      End If
   End If
   
   If PosicaoPesquisa = "SALDO DO" Then
      If Not IsNull(ext!F5) Then
         UltimoSaldo = UltimoSaldo + Format$(ext!F5, "###,##0.00")
      End If
   End If
   
   ext.MoveNext
   
Loop

txtSaldoAtual0 = Format$(UltimoSaldo, "##,##0.00")
txtSaldoAtual1 = Format$(UltimoSaldo, "##,##0.00")

Call FechaDB

Call Rotina_Tratar_Credito
Call Rotina_Tratar_Debito

Call Rotina_Tratar_Nao_Processado

End Sub

Public Sub converte_csv()
   

   Call Rotina_AbrirBanco
   
   usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
   If usu.EOF Then
      MsgBox ("ERRO: Usuário não habilitado para atualização e/ou consulta de Extrato."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   MesExtenso = txtMesExtenso
   
   MesAno = Month(Date) & "-" & MesExtenso

   'caminhoArquivo = "C:\Meus Documentos\SISTEMA SHB\"
   caminhoArquivo = usu!usuEnderecoOneDrive
   DataHojeCSV = Format$(Date, "dd-mm-yyyy")
   nomeArquivo = "Financeiro\EXTRATOS OFC e PDF\" & Year(Date) & "\" & MesAno & "\" & "Extrato_9298-606931-" & DataHojeCSV & ".xls"
   nomeArquivoNovo = "Financeiro\EXTRATOS OFC e PDF\" & Year(Date) & "\" & MesAno & "\" & "Extrato_9298-606931-" & DataHojeCSV & ".csv"
   
   Set excelApp = New Excel.Application
   
   If Dir(caminhoArquivo & nomeArquivo, vbArchive) = "" Then
70    MsgBox "Não foi possível atualizar o extrato. " & vbCrLf & _
      "O extrato do dia não foi localizado!", vbCritical
      EncontreiExtrato = 0
80    Exit Sub
   Else
      EncontreiExtrato = 1
90 End If

   Set oWB = excelApp.Workbooks.Open(FileName:=caminhoArquivo & nomeArquivo)
   
   oWB.Sheets.Copy
   
   oWB.SaveAs FileName:=caminhoArquivo & nomeArquivoNovo, FileFormat:=xlCSV, local:=True
   
    oWB.Close
520 Set oWB = Nothing
530 excelApp.Quit
540 Set excelApp = Nothing

End Sub

Public Sub carrega_tabela_extrato()
Call Rotina_AbrirBanco
Dim nomeCompleto As String
   
   nomeCompleto = caminhoArquivo & nomeArquivoNovo
   nomeCompleto = Replace$(nomeCompleto, "\", "/")

   db.BeginTrans
   
   rs.Open "Delete from Extrato", db, 3, 3
   
   db.Execute ("LOAD DATA LOCAL INFILE '" & nomeCompleto & "' INTO TABLE Extrato FIELDS TERMINATED BY ';' LINES TERMINATED BY '\n'; ")

   db.CommitTrans
   
FechaDB
End Sub

