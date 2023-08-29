VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEquipamento 
   Caption         =   "frmEquipamento"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   16560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   14760
      TabIndex        =   64
      Top             =   6240
      Width           =   1695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Novo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2520
         Width           =   1400
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H000000FF&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H0000FF00&
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   1400
      End
   End
   Begin VB.Frame frameCalibracao 
      Caption         =   "Calibração/Manutenção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   52
      Top             =   6240
      Width           =   14655
      Begin VB.OptionButton optRetornoCalib 
         Caption         =   "Retorno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cmbStatusCalibracao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11640
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1680
         Width           =   2895
      End
      Begin VB.OptionButton optRemessaCalibSim 
         Caption         =   "Enviar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optRemessaCalibNao 
         Caption         =   "Não Envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1050
         Width           =   1575
      End
      Begin VB.TextBox txtCertificado 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7680
         TabIndex        =   14
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtMetodologia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11160
         TabIndex        =   15
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox cmbLaboratorio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   -2040
         Width           =   3735
      End
      Begin VB.Frame frameEnvioCalibracao 
         Caption         =   "Envio e Recebimento de Calibração/Manutenção"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   1800
         TabIndex        =   53
         Top             =   1080
         Width           =   9735
         Begin VB.Frame frameMovCalibracao 
            Height          =   1575
            Left            =   240
            TabIndex        =   54
            Top             =   240
            Width           =   9375
            Begin MSComCtl2.DTPicker dtRetornoCalib 
               Height          =   375
               Left            =   120
               TabIndex        =   22
               Top             =   1080
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   248053761
               CurrentDate     =   44858
            End
            Begin VB.TextBox txtDocRemessa 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5760
               TabIndex        =   21
               Top             =   360
               Width           =   3495
            End
            Begin VB.ComboBox cmbLaboratorioNew 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   360
               Width           =   3615
            End
            Begin VB.TextBox txtCertificadoNew 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               TabIndex        =   23
               Top             =   1080
               Width           =   3615
            End
            Begin VB.TextBox txtMetodologiaNew 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5760
               TabIndex        =   24
               Top             =   1080
               Width           =   3495
            End
            Begin MSComCtl2.DTPicker dtEnvioCalib 
               Height          =   375
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   248053761
               CurrentDate     =   44858
            End
            Begin VB.Label Label21 
               Caption         =   "Laboratório"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2040
               TabIndex        =   67
               Top             =   120
               Width           =   1695
            End
            Begin VB.Label Label20 
               Caption         =   "Data de Retorno"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label Label19 
               Caption         =   "O.S (Doc de Remessa)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5880
               TabIndex        =   65
               Top             =   120
               Width           =   2295
            End
            Begin VB.Label Label16 
               Caption         =   "Data Envio"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   120
               Width           =   1935
            End
            Begin VB.Label Label17 
               Caption         =   "Novo Certificado"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2040
               TabIndex        =   56
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label18 
               Caption         =   "Metodologia"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5760
               TabIndex        =   55
               Top             =   840
               Width           =   1935
            End
         End
      End
      Begin VB.ComboBox cmbLaboratorioAtual 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtDataValidade 
         Height          =   375
         Left            =   2160
         TabIndex        =   59
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   247857153
         CurrentDate     =   44857
      End
      Begin MSComCtl2.DTPicker dtDataCalibracao 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   247857153
         CurrentDate     =   44857
      End
      Begin VB.Label Label4 
         Caption         =   "Data Calib.Manut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   68
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Validade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   63
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Laboratório"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   62
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Certificado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   61
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "Metodologia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11160
         TabIndex        =   60
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame 
      Height          =   4815
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   1320
      Width           =   16335
      Begin VB.CommandButton cmdFiltar 
         BackColor       =   &H0000FF00&
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cmbTipoDeEquipamento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   840
         Width           =   2655
      End
      Begin VB.ListBox lstEquipamento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   120
         TabIndex        =   71
         Top             =   1920
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker dtDataInicioOperacao 
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   4320
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   247857153
         CurrentDate     =   44720
      End
      Begin VB.ComboBox cmbJBX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   10
         Top             =   3480
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtDataImportacao 
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   247857153
         CurrentDate     =   44720
      End
      Begin MSComCtl2.DTPicker dtDataFabricacao 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   247857153
         CurrentDate     =   44720
      End
      Begin VB.TextBox txtNotaFiscal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11760
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox txtSerie 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         TabIndex        =   5
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox txtModelo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   4
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtOrigem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8160
         TabIndex        =   8
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtFabricante 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   2
         Top             =   1920
         Width           =   3735
      End
      Begin VB.ComboBox cmbTipoEquipamento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox cmbCodEquipamento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label14 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1250
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Equipto."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   860
         Width           =   1095
      End
      Begin VB.Label txtDescricaoCompleta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   51
         Top             =   1200
         Width           =   12135
      End
      Begin VB.Label txtDescResumida 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   29
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label Label9 
         Caption         =   "Data Início Operação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   50
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Unidade Operacional"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11760
         TabIndex        =   49
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Label lblUnidadeOperacional 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11760
         TabIndex        =   48
         Top             =   3480
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   "Cliente Atual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   47
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label lblCliente 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         TabIndex        =   46
         Top             =   3480
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Conjunto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   45
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Nota Fiscal/Delivery Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11760
         TabIndex        =   44
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Série"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12240
         TabIndex        =   43
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Modêlo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   42
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lblLabel11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Importação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6120
         TabIndex        =   41
         Top             =   2400
         Width           =   1650
      End
      Begin VB.Label lblLabel10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "País de Origem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8160
         TabIndex        =   40
         Top             =   2400
         Width           =   1425
      End
      Begin VB.Label lblLabel9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante(Marca)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   39
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label lblLabel8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fabricação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   38
         Top             =   2400
         Width           =   1560
      End
      Begin VB.Label lblLabel6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição Completa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   37
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label lblLabel5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição Resumida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7560
         TabIndex        =   36
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblLabel4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Equipamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLabel3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Equipto SHB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame Frame 
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   16335
      Begin VB.TextBox txtHoje 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   32
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblLabel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hoje"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12480
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro e Atualização de Equipamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         TabIndex        =   30
         Top             =   360
         Width           =   7185
      End
   End
End
Attribute VB_Name = "frmEquipamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Verifica As String
Dim Dia As String
Dim Mes As String
Dim Ano As String

Dim DiaCalib As String
Dim MesCalib As String
Dim AnoCalib As String

Dim DataProxManut As Date
Dim DataProxManutCalib As Date
Dim DataProxCalibracao As String

Dim ErroCritica As Integer
Dim Resp As String
Dim DataInvertida As String

Dim wsUnidTempoPeriodica As String
Dim wsQtdTempoPeriodica As Integer
Dim wsUnidTempoCalibracao As String
Dim wsQtdTempoCalibracao As Integer
Dim EquipTipoAnterior As String
Dim ProdutoSalvo As String
Dim EquipamentoAux As String
Dim Inclui As Integer
Dim Exclui As Integer

Dim Indice As Integer
Dim Ind As Integer


Private Sub cmbCodequipamento_LostFocus()

Call LimpaCampos

If Not cmbCodEquipamento = Empty Then
   Call CarregaCampos
End If

'Call CarregaProduto

End Sub

Private Sub cmbJBX_LostFocus()

Call Rotina_AbrirBanco

If cmbJBX = " DISPONIVEL" Then
   lblCliente = "SEMIHERMATICS"
   lblUnidadeOperacional = "ESTOQUE DISPONÍVEL"
Else
   If cmbJBX = " MANUTENÇÃO" Then
      lblCliente = "NÃO DISPONÍVEL"
      lblUnidadeOperacional = "equipamento EM MANUTENÇÃO"
   End If
End If

Prod.Open "Select * from Produto where chProduto = ('" & cmbJBX & "')", db, 3, 3

If Prod.EOF Then
   MsgBox ("Produto não cadastrado. Verificar"), vbCritical
   Call FechaDB
   Exit Sub
End If

lblCliente = Prod!prdLocadora
lblUnidadeOperacional = Prod!prdUnidadeOperacional

Call FechaDB

End Sub



Private Sub cmbTipoequipamento_LostFocus()

Call Rotina_AbrirBanco

teq.Open "Select * from EquipamentoTipo where chTipoDeequipamento = ('" & cmbTipoEquipamento & "')", db, 3, 3
If teq.EOF Then
   MsgBox ("Tipo de equipamento não encontrado. Informar um código da lista."), vbInformation
   Call FechaDB
   Exit Sub
End If

wsUnidTempoPeriodica = teq!teqUnidTempoPeriodica
wsQtdTempoPeriodica = teq!teqQtdTempoPeriodica
wsUnidTempoCalibracao = teq!teqUnidTempoCalibracao
wsQtdTempoCalibracao = teq!teqQtdTempoCalibracao

txtDescResumida = teq!teqNomeEquipamentoCurto
txtDescricaoCompleta = teq!teqNomeEquipamentoLongo

Call FechaDB

End Sub

Private Sub cmdExcluir_Click()

Call Rotina_AbrirBanco

eqpt.Open "Select * from Equipamento where chCodEquipamento = ('" & cmbCodEquipamento & "')", db, 3, 3
If eqpt.EOF Then
   MsgBox ("Equipamento oara exclusão não encontrado"), vbInformation
   Call FechaDB
   Exit Sub
End If

Resp = MsgBox("Exclusão de Equipamento. Confirma?", vbExclamation + vbYesNo)
   If Resp = vbYes Then
      eqh.Open "Select * from EquipamentoHistorico where chCodEquipamento = ('" & cmbCodEquipamento & "')", db, 3, 3
      If Not eqh.EOF Then
         eqh.MoveFirst
         Do While Not eqh.EOF
            eqh.Delete
            eqh.MoveNext
         Loop
      Else
         MsgBox ("Equipamento sem histórico."), vbInformation
      End If
      eqpt.Delete
   End If
   
   Call LimpaCampos

   'Call CarregaCampos
   
   Call CarregaProduto

   'cmbCodEquipamento = Empty

   Call Carregaequipamento
   
   Call FechaDB

End Sub

Private Sub cmdFiltar_Click()

lstEquipamento.Clear

Call Rotina_AbrirBanco

If cmbTipoDeEquipamento = " TODOS" And cmbStatus = " TODOS" Then
   Call AcessarTodos
Else
   If cmbTipoDeEquipamento = " TODOS" And Not cmbStatus = " TODOS" Then
      Call TodosComStatus
   Else
      If Not cmbTipoDeEquipamento = " TODOS" And cmbStatus = " TODOS" Then
         Call EqptoComTodos
      Else
         Call EqptocomStatus
      End If
   End If
End If

If Not eqpt.EOF Then


   eqpt.MoveFirst

   lstEquipamento.Clear
   
   eqpt.MoveFirst
   
   Do While Not eqpt.EOF
      If Not EquipTipoAnterior = eqpt!chCodEquipamento Then
         If teq.State = 1 Then
            teq.Close: Set teq = Nothing
         End If
         teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & eqpt!eqptTipoEquipamento & "')", db, 3, 3
         If teq.EOF Then
            MsgBox ("Erro no acesso a Tipo de Equipamento."), vbCritical
            Call FechaDB
            Exit Sub
         End If
      End If
   
      If Not IsNull(eqpt!eqptDataValidade) Then
         DataProxManut = eqpt!eqptDataValidade
      Else
         MsgBox ("Data validade invalida o processo."), vbCritical
         Call FechaDB
         Exit Sub
      End If
   
      lstEquipamento.AddItem eqpt!chCodEquipamento
   
      Indice = lstEquipamento.ListCount
   
       eqpt.MoveNext
   
   Loop

End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub


Private Sub cmdSalvar_Click()

If cmbCodEquipamento = Empty Then
   MsgBox ("Código do equipamento não informado."), vbInformation
   Exit Sub
End If

If cmbTipoEquipamento = Empty Then
   MsgBox ("Tipo de equipamento não informado."), vbInformation
   Exit Sub
End If

If txtFabricante = Empty Then
   MsgBox ("Fabricante do equipamento não informado."), vbInformation
   Exit Sub
End If

If txtModelo = Empty Then
   MsgBox ("Modêlo do equipamento não informado."), vbInformation
   Exit Sub
End If

If txtSerie = Empty Then
   MsgBox ("Série do equipamento não informado."), vbInformation
   Exit Sub
End If

If dtDataFabricacao = Date Then
   MsgBox ("Data de Fabricação do equipamento não informado."), vbInformation
   Exit Sub
End If

If dtDataImportacao = Date Then
   MsgBox ("Data de Importação do equipamento não informado."), vbInformation
   Exit Sub
End If

If txtOrigem = Empty Then
   MsgBox ("País de origem do equipamento não informado."), vbInformation
   Exit Sub
End If

If txtNotaFiscal = Empty Then
   MsgBox ("Número da Nota Fiscal do equipamento não informado."), vbInformation
   Exit Sub
End If

'If txtDocImportacao = Empty Then
'   MsgBox ("Número do Documento de Importação do equipamento não informado."), vbInformation
'   Exit Sub
'End If

If cmbJBX = Empty Then
   MsgBox ("Vinculação do equipamento não informado."), vbInformation
   Exit Sub
End If

If dtDataInicioOperacao = Date Then
   MsgBox ("Data de Início da Operação do equipamento não informado."), vbInformation
   Exit Sub
End If

'If dtDataUltimaManutencao = Date Then
'   MsgBox ("Data da Última Manutenção Periódica do equipamento não informado."), vbInformation
'   Exit Sub
'End If

If optRetornoCalib = True Then
   ErroCritica = 0
   Call CriticaRetornoCalibracao
   If ErroCritica = 0 Then
      Call GeraHistorico
      'MsgBox ("Retorno de Calibração realizado com sucesso"), vbInformation
   Else
      MsgBox ("Verificar inconsistências."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

Call Rotina_AbrirBanco

Inclui = 0
eqpt.Open "Select * from Equipamento where chCodequipamento = ('" & cmbCodEquipamento & "')", db, 3, 3
If eqpt.EOF Then
   Inclui = 1
   eqpt.AddNew
End If

eqpt!chCodEquipamento = cmbCodEquipamento
eqpt!eqptTipoEquipamento = cmbTipoEquipamento
eqpt!eqptModelo = txtModelo
eqpt!eqptFabricante = txtFabricante
eqpt!eqptSerie = txtSerie
eqpt!eqptDataFabricacao = dtDataFabricacao
eqpt!eqptPaisOrigem = txtOrigem
eqpt!eqptDataAquisicao = dtDataImportacao
eqpt!eqptDeliveryOrder = txtNotaFiscal
eqpt!eqptProdVinculado = cmbJBX
eqpt!eqptDataInicioVinculacao = dtDataInicioOperacao
eqpt!eqptDataInicioOperacao = dtDataInicioOperacao
'eqpt!eqptDataUltimaManutencao = dtDataUltimaManutencao
'eqpt!eqptDataProximaManutencao = dtDataProxManutencao

eqpt!eqptDataCalibracao = dtDataCalibracao
eqpt!eqptDataValidade = dtDataValidade
eqpt!eqptLaboratorio = cmbLaboratorioAtual
eqpt!eqptCertificado = txtCertificado
eqpt!eqptMetodologia = txtMetodologia
eqpt!eqptStatusCalibracao = cmbStatusCalibracao

If optRetornoCalib = False Then
   If optRemessaCalibSim = True Then
      ErroCritica = 0
      Call CriticaRemessaCalibracao
      If ErroCritica > 0 Then
         MsgBox ("Erro no comando de envio para calibração"), vbInformation
         Exit Sub
      Else
         Call MovimentaEnvioCalibracao
         eqpt.Update
         MsgBox ("Remessa de equipamento salvo com sucesso."), vbInformation
      End If
   Else
      eqpt.Update
      MsgBox ("equipamento salvo com sucesso."), vbInformation
   End If
Else
   Call AjustaEquipamento
   eqpt.Update
   MsgBox ("Equipamento retornado para a BASE."), vbInformation
End If

Call CarregaProduto

Call FechaDB

Call LimpaCampos

'cmbCodEquipamento = Empty

If Inclui = 1 Then
   Call Carregaequipamento
   Inclui = 0
End If

cmbCodEquipamento.SetFocus

End Sub

Private Sub Command1_Click()

Call LimpaCampos

'cmbCodEquipamento = Empty

Call Carregaequipamento

cmbCodEquipamento.SetFocus

End Sub

Private Sub dtDataCalibracao_LostFocus()

DataProxManutCalib = dtDataCalibracao

DiaCalib = Day(DataProxManutCalib)
MesCalib = Month(DataProxManutCalib)
AnoCalib = Year(DataProxManutCalib)

If wsUnidTempoCalibracao = "Ano" Then
   DataProxCalibracao = Format$(DiaCalib, "00") & "/" & Format$(MesCalib, "00") & "/" & AnoCalib + wsQtdTempoCalibracao
Else
   If wsUnidTempoCalibracao = "Mes" Or wsUnidTempoCalibracao = "Mês" Or wsUnidTempoCalibracao = "MES" Then
      MesCalib = MesCalib + wsQtdTempoCalibracao
      If MesCalib > 12 Then
         MesCalib = MesCalib - 12
         AnoCalib = AnoCalib + 1
      End If
      DataProxCalibracao = Format$(DiaCalib, "00") & "/" & Format$(MesCalib, "00") & "/" & AnoCalib
   End If
End If

dtDataValidade = DataProxCalibracao

If dtDataValidade < Date And cmbStatusCalibracao = "CALIBRADO" Then
   cmbStatusCalibracao = "VENCIDO"
End If

End Sub



Private Sub Form_Load()
txtHoje = Date
dtDataFabricacao = Date
dtDataImportacao = Date
dtDataInicioOperacao = Date
'dtDataUltimaManutencao = Date
'dtDataProxManutencao = Date
dtEnvioCalib = Date
dtRetornoCalib = Date

cmbTipoEquipamento.Clear
'cmbCodEquipamento.Clear
lstEquipamento.Clear
cmbJBX.Clear

optRemessaCalibNao = True
optRemessaCalibNao.ForeColor = &HFF&
optRemessaCalibSim = False
optRetornoCalib = False

dtDataFabricacao = Date
dtDataImportacao = Date
dtDataInicioOperacao = Date
'dtDataUltimaManutencao = Date
'dtDataProxManutencao = Date

cmbStatusCalibracao.AddItem "CALIBRADO"
cmbStatusCalibracao.AddItem "VENCIDO"
cmbStatusCalibracao.AddItem "EM CALIBRAÇÃO"
cmbStatusCalibracao.AddItem "EM MANUTENÇÃO"
cmbStatusCalibracao.AddItem "DISPONÍVEL"
cmbStatusCalibracao.AddItem "INATIVO"

cmbStatus.AddItem " TODOS"
cmbStatus.AddItem "CALIBRADO"
cmbStatus.AddItem "VENCIDO"
cmbStatus.AddItem "EM CALIBRAÇÃO"
cmbStatus.AddItem "EM MANUTENÇÃO"
cmbStatus.AddItem "DISPONÍVEL"
cmbStatus.AddItem "INATIVO"

cmbStatus.ListIndex = 0

Call AtualizaStatus

Call Carregaequipamento

Call Rotina_AbrirBanco

teq.Open "Select * from EquipamentoTipo", db, 3, 3
If teq.EOF Then
   MsgBox ("Tabela de Tipos de equipamentos vazia."), vbInformation
   Call FechaDB
   Exit Sub
End If

teq.MoveFirst

cmbTipoDeEquipamento.AddItem " TODOS"
   
Do While Not teq.EOF
   cmbTipoEquipamento.AddItem teq!chTipoDeEquipamento
   cmbTipoDeEquipamento.AddItem teq!chTipoDeEquipamento
   teq.MoveNext
Loop

cmbTipoDeEquipamento.ListIndex = 0

pes.Open "Select * from Pessoa where pesTipoPessoa = ('" & 1 & "') and pesRamoAtividade = ('" & 1 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Cadastro de fornecedor sem prstador de serviços especializados"), vbCritical
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

cmbLaboratorioAtual.Clear
cmbLaboratorioNew.Clear

cmbLaboratorioAtual.AddItem "SHB Brasil"
cmbLaboratorioNew.AddItem "SHB Brasil"

Do While Not pes.EOF
   cmbLaboratorioAtual.AddItem pes!chPessoa
   cmbLaboratorioNew.AddItem pes!chPessoa
   pes.MoveNext
Loop
   
frameCalibracao.Visible = True
frameEnvioCalibracao.Visible = True

Call FechaDB

End Sub

Public Sub Carregaequipamento()

'cmbCodEquipamento.Clear
lstEquipamento.Clear

Call Rotina_AbrirBanco

eqpt.Open "Select * from Equipamento", db, 3, 3
If Not eqpt.EOF Then
   eqpt.MoveFirst
   
   Do While Not eqpt.EOF
      If Not EquipTipoAnterior = eqpt!chCodEquipamento Then
         If teq.State = 1 Then
            teq.Close: Set teq = Nothing
         End If
         teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & eqpt!eqptTipoEquipamento & "')", db, 3, 3
         If teq.EOF Then
            MsgBox ("Erro no acesso a Tipo de Equipamento."), vbCritical
            Call FechaDB
            Exit Sub
         End If
      End If
      
      If Not IsNull(eqpt!eqptDataValidade) Then
         DataProxManut = eqpt!eqptDataValidade
      Else
         MsgBox ("Data validade invalida o processo."), vbCritical
         Call FechaDB
         Exit Sub
      End If
            
      lstEquipamento.AddItem eqpt!chCodEquipamento
      
      'If (Date + teq!teqDiasAntecedencia) > eqpt!eqptDataValidade Then
      '   lstEquipamento.BackColor = &HFFFF&
      'Else
      '   lstEquipamento.BackColor = &H8000000F
      'End If
      
      Indice = lstEquipamento.ListCount
      
      eqpt.MoveNext
      
   Loop
   
   'Call PoeCorNoList
   
Else
   MsgBox ("Tabela de equipamentos vazia."), vbInformation
End If

End Sub

Public Sub LimpaCampos()

frameCalibracao.Visible = True
frameEnvioCalibracao.Visible = True
cmbTipoEquipamento.ListIndex = 0
txtModelo = Empty
txtFabricante = Empty
txtSerie = Empty
dtDataFabricacao = Date
txtOrigem = Empty
dtDataImportacao = Date
txtNotaFiscal = Empty
'txtDocImportacao = Empty
txtDescResumida = Empty
txtDescricaoCompleta = Empty
lblCliente = Empty
lblUnidadeOperacional = Empty
'cmbJBX = Empty
dtDataInicioOperacao = Date
dtDataInicioOperacao = Date
'dtDataUltimaManutencao = Date
'dtDataProxManutencao = Date
dtDataCalibracao = Date
dtDataValidade = Date
cmbLaboratorioAtual.ListIndex = 0
cmbLaboratorioNew.ListIndex = 0
txtCertificado = Empty
txtMetodologia = Empty
txtCertificadoNew = Empty
txtMetodologiaNew = Empty
dtEnvioCalib = Date
cmbLaboratorioNew.ListIndex = 0
txtDocRemessa = Empty
optRemessaCalibNao = True
optRemessaCalibSim = False
optRemessaCalibNao.ForeColor = &HFF&
optRemessaCalibSim.ForeColor = &H0&
optRetornoCalib.ForeColor = &H0&

End Sub

Public Sub CarregaCampos()

Call Rotina_AbrirBanco

eqpt.Open "Select * from Equipamento where chCodequipamento = ('" & cmbCodEquipamento & "')", db, 3, 3
If eqpt.EOF Then
   Resp = MsgBox("Inclusão de Equipamento. Confirma?", vbExclamation + vbYesNo)
   If Resp = vbYes Then
      cmbTipoEquipamento.SetFocus
   End If
   
   Call FechaDB
   Exit Sub
End If

cmbCodEquipamento = eqpt!chCodEquipamento
cmbTipoEquipamento = eqpt!eqptTipoEquipamento
txtModelo = eqpt!eqptModelo
txtFabricante = eqpt!eqptFabricante
txtSerie = eqpt!eqptSerie
dtDataFabricacao = eqpt!eqptDataFabricacao
txtOrigem = eqpt!eqptPaisOrigem
dtDataImportacao = eqpt!eqptDataAquisicao
txtNotaFiscal = eqpt!eqptDeliveryOrder
cmbJBX = eqpt!eqptProdVinculado
dtDataInicioOperacao = eqpt!eqptDataInicioVinculacao
dtDataInicioOperacao = eqpt!eqptDataInicioOperacao
'dtDataUltimaManutencao = eqpt!eqptDataUltimaManutencao
'dtDataProxManutencao = eqpt!eqptDataProximaManutencao
If Not eqpt!eqptStatusCalibracao = Empty Then
   cmbStatusCalibracao = eqpt!eqptStatusCalibracao
End If

teq.Open "Select * from EquipamentoTipo where chTipoDeequipamento = ('" & cmbTipoEquipamento & "')", db, 3, 3
If teq.EOF Then
   MsgBox ("Tipo de equipamento não encontrado. Informar um código da lista."), vbInformation
   Call FechaDB
   Exit Sub
End If

wsUnidTempoPeriodica = teq!teqUnidTempoPeriodica
wsQtdTempoPeriodica = teq!teqQtdTempoPeriodica
wsUnidTempoCalibracao = teq!teqUnidTempoCalibracao
wsQtdTempoCalibracao = teq!teqQtdTempoCalibracao
txtDescResumida = teq!teqNomeEquipamentoCurto
txtDescricaoCompleta = teq!teqNomeEquipamentoLongo

frameCalibracao.Visible = True
If Not IsNull(eqpt!eqptDataCalibracao) Then
   dtDataCalibracao = eqpt!eqptDataCalibracao
End If
If Not IsNull(eqpt!eqptDataValidade) Then
   dtDataValidade = eqpt!eqptDataValidade
End If
If Not IsNull(eqpt!eqptLaboratorio) Then
   cmbLaboratorioAtual = eqpt!eqptLaboratorio
End If
If Not IsNull(eqpt!eqptCertificado) Then
   txtCertificado = eqpt!eqptCertificado
End If
If Not IsNull(eqpt!eqptMetodologia) Then
   txtMetodologia = eqpt!eqptMetodologia
End If

frameCalibracao.Visible = True
frameEnvioCalibracao.Visible = True


Prod.Open "Select * from Produto where chProduto = ('" & cmbJBX & "')", db, 3, 3

If Prod.EOF Then
   MsgBox ("Produto não cadastrado. Verificar"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbJBX = eqpt!eqptProdVinculado
lblCliente = Prod!prdLocadora
lblUnidadeOperacional = Prod!prdUnidadeOperacional

If Not eqpt!eqptOrdemServicoRemessa = Empty Then
   dtEnvioCalib = eqpt!eqptDataEnvioCalibracao
   cmbLaboratorioNew = eqpt!eqptLaboratotioNew
   txtDocRemessa = eqpt!eqptOrdemServicoRemessa
   cmbStatusCalibracao = eqpt!eqptStatusCalibracao
   frameEnvioCalibracao.Visible = True
   frameMovCalibracao.Visible = True
End If

Call FechaDB

End Sub

Private Sub lstEquipamento_Click()

Call LimpaCampos

cmbCodEquipamento = lstEquipamento.List(lstEquipamento.ListIndex)

If Not cmbCodEquipamento = Empty Then
   Call CarregaCampos
End If

Call CarregaProduto

cmbTipoEquipamento.SetFocus

End Sub

Private Sub optRemessaCalibNao_LostFocus()

optRemessaCalibNao = True
optRemessaCalibSim = False
optRemessaCalibNao.ForeColor = &HFF&
optRemessaCalibSim.ForeColor = &H0&
optRetornoCalib.ForeColor = &H0&

frameEnvioCalibracao.Visible = True

cmbStatusCalibracao.SetFocus

End Sub


Private Sub optRemessaCalibSim_LostFocus()

optRemessaCalibNao = False
optRemessaCalibSim = True
optRemessaCalibSim.ForeColor = &HFF&
optRemessaCalibNao.ForeColor = &H0&
optRetornoCalib.ForeColor = &H0&

frameEnvioCalibracao.Visible = True

dtEnvioCalib.SetFocus

End Sub

Private Sub optRetornoCalib_LostFocus()

optRemessaCalibNao = False
optRemessaCalibSim = False
optRetornoCalib = True
optRetornoCalib.ForeColor = &HFF&
optRemessaCalibNao.ForeColor = &H0&
optRemessaCalibSim.ForeColor = &H0&

frameEnvioCalibracao.Visible = True

dtRetornoCalib.SetFocus

End Sub

Public Sub CriticaRemessaCalibracao()

If dtEnvioCalib > Date Then
   MsgBox ("Data remesssa maior que data de hoje"), vbInformation
   ErroCritica = 1
End If
If cmbLaboratorioNew = Empty Then
   MsgBox ("Laboratório não informado"), vbInformation
   ErroCritica = 1
End If
If txtDocRemessa = Empty Then
   MsgBox ("Documento de remessa (O>S) não informado"), vbInformation
   ErroCritica = 1
End If

End Sub

Public Sub MovimentaEnvioCalibracao()

eqpt!eqptDataEnvioCalibracao = dtEnvioCalib
eqpt!eqptLaboratotioNew = cmbLaboratorioNew
eqpt!eqptOrdemServicoRemessa = txtDocRemessa

End Sub

Public Sub CriticaRetornoCalibracao()

If txtDocRemessa = Empty Then
   MsgBox ("Comando para retorno sem o envio para Calibração."), vbInformation
   ErroCritica = 1
   Exit Sub
End If

If dtRetornoCalib > Date Then
   MsgBox ("Data de retorno maior que Hoje."), vbInformation
   ErroCritica = 1
   Exit Sub
End If

If txtCertificadoNew = Empty Then
   MsgBox ("Novo Certificado não Informado."), vbInformation
   ErroCritica = 1
   Exit Sub
End If

If txtMetodologiaNew = Empty Then
   MsgBox ("Metodologia utilizada na Calibração atual não informada."), vbInformation
   ErroCritica = 1
   Exit Sub
End If

End Sub

Public Sub GeraHistorico()

Call Rotina_AbrirBanco

Ano = Year(dtDataCalibracao)
Mes = Month(dtDataCalibracao)
Dia = Day(dtDataCalibracao)

DataInvertida = Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

eqh.Open "Select * from EquipamentoHistorico where chCodEquipamento = ('" & cmbCodEquipamento & "') and eqhDataCalibracao = ('" & DataInvertida & "')", db, 3, 3
If eqh.EOF Then
   eqh.AddNew
End If

eqh!chCodEquipamento = cmbCodEquipamento
eqh!eqhDataCalibracao = DataInvertida
eqh!eqhDataValidade = dtDataValidade
eqh!eqhDataRemessaValidacaoNew = dtEnvioCalib
eqh!eqhDataRetornoValidacao = dtRetornoCalib
eqh!eqhLaboratorio = cmbLaboratorioAtual
eqh!eqhCertificado = txtCertificado
eqh!eqhMetodologia = txtMetodologia
eqh!eqhLaboratorioNew = cmbLaboratorioNew
eqh!eqhOrdemDeServico = txtDocRemessa
eqh!eqhCertificadoNew = txtCertificadoNew
eqh!eqhMtodologiaNew = txtMetodologiaNew
eqh!eqhCliente = lblCliente
eqh!eqhUnidadeOperacional = lblUnidadeOperacional
eqh!eqhDataInicioOperacao = dtDataInicioOperacao
eqh!eqhDataFimOperacao = dtEnvioCalib

eqh.Update

Call FechaDB

End Sub

Public Sub AjustaEquipamento()

eqpt!eqptDataInicioVinculacao = dtRetornoCalib
eqpt!eqptDataInicioOperacao = dtRetornoCalib
eqpt!eqptDataUltimaManutencao = dtRetornoCalib

'Call CalculaDataProximaManutencao

eqpt!eqptDataProximaManutencao = DataProxManut

eqpt!eqptDataCalibracao = dtRetornoCalib
eqpt!eqptDataValidade = DataProxManut
eqpt!eqptLaboratorio = cmbLaboratorioNew
eqpt!eqptCertificado = txtCertificadoNew
eqpt!eqptMetodologia = txtMetodologiaNew
eqpt!eqptStatusCalibracao = "CALIBRADO"

End Sub

'Public Sub CalculaDataProximaManutencao()
'
'DataProxManut = dtDataUltimaManutencao
'
'Dia = Day(DataProxManut)
'Mes = Month(DataProxManut)
'Ano = Year(DataProxManut)
'
'If cmbCodEquipamento = Empty Then
'   Call FechaDB
'   cmbCodEquipamento.SetFocus
'   Exit Sub
'End If
'
'If wsUnidTempoPeriodica = "Ano" Then
'   DataProxManut = Format$(Dia, "00") & "/" & Format$(Mes, "00") & "/" & Ano + wsQtdTempoPeriodica
'Else
'   If teq!teqUnidTempoPeriodica = "MES" Then
'      Mes = Mes + eqpt!teqQtdTempoPeriodica
'      If Mes > 12 Then
'         Mes = Mes - 12
'         Ano = Ano + 1
'      End If
'      DataProxManut = Format$(Dia, "00") & Format$(Mes, "00") & Ano + wsQtdTempoPeriodica
'   End If
'End If
'
'dtDataProxManutencao = DataProxManut
'
'End Sub

Public Sub CarregaProduto()

ProdutoSalvo = cmbJBX

cmbJBX.Clear

If nfd.State = 0 Then
   Call Rotina_AbrirBanco
End If

Prod.Open "Select * from Produto where prdOrdemApresentacao = ('" & 0 & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Produto vazio. Informar ao analista responsável."), vbInformation
   Call FechaDB
   Exit Sub
End If

Prod.MoveFirst

Do While Not Prod.EOF
   Verifica = Mid$(Prod!prdNomeProd, 1, 12)
   If Verifica = "JUNCTION BOX" Then
      cmbJBX.AddItem Prod!chProduto
   End If
   Prod.MoveNext
Loop

cmbJBX = ProdutoSalvo
'cmbJBX.ListIndex = 0

End Sub

Public Sub PoeCorNoList()

Call Rotina_AbrirBanco

For Ind = 1 To Indice

   If eqpt.State = 1 Then
      eqpt.Close: Set eqpt = Nothing
   End If
   
   EquipamentoAux = lstEquipamento.List(Ind)
   
   eqpt.Open "Select * from Equipamento where chCodEquipamento = ('" & EquipamentoAux & "')", db, 3, 3
   If eqpt.EOF Then
      MsgBox ("Erro. Não encontrei o registro apontado no list."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   DataProxManut = eqpt!eqptDataValidade
         
   If teq.State = 1 Then
      teq.Close: Set teq = Nothing
   End If
         
   teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & eqpt!eqptTipoEquipamento & "')", db, 3, 3
   If Not teq.EOF Then
         
      If (DataProxManut - teq!teqDiasAntecedencia) < Date Then
         lstEquipamento.ForeColor = &HFF&
      Else
         lstEquipamento.ForeColor = &H0&
      End If
   Else
      MsgBox ("Tipo de equipamento não encontrado. Analista Responsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If
Next
End Sub

Public Sub AcessarTodos()
eqpt.Open "Select * from Equipamento", db, 3, 3
If eqpt.EOF Then
   MsgBox ("Não há registros com a chave solicitada."), vbCritical
End If
End Sub

Public Sub TodosComStatus()

eqpt.Open "Select * from Equipamento where eqptStatusCalibracao = ('" & cmbStatus & "')", db, 3, 3
If eqpt.EOF Then
   MsgBox ("Não há registros com a chave solicitada."), vbCritical
End If
End Sub

Public Sub EqptoComTodos()

eqpt.Open "Select * from Equipamento where eqptTipoEquipamento = ('" & cmbTipoDeEquipamento & "')", db, 3, 3
If eqpt.EOF Then
   MsgBox ("Não há registros com a chave solicitada."), vbCritical
End If
End Sub

Public Sub EqptocomStatus()
eqpt.Open "Select * from Equipamento where eqptTipoEquipamento = ('" & cmbTipoDeEquipamento & "') and eqptStatusCalibracao = ('" & cmbStatus & "')", db, 3, 3
If eqpt.EOF Then
   MsgBox ("Não há registros com a chave solicitada."), vbCritical
End If
End Sub

Public Sub AtualizaStatus()

Call Rotina_AbrirBanco

eqpt.Open "Select * from Equipamento", db, 3, 3
If Not eqpt.EOF Then
   eqpt.MoveFirst
   
   Do While Not eqpt.EOF
 
      If Not IsNull(eqpt!eqptDataValidade) Then
         DataProxManutCalib = eqpt!eqptDataValidade
         dtDataValidade = eqpt!eqptDataValidade
      Else
         MsgBox ("Data validade invalida o processo."), vbCritical
         Call FechaDB
         Exit Sub
      End If

            
      If dtDataValidade < Date And eqpt!eqptStatusCalibracao = "CALIBRADO" Then
         cmbStatusCalibracao = "VENCIDO"
         eqpt!eqptStatusCalibracao = "VENCIDO"
         eqpt.Update
      End If
      
      eqpt.MoveNext
      
   Loop
Else
   MsgBox ("Tabela de Equipamentos vazia. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
End If

End Sub
