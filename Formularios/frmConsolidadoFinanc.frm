VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsolidadoFinanc 
   Caption         =   "frmConsolidadoFinanc"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20370
   LinkTopic       =   "Form3"
   ScaleHeight     =   9945
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20415
      Begin VB.Frame Frame10 
         Height          =   855
         Left            =   13080
         TabIndex        =   64
         Top             =   240
         Width           =   7215
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H008080FF&
            Caption         =   "Sair"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3600
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdConsulta 
            BackColor       =   &H00FFFF00&
            Caption         =   "Consulta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2565
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbFiltro 
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            TabIndex        =   65
            Top             =   360
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtDataHoje 
            Height          =   360
            Left            =   4680
            TabIndex        =   68
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   16777194
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
         Begin VB.Label Label23 
            Caption         =   "Hoje"
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
            Left            =   4920
            TabIndex        =   70
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Filtro"
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
            Left            =   795
            TabIndex        =   69
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Analítico Semanal Consultado"
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
         Left            =   7200
         TabIndex        =   58
         Top             =   240
         Width           =   5895
         Begin MSMask.MaskEdBox txtDataCtaPagDe 
            Height          =   375
            Left            =   720
            TabIndex        =   59
            Top             =   360
            Width           =   1635
            _ExtentX        =   2884
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
         Begin MSMask.MaskEdBox txtDataCtaPagAte 
            Height          =   375
            Left            =   3120
            TabIndex        =   60
            Top             =   360
            Width           =   1635
            _ExtentX        =   2884
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
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Até"
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
            Left            =   2520
            TabIndex        =   62
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "De"
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
            Left            =   195
            TabIndex        =   61
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Contas a Receber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8535
         Left            =   7200
         TabIndex        =   57
         Top             =   1080
         Width           =   5895
         Begin MSFlexGridLib.MSFlexGrid GrdCtaPagar 
            Height          =   4020
            Left            =   0
            TabIndex        =   74
            Top             =   4440
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   7091
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColor       =   16777194
            BackColorFixed  =   16776960
            ForeColorSel    =   -2147483635
            BackColorBkg    =   16777194
            AllowBigSelection=   -1  'True
            FormatString    =   "|Data       |Desp./Colaborador    |Valor           "
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
         Begin MSFlexGridLib.MSFlexGrid GrdCtaReceber 
            Height          =   3855
            Left            =   0
            TabIndex        =   73
            Top             =   360
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   6800
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColor       =   16777194
            BackColorFixed  =   16776960
            ForeColorFixed  =   0
            BackColorBkg    =   16777194
            FormatString    =   "|Data       |Cliente                        |Valor          "
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
         Begin VB.Label Label28 
            Caption         =   "Contas a Pagar"
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
            Left            =   120
            TabIndex        =   63
            Top             =   4200
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Semanal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   120
         TabIndex        =   55
         Top             =   1080
         Width           =   7095
         Begin MSFlexGridLib.MSFlexGrid GridSemana 
            Height          =   5055
            Left            =   -120
            TabIndex        =   71
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   8916
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColor       =   16777194
            ForeColor       =   -2147483630
            BackColorFixed  =   16776960
            BackColorBkg    =   16777194
            FormatString    =   "De          |Até         |Receber      |Pagar        |Saldo          "
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Até a data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   0
         TabIndex        =   30
         Top             =   6480
         Width           =   7215
         Begin VB.TextBox txtSHESaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Left            =   5325
            TabIndex        =   88
            Top             =   3120
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtSHEPagar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            Left            =   3840
            TabIndex        =   87
            Top             =   3120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtSHEMalasia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            Left            =   2160
            TabIndex        =   86
            Top             =   3120
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtInvestSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Left            =   5450
            TabIndex        =   84
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtPagarInvest 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Left            =   3960
            TabIndex        =   83
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtInvestFinanc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            Left            =   2280
            TabIndex        =   82
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtCotasSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Left            =   5450
            TabIndex        =   80
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtCotasPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   3960
            TabIndex        =   79
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtCotas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Left            =   2280
            TabIndex        =   78
            Top             =   1920
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtDataAte 
            Height          =   375
            Left            =   1200
            TabIndex        =   31
            Top             =   480
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   16777194
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDataDE 
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   480
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   16777194
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SHE Malásia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   -60
            TabIndex        =   85
            Top             =   3120
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Invest. Financ"
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
            Left            =   60
            TabIndex        =   81
            Top             =   2280
            Width           =   2176
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cotas"
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
            Left            =   60
            TabIndex        =   77
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "De"
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
            Left            =   60
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Até"
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
            TabIndex        =   53
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pend. na semana"
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
            Left            =   60
            TabIndex        =   52
            Top             =   840
            Width           =   2176
         End
         Begin VB.Label txtPendenteAReceber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2280
            TabIndex        =   51
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label txtPendenteAPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3980
            TabIndex        =   50
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label txtSaldoPendente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5450
            TabIndex        =   49
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Receber"
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
            Left            =   2280
            TabIndex        =   48
            Top             =   555
            Width           =   1725
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pagar"
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
            Left            =   3980
            TabIndex        =   47
            Top             =   555
            Width           =   1455
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo"
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
            Left            =   5450
            TabIndex        =   46
            Top             =   555
            Width           =   1725
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Process até hoje"
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
            Left            =   60
            TabIndex        =   45
            Top             =   1560
            Width           =   2176
         End
         Begin VB.Label txtProcesAReceber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2280
            TabIndex        =   44
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label txtProcesAPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3980
            TabIndex        =   43
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label txtSaldoProces 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5450
            TabIndex        =   42
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total................................"
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
            Left            =   60
            TabIndex        =   41
            Top             =   2650
            Width           =   2175
         End
         Begin VB.Label txtTotalAReceber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2280
            TabIndex        =   40
            Top             =   2650
            Width           =   1695
         End
         Begin VB.Label txtTotalAPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3975
            TabIndex        =   39
            Top             =   2650
            Width           =   1455
         End
         Begin VB.Label txtSaldoTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5445
            TabIndex        =   38
            Top             =   2650
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "V  a  l  o  r"
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
            Left            =   2280
            TabIndex        =   37
            Top             =   240
            Width           =   4860
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pend até fim do mes"
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
            Left            =   60
            TabIndex        =   36
            Top             =   1200
            Width           =   2176
         End
         Begin VB.Label txtReceberAteFim 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2280
            TabIndex        =   35
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label txtPagarAteFim 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3980
            TabIndex        =   34
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label txtSaldoAteFim 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   5450
            TabIndex        =   33
            Top             =   1200
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Mensal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   12960
         TabIndex        =   29
         Top             =   1080
         Width           =   7455
         Begin MSFlexGridLib.MSFlexGrid GridMensal 
            Height          =   2535
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   4471
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColor       =   16777194
            BackColorFixed  =   16776960
            BackColorBkg    =   16777194
            FormatString    =   "Data   |Receber        |Pagar         |Saldo            |%  Distr."
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
      End
      Begin VB.Frame Frame8 
         Caption         =   "No Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   13080
         TabIndex        =   4
         Top             =   6720
         Width           =   7335
         Begin VB.TextBox txtChaveClassifica 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   360
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label txtSaldoEmAtraso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            Left            =   3840
            TabIndex        =   28
            Top             =   1200
            Width           =   1785
         End
         Begin VB.Label txtPagtosEmAtraso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3840
            TabIndex        =   27
            Top             =   840
            Width           =   1785
         End
         Begin VB.Label txtRecebEmAtraso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   3840
            TabIndex        =   26
            Top             =   480
            Width           =   1785
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Atrasados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Atividade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   24
            Top             =   1680
            Width           =   1305
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Contas a Receber...."
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
            Left            =   1560
            TabIndex        =   23
            Top             =   480
            Width           =   2250
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Contas a Pagar........"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   22
            Top             =   840
            Width           =   2250
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo em Atraso......."
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
            Left            =   1560
            TabIndex        =   21
            Top             =   1200
            Width           =   2250
         End
         Begin VB.Label txtTotalCreditoMensal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            Left            =   3840
            TabIndex        =   20
            Top             =   1680
            Width           =   1785
         End
         Begin VB.Label txtTotalDebitoMensal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            Left            =   3840
            TabIndex        =   19
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label txtSaldoMensal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
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
            Left            =   3840
            TabIndex        =   18
            Top             =   2400
            Width           =   1785
         End
         Begin VB.Label txtReceberAtrasadoPerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5625
            TabIndex        =   17
            Top             =   480
            Width           =   945
         End
         Begin VB.Label txtPagarAtrasadoPerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5625
            TabIndex        =   16
            Top             =   840
            Width           =   945
         End
         Begin VB.Label txtSaldoAtrasadoPerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5625
            TabIndex        =   15
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Height          =   1095
            Left            =   6600
            TabIndex        =   14
            Top             =   480
            Width           =   300
         End
         Begin VB.Label txtPercRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5625
            TabIndex        =   13
            Top             =   1680
            Width           =   945
         End
         Begin VB.Label txtPercPag 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5625
            TabIndex        =   12
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label txtPercSaldo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5625
            TabIndex        =   11
            Top             =   2400
            Width           =   945
         End
         Begin VB.Label Label25 
            BorderStyle     =   1  'Fixed Single
            Height          =   1095
            Left            =   6600
            TabIndex        =   10
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor Total Credor...."
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
            Left            =   1560
            TabIndex        =   9
            Top             =   1680
            Width           =   2250
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo No Período...."
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
            Left            =   1560
            TabIndex        =   8
            Top             =   2400
            Width           =   2250
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor Total Devedor."
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
            Left            =   1560
            TabIndex        =   7
            Top             =   2040
            Width           =   2250
         End
         Begin VB.Label txtDataDeMensal 
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   6
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label txtDataAteMensal 
            BackColor       =   &H00FFFFEA&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   5
            Top             =   2400
            Width           =   1275
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Distribuição do Faturamento do mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   13080
         TabIndex        =   1
         Top             =   3960
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid GridDistrib 
            Height          =   1935
            Left            =   600
            TabIndex        =   75
            Top             =   600
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   3413
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColor       =   16777194
            BackColorFixed  =   16776960
            BackColorBkg    =   16777194
            FormatString    =   "Vencimento          |Valor                  |%Distrib  "
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
         Begin VB.Label Label20 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Faturado no mes......"
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
            Left            =   600
            TabIndex        =   3
            Top             =   240
            Width           =   2250
         End
         Begin VB.Label txtTotalFaturamento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2880
            TabIndex        =   2
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFEA&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Posição Financeira Consolidada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmConsolidadoFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim IndReceber As Byte
Dim IndPagar As Byte
Dim IndMes As Byte
Dim indice As Byte
Dim IndLimiteMensal As Integer
Dim Linha As Integer
Dim LinhaAnalitico As Integer
Dim coluna As Integer
Dim Fim As Byte

Dim FimReceber As Integer
Dim FimPagar As Integer

Dim fimsemana As Byte
Dim fimmensal As Byte
Dim UltimoMes As Date

Dim AcumProcesPagar As Currency
Dim AcumProcesPagarMensal As Currency
Dim AcumPendPagar As Currency
Dim AcumPagarAteFim As Currency
Dim AcumPendPagarMensal As Currency
Dim AcumPagarAtrasado As Currency
Dim AcumPagarAtrasadoMensal As Currency
Dim AcumFaturamentoMensal As Currency
Dim AcumFaturamentoDistrib As Currency
Dim AcumulaConsignacao As Currency

Dim AcumGeralReceber As Currency
Dim AcumGeralPagar As Currency

Dim AcumPendReceber As Currency
Dim AcumPendReceberMensal As Currency
Dim AcumReceberAtrasado As Currency
Dim AcumReceberAteFim As Currency
Dim AcumReceberAtrasadoMensal As Currency
Dim AcumProcesReceber As Currency
Dim AcumProcesReceberMensal As Currency

Dim ValorDebCre As Currency
Dim TotalDebCre As Currency
Dim PercDebCre As Currency

Dim DiaUtilAnterior As Date
Dim dataInicio As Date
Dim dataFim As Date
Dim DataParaCalculo As Date

Dim DataMensal As Date
Dim DataMensalInicio As Date
Dim DataMensalFim As Date
Dim DataIniMesAtu As Date

Dim Dia As Integer
Dim mes As Integer
Dim ano As Integer

Dim AnoAtu As Integer
Dim MesAtu As Integer

Dim DiadaSemana As Integer
Dim NDias As Integer
Public DataInformada As Date

Dim tabDataIni(150) As Date
Dim tabDataFim(150) As Date
Dim TabMesAnoIni(150) As Date
Dim TabMesAnoFim(150) As Date
Dim TabMesFaturamento(24) As Date
Dim TabMesFaturamentoFim(24) As Date
Dim tabValor(150, 3) As Currency
Dim tabValorMensal(150, 4) As Currency
Dim tabValorDistrib(150) As Currency
Dim ChaveClassifica As String

Dim AcumulaCotas As Currency
Dim AcumulaSHEMalasia As Currency
Dim AcumulaInvestFinanc As Currency
Dim ChavePesquisa As String



Private Sub cmdConsulta_Click()

indice = cmbFiltro.ListIndex

Call Rotina_00_Principal

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

cmbFiltro.Clear
cmbFiltro.AddItem "Geral"

AnoAtu = Year(Date)
MesAtu = Month(Date)

Call Rotina_AbrirBanco

Bco.Open "Select * from banco", db, 3, 3
If Bco.EOF Then
   MsgBox ("tabela de bancos vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

Bco.MoveFirst

Do While Not Bco.EOF
   cmbFiltro.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbFiltro.ListIndex = 0

Call Rotina_00_Principal

Call FechaDB

End Sub

Public Sub Rotina_00_Principal()

txtDataHoje = Date

NDias = 1
DataInformada = Date
'DataRetorno = ObterDataUtilAnterior(Datainformada, NDias)

'DiaUtilAnterior = DataRetorno.DiaUtil

Call Rotina_070_Ajusta_Data

'DataInicio = Format$(DataInicio, "dd/mm/yy")Format$(tabDataIni(IndReceber) + 1, "dd/mm/yy")
'txtDataDE = Format$(DataInicio, "dd/mm/yy")
'If DataFim > DataMensalFim Then
'   txtDataAte = DataMensalFim
'Else
'   txtDataAte = DataFim
'End If

Call Rotina_012_Limpa_Cta_Receber

tabDataIni(1) = dataInicio - 1
tabDataFim(1) = dataFim - 1
TabMesAnoIni(1) = DataMensalInicio
TabMesFaturamento(1) = DataMensalInicio
TabMesAnoFim(1) = DataMensalFim
TabMesFaturamentoFim(1) = DataMensalFim

AcumProcesPagar = 0
AcumProcesPagarMensal = 0
AcumPendPagar = 0

AcumPendPagarMensal = 0
AcumPagarAtrasado = 0
AcumPagarAtrasadoMensal = 0

AcumPendReceber = 0
AcumPendReceberMensal = 0
AcumReceberAtrasado = 0
AcumReceberAtrasadoMensal = 0
AcumProcesReceber = 0
AcumProcesReceberMensal = 0
AcumReceberAteFim = 0
AcumPagarAteFim = 0
AcumFaturamentoMensal = 0
AcumFaturamentoDistrib = 0
AcumulaCotas = 0
AcumulaInvestFinanc = 0

Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber", db, 3, 3
If Not ctr.EOF Then

   ctr.MoveFirst

    
   Do While Not ctr.EOF
      
      IndReceber = 1
      IndMes = 1
      
      If (cmbFiltro = "Geral") Or (cmbFiltro = ctr!chCodBcoLart) Then
      
         Call Rotina_030_TabSemana
         
         Call Rotina_040_TabMensal
      End If
         
      ctr.MoveNext
          
   Loop
End If

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If


ctp.Open "Select * from contas_a_pagar", db, 3, 3
If ctp.EOF Then
   'MsgBox ("Não há lançamentos a Débito até a apresente data."), vbInformation
   Call FechaDB
   'Exit Sub
Else

    ctp.MoveFirst


        Do While Not ctp.EOF

           IndPagar = 1
           IndMes = 1

           If (cmbFiltro = "Geral") Or (cmbFiltro = ctp!chCodBcoLart) Then

              If Not ctp!ctpStatus = 2 Then

                 Call Rotina_050_CtaPagar_Semana

                 Call Rotina_060_CtaPagar_Mes
              End If

           End If

           ctp.MoveNext

        Loop
End If

'Descarregar Tabelas na tela de consulta

IndReceber = 1

Do While Not (tabDataIni(IndReceber) = Empty)
   GridSemana.Rows = IndReceber + 1
   GridSemana.TextMatrix(IndReceber, 0) = Format$(tabDataIni(IndReceber) + 1, "dd/mm/yy")
   GridSemana.TextMatrix(IndReceber, 1) = Format$(tabDataFim(IndReceber) + 1, "dd/mm/yy")
   GridSemana.TextMatrix(IndReceber, 2) = Format$(tabValor(IndReceber, 0), "##,##0.00")
   GridSemana.TextMatrix(IndReceber, 3) = Format$(tabValor(IndReceber, 1), "##,##0.00")
   GridSemana.TextMatrix(IndReceber, 4) = Format$(tabValor(IndReceber, 2), "##,##0.00")
   
   GridSemana.Col = 0
   GridSemana.ColSel = 0
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = RGB(100, 51, 51)
   GridSemana.CellFontBold = True
     
   GridSemana.Col = 1
   GridSemana.ColSel = 1
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = RGB(100, 51, 51)
   GridSemana.CellFontBold = True
   
   GridSemana.Col = 2
   GridSemana.ColSel = 2
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = RGB(51, 0, 256)
   GridSemana.CellFontBold = True
     
   GridSemana.Col = 3
   GridSemana.ColSel = 3
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = vbRed
   GridSemana.CellFontBold = True
   
   If tabValor(IndReceber, 2) < 0 Then
      GridSemana.Col = 4
      GridSemana.ColSel = 4
      GridSemana.Row = IndReceber
      GridSemana.RowSel = IndReceber
      GridSemana.CellForeColor = vbRed
      GridSemana.CellFontBold = True
   Else
      GridSemana.Col = 4
      GridSemana.ColSel = 4
      GridSemana.Row = IndReceber
      GridSemana.RowSel = IndReceber
      GridSemana.CellForeColor = vbBlue
   End If
   IndReceber = IndReceber + 1
Loop

txtPendenteAReceber = Format$(AcumPendReceber, "##,##0.00")
txtPendenteAReceber.ForeColor = vbBlue
txtPendenteAPagar = Format$(AcumPendPagar, "##,##0.00")
txtPendenteAPagar.ForeColor = vbRed
txtSaldoPendente = Format$(AcumPendReceber - AcumPendPagar, "##,##0.00")
If AcumPendPagar > AcumPendReceber Then
   txtSaldoPendente.ForeColor = vbRed
Else
   txtSaldoPendente.ForeColor = vbBlue
End If

txtProcesAReceber = Format$(AcumProcesReceber, "##,##0.00")
txtProcesAReceber.ForeColor = vbBlue
txtProcesAPagar = Format$(AcumProcesPagar, "##,##0.00")
txtProcesAPagar.ForeColor = vbRed
txtSaldoProces = Format$(AcumProcesReceber - AcumProcesPagar, "##,##0.00")
If AcumProcesPagar > AcumProcesReceber Then
   txtSaldoProces.ForeColor = vbRed
Else
   txtSaldoProces.ForeColor = vbBlue
End If

txtReceberAteFim = Format$(AcumReceberAteFim, "##,##0.00")
txtReceberAteFim.ForeColor = vbBlue
txtPagarAteFim = Format$(AcumPagarAteFim, "##,##0.00")
txtPagarAteFim.ForeColor = vbRed
txtSaldoAteFim = Format$(AcumReceberAteFim - AcumPagarAteFim, "###,##0.00")
If AcumPagarAteFim > AcumReceberAteFim Then
   txtSaldoAteFim.ForeColor = vbRed
Else
   txtSaldoAteFim.ForeColor = vbBlue
End If

'CONSULTA CONTAS A PAGAR

ChavePesquisa = "COTAS"

rs.Open "SELECT ctpValorDaBoleta FROM contas_a_pagar where chPessoa = ('" & ChavePesquisa & "')", db, 3, 3
If rs.EOF Then
   AcumulaCotas = 0
Else
   Do While Not rs.EOF
      AcumulaCotas = AcumulaCotas + rs!ctpValorDaBoleta
      rs.MoveNext
   Loop
End If

rs.Close

ChavePesquisa = "INVEST FINANC"

rs.Open "SELECT ctpValorDaBoleta FROM contas_a_pagar where chPessoa = ('" & ChavePesquisa & "')", db, 3, 3
If rs.EOF Then
   AcumulaInvestFinanc = 0
Else
   Do While Not rs.EOF
      AcumulaInvestFinanc = AcumulaInvestFinanc + rs!ctpValorDaBoleta
      rs.MoveNext
   Loop
End If

rs.Close

ChavePesquisa = "SHE MALASIA"

rs.Open "SELECT ctpValorDaBoleta FROM contas_a_pagar where chPessoa = ('" & ChavePesquisa & "')", db, 3, 3
If rs.EOF Then
   AcumulaSHEMalasia = 0
Else
   Do While Not rs.EOF
      AcumulaSHEMalasia = AcumulaSHEMalasia + rs!ctpValorDaBoleta
      rs.MoveNext
   Loop
End If

rs.Close

txtCotas = Format$(AcumulaCotas, "##,##0.00")
txtCotasPagar = Format$(0, "##,##0.00")
txtCotasSaldo = Format$(AcumulaCotas, "##,##0.00")

txtInvestFinanc = Format$(AcumulaInvestFinanc, "##,##0.00")
txtPagarInvest = Format$(0, "##,##0.00")
txtInvestSaldo = Format$(AcumulaInvestFinanc, "##,##0.00")

'txtSHEMalasia = Format$(AcumulaSHEMalasia, "##,##0.00")
'txtSHEPagar = Format$(0, "##,##0.00")
'txtSHESaldo = Format$(AcumulaSHEMalasia, "##,##0.00")

'txtTotalAReceber = Format$(AcumPendReceber + AcumProcesReceber + AcumReceberAteFim + AcumulaCotas + AcumulaInvestFinanc + AcumulaSHEMalasia, "##,##0.00")
txtTotalAReceber = Format$(AcumPendReceber + AcumProcesReceber + AcumReceberAteFim + AcumulaCotas + AcumulaInvestFinanc, "##,##0.00")
txtTotalAReceber.ForeColor = vbBlue
txtTotalAPagar = Format$(AcumPendPagar + AcumProcesPagar + AcumPagarAteFim, "##,##0.00")
txtTotalAPagar.ForeColor = vbRed
'txtSaldoTotal = Format$((AcumPendReceber + AcumProcesReceber + AcumReceberAteFim + AcumulaCotas + AcumulaInvestFinanc + AcumulaSHEMalasia) - (AcumPendPagar + AcumProcesPagar + AcumPagarAteFim), "##,##0.00")
txtSaldoTotal = Format$((AcumPendReceber + AcumProcesReceber + AcumReceberAteFim + AcumulaCotas + AcumulaInvestFinanc) - (AcumPendPagar + AcumProcesPagar + AcumPagarAteFim), "##,##0.00")
If txtSaldoTotal < 0 Then
   txtSaldoTotal.ForeColor = vbRed
Else
   txtSaldoTotal.ForeColor = vbBlue
End If

txtCotas.ForeColor = vbBlue
txtCotasSaldo.ForeColor = vbBlue
txtInvestFinanc.ForeColor = vbBlue
txtInvestSaldo.ForeColor = vbBlue
txtSHEMalasia.ForeColor = vbBlue
txtSHESaldo.ForeColor = vbBlue
txtCotasPagar.ForeColor = vbRed
txtPagarInvest.ForeColor = vbRed
txtSHEPagar.ForeColor = vbRed

txtRecebEmAtraso = Format$(AcumReceberAtrasado, "##,##0.00")
txtPagtosEmAtraso = Format$(AcumPagarAtrasado, "##,##0.00")
txtSaldoEmAtraso = Format$(AcumReceberAtrasado - AcumPagarAtrasado, "##,##0.00")
If AcumReceberAtrasado < AcumPagarAtrasado Then
   txtSaldoEmAtraso.ForeColor = vbRed
Else
   txtSaldoEmAtraso.ForeColor = vbBlue
End If

'MENSAL

AcumGeralReceber = 0
AcumGeralPagar = 0

IndMes = 1

Do While Not (TabMesAnoIni(IndMes) = Empty)
   GridMensal.Rows = IndMes + 1
   GridMensal.TextMatrix(IndMes, 0) = Format$(TabMesAnoIni(IndMes), "mmm/yy")
   GridMensal.TextMatrix(IndMes, 1) = Format$(tabValorMensal(IndMes, 0), "##,##0.00")
   GridMensal.TextMatrix(IndMes, 2) = Format$(tabValorMensal(IndMes, 1), "##,##0.00")
   GridMensal.TextMatrix(IndMes, 3) = Format$(tabValorMensal(IndMes, 2), "##,##0.00")
   
   AcumGeralReceber = AcumGeralReceber + tabValorMensal(IndMes, 0)
   AcumGeralPagar = AcumGeralPagar + tabValorMensal(IndMes, 1)
   
   GridMensal.Col = 0
   GridMensal.ColSel = 0
   GridMensal.Row = IndMes
   GridMensal.RowSel = IndMes
   GridMensal.CellForeColor = RGB(100, 51, 51)
   GridMensal.CellFontBold = True
   
   GridMensal.Col = 1
   GridMensal.ColSel = 1
   GridMensal.Row = IndMes
   GridMensal.RowSel = IndMes
   GridMensal.CellForeColor = vbBlue
   GridMensal.CellFontBold = True
     
   GridMensal.Col = 2
   GridMensal.ColSel = 2
   GridMensal.Row = IndMes
   GridMensal.RowSel = IndMes
   GridMensal.CellForeColor = vbRed
   GridMensal.CellFontBold = True
   
   If tabValorMensal(IndMes, 2) < 0 Then
      GridMensal.Col = 3
      GridMensal.ColSel = 3
      GridMensal.Row = IndMes
      GridMensal.RowSel = IndMes
      GridMensal.CellForeColor = vbRed
      GridMensal.CellFontBold = True
   Else
      GridMensal.Col = 3
      GridMensal.ColSel = 3
      GridMensal.Row = IndMes
      GridMensal.RowSel = IndMes
      GridMensal.CellForeColor = vbBlue
      GridMensal.CellFontBold = True
   End If
   IndMes = IndMes + 1
Loop

'AcumulaConsignacao = 0

'If TabConsignacao.RecordCount > 0 Then

'TabConsignacao.MoveFirst

'    Do While Not TabConsignacao.EOF
'       If TabConsignacao("constatus  < 2 Then
'          AcumulaConsignacao = AcumulaConsignacao + (TabConsignacao("convaloratual  - TabConsignacao("convalorrecebido )
'       End If
'       TabConsignacao.MoveNext
'    Loop
'End If

AcumGeralReceber = AcumGeralReceber '+ AcumReceberAtrasado
AcumGeralPagar = AcumGeralPagar '+ AcumPagarAtrasado

'txtTotalConsignado = Format$(AcumulaConsignacao, "###,##0.00")

IndLimiteMensal = IndMes

txtDataDeMensal = Format$(TabMesAnoIni(1), "mmm/yy")
txtDataAteMensal = Format$(UltimoMes, "mmm/yy")

txtTotalCreditoMensal = Format$(AcumGeralReceber, "##,##0.00")
txtTotalCreditoMensal.ForeColor = vbBlue
txtTotalDebitoMensal = Format$(AcumGeralPagar, "##,##0.00")
txtTotalDebitoMensal.ForeColor = vbRed
txtSaldoMensal = Format$(AcumGeralReceber - AcumGeralPagar, "##,##0.00")
If AcumGeralPagar > AcumGeralReceber Then
   txtSaldoMensal.ForeColor = vbRed
Else
   txtSaldoMensal.ForeColor = vbBlue
End If
If (AcumGeralReceber + AcumGeralPagar) > 0 Then
   txtPercRec = Format$((AcumGeralReceber / ((AcumGeralReceber + AcumGeralPagar)) * 100), "#0.00") & "%"
   txtPercPag = Format$((AcumGeralPagar / (AcumGeralReceber + AcumGeralPagar) * 100), "#0.00") & "%"
   txtPercSaldo = Format$(((AcumGeralReceber - AcumGeralPagar) / (AcumGeralReceber + AcumGeralPagar) * 100), "#0.00") & "%"
End If
IndMes = 1

txtTotalFaturamento = Format$(AcumFaturamentoDistrib, "##,##0.00")
txtTotalFaturamento.ForeColor = vbBlue

Do While Not TabMesFaturamento(IndMes) = Empty
   GridDistrib.Rows = IndMes + 1
   GridDistrib.TextMatrix(IndMes, 0) = Format$(TabMesAnoIni(IndMes), "Mmmm/yyyy")
   GridDistrib.TextMatrix(IndMes, 1) = Format$(tabValorDistrib(IndMes), "##,##0.00")
   If AcumFaturamentoDistrib = 0 Then
      GridDistrib.TextMatrix(IndMes, 2) = 0
   Else
      GridDistrib.TextMatrix(IndMes, 2) = Format$((tabValorDistrib(IndMes) / AcumFaturamentoDistrib) * 100, "#0.00") & "%"
   End If
   GridDistrib.Col = 1
   GridDistrib.ColSel = 1
   GridDistrib.Row = IndMes
   GridDistrib.RowSel = IndMes
   GridDistrib.CellForeColor = vbBlue
   GridDistrib.CellFontBold = True
   IndMes = IndMes + 1
Loop

'Atenção

If AcumGeralReceber > 0 Then
   'txtConsigPercent = Format$(AcumulaConsignacao / AcumGeralReceber, "#0.00") & "%"
   txtReceberAtrasadoPerc = Format$((AcumReceberAtrasado / AcumGeralReceber) * 100, "#0.00") & "%"
End If
If AcumGeralPagar > 0 Then
   txtPagarAtrasadoPerc = Format$((AcumPagarAtrasado / AcumGeralPagar) * 100, "#0.00") & "%"
Else
   txtPagarAtrasadoPerc = Format$(0, "#0.00") & "%"
End If
If (AcumGeralReceber + AcumGeralPagar) > 0 Then
   txtSaldoAtrasadoPerc = Format$(((AcumReceberAtrasado - AcumPagarAtrasado) / (AcumGeralReceber - AcumGeralPagar) * 100), "#0.00") & "%"
End If
End Sub

Public Sub Rotina_012_Limpa_Cta_Receber()

GridSemana.Rows = 2
GridDistrib.Rows = 2
IndReceber = 1
GridSemana.TextMatrix(IndReceber, 0) = Empty
GridSemana.TextMatrix(IndReceber, 1) = Empty
GridSemana.TextMatrix(IndReceber, 2) = Empty
GridSemana.TextMatrix(IndReceber, 3) = Empty
GridSemana.TextMatrix(IndReceber, 4) = Empty
GridMensal.TextMatrix(IndReceber, 0) = Empty
GridMensal.TextMatrix(IndReceber, 1) = Empty
GridMensal.TextMatrix(IndReceber, 2) = Empty
GridMensal.TextMatrix(IndReceber, 3) = Empty
GridMensal.TextMatrix(IndReceber, 4) = Empty
GridMensal.TextMatrix(0, 4) = "% Distr."

GridDistrib.TextMatrix(IndReceber, 0) = Empty
GridDistrib.TextMatrix(IndReceber, 1) = Empty
GridDistrib.TextMatrix(IndReceber, 2) = Empty

For IndReceber = 1 To 150
    
    If IndReceber < 10 Then
       TabMesFaturamento(IndReceber) = Empty
    End If
    
    tabDataIni(IndReceber) = Empty
    tabDataFim(IndReceber) = Empty
    TabMesAnoIni(IndReceber) = Empty
    TabMesAnoFim(IndReceber) = Empty
    
    tabValor(IndReceber, 0) = 0
    tabValor(IndReceber, 1) = 0
    tabValor(IndReceber, 2) = 0
    tabValorMensal(IndReceber, 0) = 0
    tabValorMensal(IndReceber, 1) = 0
    tabValorMensal(IndReceber, 2) = 0
    tabValorMensal(IndReceber, 3) = 0
    
    tabValorDistrib(IndReceber) = 0

Next
End Sub
Public Sub Rotina_020_Acumula_Receber()

If ctr!ctrStatus = 0 Then
   If ctr!ctrDataBanco < Date Then
      AcumReceberAtrasado = AcumReceberAtrasado + ctr!ctrValorDaBoleta
   Else
      If ctr!ctrDataBanco > (tabDataIni(1) - 1) And ctr!ctrDataBanco < (tabDataFim(1) + 2) And ctr!ctrDataVencito < (TabMesAnoFim(1) + 1) Then
         AcumPendReceber = AcumPendReceber + ctr!ctrValorDaBoleta
      Else
         If ctr!ctrDataBanco < (TabMesAnoFim(1) + 1) And ctr!ctrDataBanco > (tabDataIni(1) - 1) Then
            AcumReceberAteFim = AcumReceberAteFim + ctr!ctrValorDaBoleta
         End If
      End If
   End If
Else
   AcumProcesReceber = AcumProcesReceber + ctr!ctrValorDaBoleta
End If

End Sub

Public Sub Rotina_025_Acumula_Pagar()


If ctp!ctpStatus = 0 Then
   If ctp!chDataVencito < Date Then
      AcumPagarAtrasado = AcumPagarAtrasado + ctp!ctpValorDaBoleta
   Else
      If ctp!chDataVencito < tabDataFim(1) + 2 And ctp!chDataVencito < TabMesAnoFim(1) + 1 Then
         AcumPendPagar = AcumPendPagar + ctp!ctpValorDaBoleta
      Else
         If ctp!chDataVencito < TabMesAnoFim(1) + 1 Then
            AcumPagarAteFim = AcumPagarAteFim + ctp!ctpValorDaBoleta
         End If
      End If
   End If
Else
   AcumProcesPagar = AcumProcesPagar + ctp!ctpValorDaBoleta
End If

End Sub

Public Sub Rotina_026_Acumula_Mensal()

If ctr!ctrStatus = 0 Then
   If ctr!ctrDataVencito < Date Then
      AcumReceberAtrasadoMensal = AcumReceberAtrasadoMensal + ctr!ctrValorDaBoleta
   End If
   If Not ctr!ctrDataVencito < DiaUtilAnterior Then
      If ctr!ctrDataVencito > (tabDataIni(1) - 1) And ctr!ctrDataVencito < (tabDataFim(1) + 1) Then
         AcumPendReceberMensal = AcumPendReceberMensal * 1
      Else
         AcumPendReceberMensal = AcumPendReceberMensal + ctr!ctrValorDaBoleta
      End If
   End If
Else
   AcumProcesReceberMensal = AcumProcesReceberMensal + ctr!ctrValorDaBoleta
End If

End Sub
Public Sub Rotina_027_Acumula_Mensal_Pagar()

   If ctp!ctpdatabanco < Date Then  'DiaUtilAnterior + 1
      AcumPagarAtrasadoMensal = AcumPagarAtrasadoMensal + ctp!ctpValorDaBoleta
   Else
      AcumPendPagarMensal = AcumPendPagarMensal + ctp!ctpValorDaBoleta
   End If

End Sub

Public Sub Rotina_030_TabSemana()

fimsemana = 0

Do While fimsemana = 0

If tabDataFim(IndReceber) = Empty Then
   tabDataIni(IndReceber) = tabDataFim(IndReceber - 1) + 1
   tabDataFim(IndReceber) = tabDataIni(IndReceber) + 6
   tabValor(IndReceber, 0) = Format$(0#, "#,##0.00")
   tabValor(IndReceber, 1) = Format$(0#, "#,##0.00")
   tabValor(IndReceber, 2) = Format$(0#, "#,##0.00")
End If

If ctr!ctrDataVencito > dataInicio - 1 Then
      If ctr!ctrDataVencito > (tabDataFim(IndReceber) + 1) Then
         IndReceber = IndReceber + 1
      Else
         If ctr!ctrStatus = 0 Then ' And (ctr!ctrDataBanco > Date) Then   DiaUtilAnterior
            tabValor(IndReceber, 0) = Format$(tabValor(IndReceber, 0) + ctr!ctrValorDaBoleta, "#,##0.00")
         End If
         Call Rotina_020_Acumula_Receber
         tabValor(IndReceber, 2) = tabValor(IndReceber, 0) - tabValor(IndReceber, 1)
         IndReceber = 1
         fimsemana = 1
      End If
Else
   Call Rotina_020_Acumula_Receber
   fimsemana = 1
   IndReceber = 1
End If

Loop
End Sub

Public Sub Rotina_040_TabMensal()

fimmensal = 0

'MsgBox ("Nota Fiscal Fatura , , ctr!chnotafiscal  & " - " & ctr!chfatura

Do While fimmensal = 0

If ctr!ctrStatus = 0 Then
   DataMensal = ctr!ctrDataVencito
Else
   DataMensal = ctr!ctrDataRecebimento
End If

If DataMensal > (DataMensalInicio - 1) Then
   If TabMesAnoFim(IndMes) = Empty Then
         TabMesAnoIni(IndMes) = TabMesAnoFim(IndMes - 1) + 1
         ano = Year(TabMesAnoIni(IndMes))
         mes = Month(TabMesAnoIni(IndMes))
         mes = mes + 1
         If mes = 13 Then
            mes = 1
            ano = ano + 1
         End If
         Dia = 1
         DataParaCalculo = (Dia & "/" & mes & "/" & ano)
         UltimoMes = DataParaCalculo - 1
         TabMesAnoFim(IndMes) = DataParaCalculo - 1
         tabValorMensal(IndMes, 0) = Format$(0#, "#,##0.00")
         tabValorMensal(IndMes, 1) = Format$(0#, "#,##0.00")
         tabValorMensal(IndMes, 2) = Format$(0#, "#,##0.00")
         tabValorDistrib(IndMes) = Format$(0#, "#,##0.00")
       
    End If
End If

'Fiz alteracao nesta  rotina. if apos indmes = indmes + 1

If DataMensal > TabMesAnoFim(IndMes) Then
   IndMes = IndMes + 1
Else

   If ctr!ctrDataEmissao > (DataIniMesAtu - 1) Then
      tabValorDistrib(IndMes) = Format$(tabValorDistrib(IndMes) + ctr!ctrValorDaBoleta, "#,##0.00")
   End If
   If ctr!ctrStatus = 0 And DataMensal < (DiaUtilAnterior + 1) Then
      IndMes = IndMes
   Else
      tabValorMensal(IndMes, 0) = Format$(tabValorMensal(IndMes, 0) + ctr!ctrValorDaBoleta, "#,##0.00")
      Call Rotina_026_Acumula_Mensal
      tabValorMensal(IndMes, 2) = tabValorMensal(IndMes, 0) - tabValorMensal(IndMes, 1)
   End If
   
   IndMes = 1
   fimmensal = 1

End If

Loop

'Carga da tabela de vendas do mes

fimmensal = 0

Do While fimmensal = 0

If ctr!ctrDataEmissao > DataMensalInicio - 1 Then
   If ctr!ctrDataBanco > DataMensalInicio - 1 Then

      If TabMesFaturamento(IndMes) = Empty Then
         TabMesFaturamento(IndMes) = TabMesFaturamentoFim(IndMes - 1) + 1
         ano = Year(TabMesFaturamento(IndMes))
         mes = Month(TabMesFaturamento(IndMes))
         mes = mes + 1
         If mes = 13 Then
            mes = 1
            ano = ano + 1
         End If
         Dia = 1
         DataParaCalculo = (Dia & "/" & mes & "/" & ano)
       
         TabMesFaturamentoFim(IndMes) = DataParaCalculo - 1

         tabValorMensal(IndMes, 3) = Format$(0#, "#,##0.00")
         
       End If
   End If

   If ctr!ctrDataBanco > TabMesFaturamentoFim(IndMes) Then
      IndMes = IndMes + 1
   Else
      'If ctr!ctrstatus  = 0 And ctr!ctrdatavencito  < DiaUtilAnterior + 1 Then
      '   IndMes = IndMes
      'Else
         tabValorMensal(IndMes, 3) = Format$(tabValorMensal(IndMes, 3) + ctr!ctrValorDaBoleta, "#,##0.00")
         AcumFaturamentoMensal = Format$(AcumFaturamentoMensal + ctr!ctrValorDaBoleta, "#,##0.00")
         If ctr!ctrDataEmissao > DataIniMesAtu - 1 Then
            AcumFaturamentoDistrib = Format$(AcumFaturamentoDistrib + ctr!ctrValorDaBoleta, "#,##0.00")
         End If
      'End If
      IndMes = 1
      fimmensal = 1

   End If
Else
   IndMes = 1
   fimmensal = 1
End If
Loop

End Sub

Public Sub Rotina_050_CtaPagar_Semana()

fimsemana = 0

Do While fimsemana = 0

If tabDataFim(IndPagar) = Empty Then
   tabDataIni(IndPagar) = tabDataFim(IndPagar - 1) + 1
   tabDataFim(IndPagar) = tabDataIni(IndPagar) + 6
   tabValor(IndPagar, 0) = Format$(0#, "#,##0.00")
   tabValor(IndPagar, 1) = Format$(0#, "#,##0.00")
   tabValor(IndPagar, 2) = Format$(0#, "#,##0.00")
End If

If ctp!ctpStatus = 0 Or 2 Then
   DataMensal = ctp!chDataVencito
Else
   DataMensal = ctp!ctpDataPagamento
End If

If DataMensal > dataInicio - 1 Then
      If DataMensal > tabDataFim(IndPagar) + 1 Then
         IndPagar = IndPagar + 1
      Else
         If ctp!ctpStatus = 0 And ctp!chDataVencito > Date - 1 Then
            tabValor(IndPagar, 1) = Format$(tabValor(IndPagar, 1) + ctp!ctpValorDaBoleta, "#,##0.00")
         End If
         Call Rotina_025_Acumula_Pagar
         tabValor(IndPagar, 2) = tabValor(IndPagar, 0) - tabValor(IndPagar, 1)
         IndPagar = 1
         fimsemana = 1
      End If
   Else
      Call Rotina_025_Acumula_Pagar
      
      fimsemana = 1
      IndPagar = 1
   End If
Loop
End Sub

Public Sub Rotina_060_CtaPagar_Mes()

fimmensal = 0

Do While fimmensal = 0

   If (ctp!ctpStatus = 0) Or (ctp!ctpStatus = 2) Then
      DataMensal = ctp!chDataVencito
   Else
      DataMensal = ctp!ctpDataPagamento
   End If
   
   'If ctp!chdatavencito  > DataMensalInicio - 1 Then
   If DataMensal > DataMensalInicio - 1 Then
      If TabMesAnoFim(IndMes) = Empty Then
            TabMesAnoIni(IndMes) = TabMesAnoFim(IndMes - 1) + 1
            ano = Year(TabMesAnoIni(IndMes))
            mes = Month(TabMesAnoIni(IndMes))
            mes = mes + 1
            If mes = 13 Then
               mes = 1
               ano = ano + 1
            End If
            Dia = 1
            DataParaCalculo = (Dia & "/" & mes & "/" & ano)
            TabMesAnoFim(IndMes) = DataParaCalculo - 1
            UltimoMes = TabMesAnoFim(IndMes)
            tabValorMensal(IndMes, 0) = Format$(0#, "#,##0.00")
            tabValorMensal(IndMes, 1) = Format$(0#, "#,##0.00")
            tabValorMensal(IndMes, 2) = Format$(0#, "#,##0.00")
       End If
   End If
   
   'If ctp!chdatavencito  > TabMesAnoFim(IndMes) Then
   If DataMensal > TabMesAnoFim(IndMes) Then
      IndMes = IndMes + 1
   Else
      'If ctp!ctpstatus  = 0 And ctp!chdatavencito  < Date Then
      If ctp!ctpStatus = 0 And ctp!chDataVencito < Date Then
         IndMes = IndMes
      Else
         tabValorMensal(IndMes, 1) = Format$(tabValorMensal(IndMes, 1) + ctp!ctpValorDaBoleta, "#,##0.00")
      End If
      Call Rotina_027_Acumula_Mensal_Pagar
      tabValorMensal(IndMes, 2) = tabValorMensal(IndMes, 0) - tabValorMensal(IndMes, 1)
      IndMes = 1
      fimmensal = 1
   End If


Loop

End Sub

Public Sub Rotina_070_Ajusta_Data()

DiadaSemana = Weekday(DataInformada)

'Calcular Range de datas

dataInicio = DataInformada - (DiadaSemana)
dataFim = dataInicio + 6

Dia = 1
If Day(Date) < 15 Then
   mes = Month(dataFim)
   ano = Year(dataFim)
Else
   mes = Month(dataInicio)
   ano = Year(dataInicio)
End If

'Ano = Year(DataInicio)

DataMensalInicio = Dia & "/" & mes & "/" & ano

mes = mes + 1

If mes > 12 Then
   mes = 1
   ano = ano + 1
End If

DataMensalFim = (Dia & "/" & mes & "/" & ano)

DataMensalFim = DataMensalFim - 1

DataIniMesAtu = 1 & "/" & Month(Date) & "/" & Year(Date)

End Sub

Private Sub GridMensal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If GridMensal.Col = 1 Then
   TotalDebCre = txtTotalCreditoMensal
   GridMensal.TextMatrix(0, 4) = "%Crédito"
Else
   If GridMensal.Col = 2 Then
      TotalDebCre = txtTotalDebitoMensal
      GridMensal.TextMatrix(0, 4) = "%Débito"
   Else
      If GridMensal.Col = 3 Then
         TotalDebCre = txtSaldoMensal
         GridMensal.TextMatrix(0, 4) = "%Saldo"
      Else
         TotalDebCre = 0
         GridMensal.TextMatrix(0, 4) = "% Distr."
         If IndLimiteMensal > 0 Then
             For Linha = 1 To (IndLimiteMensal - 1)
                GridMensal.TextMatrix(Linha, 4) = Empty
             Next
         End If
      End If
   End If
End If

If TotalDebCre > 0 Then
   For Linha = 1 To (IndLimiteMensal - 1)
       ValorDebCre = GridMensal.TextMatrix(Linha, GridMensal.Col)
       PercDebCre = ((ValorDebCre * 100) / TotalDebCre) / 100
       GridMensal.TextMatrix(Linha, 4) = Format$((PercDebCre) * 100, "##0.00") & "%"
   Next
End If
End Sub


Private Sub GridSemana_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim DataAnaliticoDe As Date
Dim DataAnaliticoAte As Date

If GridSemana.Row = 0 Then
   Linha = 1
Else
   Linha = GridSemana.Row
End If

DataAnaliticoDe = GridSemana.TextMatrix(Linha, 0)
DataAnaliticoAte = GridSemana.TextMatrix(Linha, 1)
txtDataCtaPagDe = Format$(GridSemana.TextMatrix(Linha, 0), "dd/mm/yyyy")
txtDataCtaPagAte = Format$(GridSemana.TextMatrix(Linha, 1), "dd/mm/yyyy")

DataAnaliticoDe = DataAnaliticoDe - 1
DataAnaliticoAte = DataAnaliticoAte + 1

GrdCtaReceber.Rows = 2
GrdCtaReceber.TextMatrix(1, 0) = Empty
GrdCtaReceber.TextMatrix(1, 1) = Empty
GrdCtaReceber.TextMatrix(1, 2) = Empty
GrdCtaReceber.TextMatrix(1, 3) = Empty

GrdCtaPagar.Rows = 2
GrdCtaPagar.TextMatrix(1, 0) = Empty
GrdCtaPagar.TextMatrix(1, 1) = Empty
GrdCtaPagar.TextMatrix(1, 2) = Empty
GrdCtaPagar.TextMatrix(1, 3) = Empty

Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber", db, 3, 3
If ctr.EOF Then
   'MsgBox ("Não há lançamentos a Crédito até a apresente data."), vbInformation
   Fim = 1
Else
   ctr.MoveFirst
   Fim = 0
End If

LinhaAnalitico = 0

Do While Fim = 0
   If (ctr!ctrDataVencito > DataAnaliticoDe And ctr!ctrDataVencito < DataAnaliticoAte) And ctr!ctrStatus = 0 Then
      LinhaAnalitico = LinhaAnalitico + 1
      GrdCtaReceber.Rows = LinhaAnalitico + 1
      ChaveClassifica = Year(ctr!ctrDataVencito) & Month(ctr!ctrDataVencito) & Day(ctr!ctrDataVencito) & ctr!chPessoa
      GrdCtaReceber.TextMatrix(LinhaAnalitico, 0) = ChaveClassifica
      GrdCtaReceber.TextMatrix(LinhaAnalitico, 1) = Format$(ctr!ctrDataVencito, "dd/mm/yy")
      GrdCtaReceber.TextMatrix(LinhaAnalitico, 2) = ctr!chPessoa & "-" & ctr!chNotafiscal
      GrdCtaReceber.TextMatrix(LinhaAnalitico, 3) = Format$(ctr!ctrValorDaBoleta, "#,##0.00")
   End If
   ctr.MoveNext
   If ctr.EOF Then
      FimReceber = LinhaAnalitico
      Fim = 1
   End If
Loop


ctp.Open "Select * from contas_a_pagar", db, 3, 3
If ctp.EOF Then
   'MsgBox ("Não há lançamentos a Débito até a apresente data."), vbInformation
   Fim = 1
Else
   ctp.MoveFirst
   Fim = 0
End If

LinhaAnalitico = 0

Do While Fim = 0
   If (ctp!chDataVencito > DataAnaliticoDe And ctp!chDataVencito < DataAnaliticoAte) And ctp!ctpStatus = 0 Then
      LinhaAnalitico = LinhaAnalitico + 1
      GrdCtaPagar.Rows = LinhaAnalitico + 1
      ChaveClassifica = Format$(ctp!chDataVencito, "yy/mm/dd")
      ChaveClassifica = ChaveClassifica & ctp!ctpdescricaooperacao
      GrdCtaPagar.TextMatrix(LinhaAnalitico, 0) = ChaveClassifica
      GrdCtaPagar.TextMatrix(LinhaAnalitico, 1) = Format$(ctp!chDataVencito, "dd/mm/yy")
      GrdCtaPagar.TextMatrix(LinhaAnalitico, 2) = ctp!ctpdescricaooperacao
      GrdCtaPagar.TextMatrix(LinhaAnalitico, 3) = Format$(ctp!ctpValorDaBoleta, "#,##0.00")
   End If
   ctp.MoveNext
   If ctp.EOF Then
      FimPagar = LinhaAnalitico
      Fim = 1
   End If
Loop

GrdCtaReceber.Col = 0
GrdCtaReceber.ColSel = 0
GrdCtaReceber.Row = 1
GrdCtaReceber.RowSel = FimReceber
GrdCtaReceber.Sort = 1
GrdCtaReceber.Col = 0
GrdCtaReceber.ColSel = 0
GrdCtaReceber.Row = 0
GrdCtaReceber.RowSel = 0

GrdCtaPagar.Col = 0
GrdCtaPagar.ColSel = 0
GrdCtaPagar.Row = 1
GrdCtaPagar.RowSel = FimPagar
GrdCtaPagar.Sort = 7
GrdCtaPagar.Col = 0
GrdCtaPagar.ColSel = 0
GrdCtaPagar.Row = 0
GrdCtaPagar.RowSel = 0
   
End Sub


