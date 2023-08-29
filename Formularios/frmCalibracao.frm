VERSION 5.00
Begin VB.Form frmCalibracao 
   Caption         =   "frmCalibracao"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbStatus 
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
      Left            =   9720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtEquipamento 
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
      Left            =   4320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   5175
   End
   Begin VB.ComboBox cmbCodEquipamento 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Status"
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
      Left            =   9840
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Equipamento"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Cod. Equipamento"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Registro e Atualização de Equipamentos em Calibração"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmCalibracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
