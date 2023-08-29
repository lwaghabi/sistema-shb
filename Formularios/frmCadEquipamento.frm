VERSION 5.00
Begin VB.Form frmCadEquipamento 
   Caption         =   "frmCadEquipamento"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   16020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDenominacao 
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
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox txtCodEquipamento 
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
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Denominação(curto)"
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
      Caption         =   "Código do Equipamento"
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
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "REGISTRO E ATUALIZAÇÃO DE EQUIPAMENTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmCadEquipamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
