VERSION 5.00
Begin VB.Form frmDeliveryOrder 
   Caption         =   "frmDeliveryOrder"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEquipOperDisp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
   End
   Begin VB.ComboBox cmbNumOS 
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cmbCliente 
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Equipamentos/Operadores Disponíveis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Número da OS"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Registro e Atualização Delivery Order "
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
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmDeliveryOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCliente_Change()
Call Rotina_AbrirBanco

rs.Open "Select numOS from ordemservico where cliente = ('" & cmbCliente & "')", db, 3, 3

If rs.EOF Then

      MsgBox ("Nenhuma OS foi registrada para este cliente")
      Call FechaDB
      Exit Sub

End If

rs.MoveFirst

Do While Not rs.EOF

   cmbNumOS.AddItem rs!NumOS
   rs.MoveNext
   
Loop

rs.Close
Call FechaDB
End Sub

Private Sub Form_Load()
Call Rotina_AbrirBanco

rs.Open "Select cliente from ordemservico where status = 0", db, 3, 3

If rs.EOF Then

      MsgBox ("Nenhuma OS foi registrada")
      Call FechaDB
      Exit Sub

End If

rs.MoveFirst

Do While Not rs.EOF

   cmbCliente.AddItem rs!Cliente
   rs.MoveNext
   
Loop

rs.Close
Call FechaDB
End Sub
