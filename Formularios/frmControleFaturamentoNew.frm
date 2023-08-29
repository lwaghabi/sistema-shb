VERSION 5.00
Begin VB.Form frmControleFaturamentoNew 
   Caption         =   "Controle de Faturamento- (frmControleFaturamentoNew)"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Controle de Faturamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox cmbTipoProces 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cmbPedido 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cmbDataNeg 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Imprime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Data"
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
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Nota Fiscal "
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
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label txtCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmControleFaturamentoNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fim As Byte
Dim DataHoje As Date
Dim DiaHoje As Integer
Dim MesHoje As Integer
Dim AnoHoje As Integer

Dim DataCombo As Date
Dim DiaCombo As Integer
Dim Indice As Byte
Dim GeradorCntrl As Byte
Dim Resp As String
Dim FlagAtu As Byte

Private Sub cmbTipoProces_LostFocus()
Fim = 0

cmbPedido.Clear

cmbPedido.AddItem " Geral"

TabNegociacao.MoveFirst

Do While Fim = 0
   If cmbTipoProces.ListIndex = 0 Then
      If (TabNegociacao("negstatus") = 1 And TabNegociacao("negcntrlfaturamento") = 0) Or (TabNegociacao("negstatus") = 3 And TabNegociacao("negcntrlfaturamento") = 0) Then
         cmbPedido.AddItem TabNegociacao("negnotafiscal")
      End If
   Else
      If (TabNegociacao("negstatus") = 1 And TabNegociacao("negcntrlfaturamento") = 1) Or (TabNegociacao("negstatus") = 3 And TabNegociacao("negcntrlfaturamento") = 1) Then
         cmbPedido.AddItem TabNegociacao("negnotafiscal")
      End If
   End If
   TabNegociacao.MoveNext
   If TabNegociacao.EOF Then
      Fim = 1
   End If
Loop
cmbPedido.ListIndex = 0
End Sub

Private Sub Command1_Click()
Fim = 0

If cmbPedido = Empty Then
   MsgBox "Nota Fiscal ou Geral não Informado. Posicionar o Cursor em Tipo e clicar em Tab"""
   Command2.SetFocus
   Exit Sub
End If

If cmbPedido.ListIndex = 0 Then
   TabNegociacao.MoveFirst
   Do While Fim = 0
      If (TabNegociacao("negstatus") = 1 And TabNegociacao("negcntrlfaturamento") = 0) Or (TabNegociacao("negstatus") = 3 And TabNegociacao("negcntrlfaturamento") = 0) Then
         TabNegociacao.Edit
         TabNegociacao("negcntrlfaturamento") = 3
         TabNegociacao.Update
       End If
       TabNegociacao.MoveNext
    
       If TabNegociacao.EOF Then
          Fim = 1
       End If
   Loop
Else
   TabNegNF.Seek "=", cmbPedido
   If TabNegNF.NoMatch Then
      MsgBox ("Negociacao nao encontrada")
      Fim = Fim / 0
   End If

   TabNegNF.Edit
   TabNegNF("negcntrlfaturamento") = 3
   TabNegNF.Update
End If

Resp = MsgBox("Impressão Solicitada. Confirma???", vbYesNo)
If Resp = vbYes Then
   impControleFaturamentoNew.Show vbModal
  
   deControleFaturamento.rscmdControleFaturamento.Close

   Resp = MsgBox("Impressão Realizada com Exito???", vbYesNo)
   If Resp = vbYes Then
      FlagAtu = 1
      Fim = 0
   Else
      MsgBox ("Atenção: Atualização de Impressão não Processada")
      FlagAtu = 0
      Fim = 1
   End If
Else
   MsgBox ("Impressão Abortada")
   FlagAtu = 0
   Fim = 1
End If



TabNegociacao.MoveFirst

Do While Not TabNegociacao.EOF
   If TabNegociacao("negcntrlfaturamento") = 3 Then
      TabNegociacao.Edit
      TabNegociacao("negcntrlfaturamento") = FlagAtu
      TabNegociacao.Update
   End If
   TabNegociacao.MoveNext
Loop

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
cmbTipoProces.AddItem "Impressão"
cmbTipoProces.AddItem "Reimpressão"

cmbTipoProces.ListIndex = 0
GeradorCntrl = 0

DataHoje = Date

DiaHoje = Day(DataHoje)
DiaCombo = 1
cmbDataNeg.AddItem " Geral"
For DiaCombo = 1 To DiaHoje
    DataCombo = DiaCombo & "/" & Month(DataHoje) & "/" & Year(DataHoje)
    cmbDataNeg.AddItem DataCombo
    DiaCombo = DiaCombo
Next

cmbDataNeg.ListIndex = 0

End Sub




