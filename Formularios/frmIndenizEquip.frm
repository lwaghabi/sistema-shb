VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIndenizEquip 
   Caption         =   "frmIndenizEquip"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   15810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
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
      Left            =   13440
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
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
      Left            =   11640
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
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
      Left            =   9360
      MaskColor       =   &H00808080&
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid grdIndenizacao 
      Height          =   3255
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   5741
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Descrição do Equipamento                                                                                   |Valor indenização "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtValorIndenizacao 
      Alignment       =   1  'Right Justify
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
      Left            =   6960
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtEquipamento 
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
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   10
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Valor da Indenização"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Descrição do Equipamento"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Registro e Atualização de Indenizações de Equipamentos Locados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
   End
End
Attribute VB_Name = "frmIndenizEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ind As Integer
Dim Coluna As Integer
Dim Linha As Integer
Dim ValorTotal As Currency

Private Sub cmdExcluir_Click()

Dim Resp As String

Call Rotina_AbrirBanco

rs.Open "Select * from IndenizEquip where descEquip = ('" & txtEquipamento & "')", db, 3, 3
If rs.EOF Then
   MsgBox ("Comando para exclusão de registro inexistente."), vbCritical
   Call FechaDB
   Exit Sub
End If

Resp = MsgBox("Exclusão de registro. Confirma???", vbExclamation + vbYesNo)
If Resp = vbYes Then
   rs.Delete
   MsgBox ("Registro excluído com sucesso."), vbInformation
End If

Call FechaDB

Call CarregaGrid

Call GerarExcelWord

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()

Call Rotina_AbrirBanco

If txtEquipamento = Empty Then
   MsgBox ("Equipamento não informado"), vbInformation
   Call FechaDB
   Exit Sub
End If

If txtValorIndenizacao = Empty Then
   MsgBox ("Valor da indenização não informado."), vbInformation
   Call FechaDB
   Exit Sub
End If
   
rs.Open "Select * from IndenizEquip WHERE descEquip = ('" & txtEquipamento & "')", db, 3, 3
If rs.EOF Then
   rs.AddNew
End If

rs!descEquip = txtEquipamento
rs!valor = txtValorIndenizacao

rs.Update

Call FechaDB

grdIndenizacao.Rows = 1

Call CarregaGrid

Call GerarExcelWord

Call FechaDB
   
End Sub

Private Sub Form_Load()

Call CarregaGrid

Call FechaDB

End Sub

Public Sub CarregaGrid()

Call Rotina_AbrirBanco

rs.Open "Select * from IndenizEquip", db, 3, 3
If Not rs.EOF Then
   Ind = 1
   rs.MoveFirst
End If

ValorTotal = 0

Do While Not rs.EOF
   grdIndenizacao.Rows = Ind + 1
   grdIndenizacao.TextMatrix(Ind, 0) = rs!descEquip
   grdIndenizacao.TextMatrix(Ind, 1) = Format$(rs!valor, "##,##0.00")
   ValorTotal = ValorTotal + rs!valor
   Ind = Ind + 1
   rs.MoveNext
Loop

lblTotal = Format$(ValorTotal, "##,###,##0.00")
txtEquipamento = Empty
txtValorIndenizacao = Empty
End Sub

Private Sub grdIndenizacao_Click()
Coluna = grdIndenizacao.Col
Linha = grdIndenizacao.Row

If Linha > grdIndenizacao.Rows Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

If grdIndenizacao.TextMatrix(Linha, 1) = Empty Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

txtEquipamento = grdIndenizacao.TextMatrix(Linha, 0)
txtValorIndenizacao = grdIndenizacao.TextMatrix(Linha, 1)

End Sub

Public Sub GerarExcelWord()

        Dim CaminhoNew As String
                
        CaminhoNew = "C:\Meus Documentos\SISTEMA SHB" & "\docPadrao\"
        
        Dim oApp As Excel.Application
        Dim oWB As Excel.Workbook
        Dim i As Integer
        Dim Ex As Object
        Set Ex = CreateObject("Excel.Application")

        i = 3
         On Error GoTo Erro
            'Create an Excel instance.
50          Set oApp = New Excel.Application

            'Open the desired workbook

60          If Dir(CaminhoNew & "ModelExcelWord.xlsx", vbArchive) = "" Then
70             MsgBox "Não foi possível gerar o documento porque" & vbCrLf & _
               "O arquivo padrão não foi localizado!", vbCritical
80             Exit Sub
90          End If
            
100         Set oWB = oApp.Workbooks.Open(FileName:=CaminhoNew & "ModelExcelWord.xlsx")
            
            'Do any modifications to the workbook.
            Rotina_AbrirBanco
            rs.Open "SELECT * FROM IndenizEquip", db, 3, 3
            Do Until rs.EOF
               oApp.Cells(i, 1) = rs!descEquip
               oApp.Cells(i, 2) = rs!valor
               rs.MoveNext
               i = i + 1
            Loop
110
          FechaDB

490       oWB.SaveAs FileName:=CaminhoNew & "ExcelWord.xlsx"

510       oWB.Close SaveChanges:=False
520       Set oWB = Nothing
530       oApp.Quit
540       Set oApp = Nothing

400       Ex.Workbooks.Open (CaminhoNew & "ExcelWord.xlsx")
410       Ex.Visible = True

Exit Sub
Erro:
MsgBox "Ocorreu o seguinte erro:" + vbNewLine & _
        Err.Description + vbNewLine & "Na linha: " & Erl
End Sub


