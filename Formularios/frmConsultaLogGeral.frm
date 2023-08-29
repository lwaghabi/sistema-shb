VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaLogGeral 
   Caption         =   "frmConsultaLogGeral"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
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
      Height          =   495
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
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
      Left            =   17760
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid grdLogistica 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   12515
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Label txtMesRef 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Mês de Referência"
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
      Left            =   11280
      TabIndex        =   4
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
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
      Left            =   17760
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Consulta Logística Geral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
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
      Width           =   8775
   End
End
Attribute VB_Name = "frmConsultaLogGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CabecalhoDias As String
Dim Ind As Integer
Dim IndLinha As Integer
Dim IndCol As Integer
Dim ColaboradorAnterior As String
Dim PessoaAnterior As String
Dim UnidadeOperacionalAnterior As String

'Datas

Dim InicioDataChLeitura As Date
Dim FinalDataChLeitura As Date
Dim DataInicioInvertida As String
Dim DataFinalInvertida As String
Dim DataHoje As Date
Dim AnoMesRef As String

Dim Ano As Integer
Dim Mes As Integer
Dim Dia As Integer

Dim MesProximo As Integer
Dim NumDiasMes As Integer


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

txtHoje = Date

txtMesRef = Format$(txtHoje, "mmm-yyyy")


Call GeraGridlogistica

Call CarregaGrid

End Sub

Public Sub GeraGridlogistica()

Call GeraDataInicioDataFim
grdLogistica.Cols = NumDiasMes + 5

Ind = 1

CabecalhoDias = Empty

Do While Ind < NumDiasMes + 1
   CabecalhoDias = CabecalhoDias & ("|" & Format$(Ind, "00"))
   Ind = Ind + 1
Loop

grdLogistica.FormatString = "Cliente                                     " & "|" & "UnidadeOperacional          " & "|" & "Colaborador                                         " & "|" & "Evento                                  " & "|" & "Inicio Evento" & "|" & "Final Evento" & CabecalhoDias

If NumDiasMes = 28 Then
   txtMesRef.Width = 9190
Else
   If NumDiasMes = 29 Then
      txtMesRef.Width = 99550
   Else
      If NumDiasMes = 30 Then
         txtMesRef.Width = 9940
      Else
         If NumDiasMes = 31 Then
            txtMesRef.Width = 10250
         End If
      End If
   End If
End If

End Sub

Public Sub GeraDataInicioDataFim()

Ano = Year(txtHoje)
Mes = Month(txtHoje)
Dia = Day(txtHoje)

DataInicioInvertida = Format$(Ano & "-" & Mes & "-" & "01", "yyyy-mm-dd")
InicioDataChLeitura = Format$("01" & "-" & Mes & "-" & Ano, "dd-mm-yyyy")

MesProximo = Format$(Mes, "00")
DataHoje = txtHoje
Do While Mes = MesProximo
   DataHoje = DataHoje + 1
   MesProximo = Format$(Month(DataHoje), "00")
Loop

DataHoje = DataHoje - 1

DataFinalInvertida = Format$(DataHoje + 1, "yyyy-mm-dd")
FinalDataChLeitura = Format$(DataHoje + 1, "dd-mm-yyyy")
AnoMesRef = Year(InicioDataChLeitura) & Format$(Month(InicioDataChLeitura), "00")

NumDiasMes = FinalDataChLeitura - InicioDataChLeitura

End Sub

Public Sub CarregaGrid()

Call LimpaGrid

Call Rotina_AbrirBanco

lgt.Open "Select * from logistica where chAnoMesRef = ('" & AnoMesRef & "')", db, 3, 3
If lgt.EOF Then
   Call FechaDB
   Exit Sub
End If

lgt.MoveFirst

IndLinha = 1
IndCol = 0
ColaboradorAnterior = Empty

Do While Not lgt.EOF

   grdLogistica.Rows = IndLinha + 1
   If Not lgt!chPessoa = PessoaAnterior Then
      grdLogistica.TextMatrix(IndLinha, 0) = lgt!chPessoa
      PessoaAnterior = lgt!chPessoa
   Else
      grdLogistica.TextMatrix(IndLinha, 0) = Empty
   End If
   If Not lgt!chUnidadeOperacional = UnidadeOperacionalAnterior Then
      grdLogistica.TextMatrix(IndLinha, 1) = lgt!chUnidadeOperacional
      UnidadeOperacionalAnterior = lgt!chUnidadeOperacional
   Else
      grdLogistica.TextMatrix(IndLinha, 1) = Empty
   End If
   
   If Not lgt!chColaborador = ColaboradorAnterior Then
      grdLogistica.TextMatrix(IndLinha, 2) = lgt!chColaborador
      ColaboradorAnterior = lgt!chColaborador
   Else
      grdLogistica.TextMatrix(IndLinha, 2) = Empty
   End If

   grdLogistica.TextMatrix(IndLinha, 3) = lgt!chEvento
   grdLogistica.TextMatrix(IndLinha, 4) = lgt!chInicioEvento
   grdLogistica.TextMatrix(IndLinha, 5) = lgt!lgtFimEvento
   
   For IndCol = Day(lgt!chInicioEvento) To Day(lgt!lgtFimEvento)
       grdLogistica.TextMatrix(IndLinha, (IndCol + 5)) = lgt!lgtCodEvento
       grdLogistica.Col = (IndCol + 5)
       grdLogistica.Row = IndLinha

       If lgt!lgtCodEvento = "H" Or lgt!lgtCodEvento = "HTL" Then
          grdLogistica.CellBackColor = vbGreen
       Else
          If lgt!lgtCodEvento = "E" Or lgt!lgtCodEvento = "EM" Or lgt!lgtCodEvento = "EMB" Then
            grdLogistica.CellBackColor = vbRed
          Else
             If lgt!lgtCodEvento = "R" Or lgt!lgtCodEvento = "FLG" Or lgt!lgtCodEvento = "FOL" Or lgt!lgtCodEvento = "RR" Then
                grdLogistica.CellBackColor = vbBlue
             Else
                If lgt!lgtCodEvento = "F" Or lgt!lgtCodEvento = "FER" Then
                   grdLogistica.CellBackColor = vbYellow
                Else
                   grdLogistica.TextMatrix(IndLinha, (IndCol + 4)) = lgt!lgtCodEvento
                   grdLogistica.CellBackColor = &HFF00FF
                End If
             End If
           End If
        End If
        
       grdLogistica.CellForeColor = &H80000004
   Next
   
   IndLinha = IndLinha + 1
   
   lgt.MoveNext

Loop
 
End Sub
Public Sub LimpaGrid()

   grdLogistica.Rows = 2

    grdLogistica.TextMatrix(1, 0) = Empty
    grdLogistica.TextMatrix(1, 1) = Empty
    
       If grdLogistica.Cols > 2 Then
       grdLogistica.TextMatrix(1, 2) = Empty
       grdLogistica.TextMatrix(1, 3) = Empty
       grdLogistica.TextMatrix(1, 4) = Empty
       grdLogistica.TextMatrix(1, 5) = Empty

       For Ind = 6 To NumDiasMes + 4
           grdLogistica.TextMatrix(1, Ind) = Empty
           grdLogistica.Col = Ind
           grdLogistica.Row = 1
           grdLogistica.CellBackColor = Empty
       Next
    End If
End Sub
