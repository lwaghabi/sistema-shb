VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaFaturamento 
   Caption         =   "frmConsultaFaturamento"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   16500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   16215
      Begin MSFlexGridLib.MSFlexGrid GridFatura 
         Height          =   4695
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         FormatString    =   $"frmConsultaFaturamento.frx":0000
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
   Begin VB.Frame Frame1 
      Caption         =   "Parâmetros de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   12855
      Begin VB.ComboBox cmbCliente 
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
         TabIndex        =   0
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cmbAnoFim 
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
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbMesFim 
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
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
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
         Height          =   615
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H0000FF00&
         Caption         =   "Consulta"
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
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbAno 
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
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbMes 
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
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
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Ano Final"
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
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Mês Final"
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
         Left            =   6360
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Ano Início"
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
         Left            =   4920
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Mês Início"
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
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label lblHoje 
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
      Left            =   13800
      TabIndex        =   16
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
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
      Left            =   13800
      TabIndex        =   15
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Total Faturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13800
      TabIndex        =   14
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblTotalFaturado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   13320
      TabIndex        =   13
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Consulta Faturamento Por Período"
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
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmConsultaFaturamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ind As Integer
Dim ano As Integer
Dim mes As Integer
Dim Dia As Integer
Dim AnoInicioOperacao
Dim DataHoje As Date
Dim AnoHoje As Integer
Dim MesHoje As Integer
Dim DataInicioInvertida As String
Dim DataFimInvertida As String
Dim DataInvertida As String
Dim DataInicioOperacao As Date
Dim DataFimOperacao As Date
Dim DataFinalInvertida As String
Dim MesProximo As Integer
Dim Periodo As Integer
Dim ValorAcumulado As Currency
Dim TotalRecebido As Currency
Dim TotalReceber As Currency
Dim TotalEmAtraso As Currency
Dim Resp As String
Dim Historico As Integer
Dim Status As Integer
Dim Limite As Integer
Dim CargaAnterior As Integer
Dim TipoDeLancamento As String

Dim linhaInicio As Integer
Dim colunaInicio As Integer
Dim linhaFim As Integer
Dim colunaFim As Integer


Dim dataInicio As String
Dim dataFim As String
Dim AcumulaFatura As Currency
Dim TotalFatura As Currency

Dim ClienteAnterior As String

Private Sub cmdConsulta_Click()

GridFatura.Rows = 2

GridFatura.TextMatrix(1, 0) = Empty
GridFatura.TextMatrix(1, 1) = Empty
GridFatura.TextMatrix(1, 2) = Empty
GridFatura.TextMatrix(1, 3) = Empty
GridFatura.TextMatrix(1, 4) = Empty
GridFatura.TextMatrix(1, 5) = Empty
GridFatura.TextMatrix(1, 6) = Empty
GridFatura.TextMatrix(1, 7) = Empty
GridFatura.TextMatrix(1, 8) = Empty
GridFatura.TextMatrix(1, 9) = Empty
GridFatura.TextMatrix(1, 10) = Empty


ValorAcumulado = 0
CargaAnterior = 0

If AnoHoje = cmbAno Then
   If cmbMes > MesHoje Then
      MsgBox ("Mês para consulta inválido. Maior que o mês da data atual."), vbInformation
      Exit Sub
   End If
End If

If AnoHoje = cmbAnoFim Then
   If cmbMesFim > MesHoje Then
      MsgBox ("Mês final para consulta inválido. Maior que o mês da data atual."), vbInformation
      Exit Sub
   End If
End If

If cmbAno = cmbAnoFim Then
   If cmbMes > cmbMesFim Then
      MsgBox ("Mês final para consulta inválido. Menor que o mês de início da pesquisa."), vbInformation
      Exit Sub
   End If
End If
 
If cmbAnoFim < cmbAno Then
   MsgBox ("Ano inicio não pode ser menor que o ano fim da pesquisa."), vbInformation
   Exit Sub
End If

Call CriaDatasPesquisa

Call Rotina_AbrirBanco

If cmbCliente = " Todos" Then
   ctr.Open "Select * from contas_a_receber where ctrDataEmissao > ('" & DataInicioInvertida & "') and ctrDataEmissao < ('" & DataFinalInvertida & "') and chFatura = ('" & 1 & "')", db, 3, 3
   If ctr.EOF Then
      CargaAnterior = 0
      AcumulaFatura = 0
      TotalFatura = 0
      Ind = 1
   Else
      CargaAnterior = 1
   End If
Else
   ctr.Open "Select * from contas_a_receber where chPessoa = ('" & cmbCliente & "') and ctrDataEmissao > ('" & DataInicioInvertida & "') and ctrDataEmissao < ('" & DataFinalInvertida & "') and chFatura = ('" & 1 & "')", db, 3, 3
   If ctr.EOF Then
      CargaAnterior = 0
      AcumulaFatura = 0
      TotalFatura = 0
      Ind = 1
   Else
      CargaAnterior = 1
   End If
End If

If CargaAnterior = 1 Then
   Call ProcessaConsulta
End If

If cmbCliente = " Todos" Then
   If ctr.State = 1 Then
      ctr.Close: Set ctr = Nothing
   End If
   
   ctr.Open "Select * from historicocontasreceber where ctrDataEmissao > ('" & DataInicioInvertida & "') and ctrDataEmissao < ('" & DataFinalInvertida & "')", db, 3, 3
   
   If ctr.EOF Then
      Periodo = 0
   Else
      Periodo = 1
   End If
Else
   If ctr.State = 1 Then
      ctr.Close: Set ctr = Nothing
   End If
   ctr.Open "Select * from historicocontasreceber where chPessoa = ('" & cmbCliente & "') and ctrDataEmissao > ('" & DataInicioInvertida & "') and ctrDataEmissao < ('" & DataFinalInvertida & "')", db, 3, 3
   
   If ctr.EOF Then
      Periodo = 0
   Else
      Periodo = 1
   End If

End If

If Periodo = 1 Then
   Call ProcessaConsulta
End If

Call FechaDB

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

AnoHoje = Year(Date)
MesHoje = Month(Date)
lblHoje = Date

AnoInicioOperacao = 2019

ano = Year(Date)

For Ind = 1 To 12
    cmbMes.AddItem Format$(Ind, "00")
    cmbMesFim.AddItem Format$(Ind, "00")
Next

cmbMes.ListIndex = 0
cmbMesFim.ListIndex = 0

Do While (ano + 1) > AnoInicioOperacao
   cmbAno.AddItem ano
   cmbAnoFim.AddItem ano
   ano = ano - 1
Loop

cmbAno.ListIndex = 0
cmbAnoFim.ListIndex = 0

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where pesTipoPessoa = ('" & 0 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("ERRO; Acesso a pessoa sem Clientes"), vbInformation
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

cmbCliente.AddItem " Todos"

Do While Not pes.EOF

   cmbCliente.AddItem pes!chPessoa
   pes.MoveNext
   
Loop

cmbCliente.ListIndex = 0

End Sub

Public Sub ProcessaConsulta()

If CargaAnterior = 1 Then
   Ind = 1
   CargaAnterior = 0
   AcumulaFatura = 0
   TotalFatura = 0
End If

Do While Not ctr.EOF
   GridFatura.Rows = Ind + 1
   
   If Not ctr!chPessoa = ClienteAnterior Then
      AcumulaFatura = 0
      ClienteAnterior = ctr!chPessoa
   End If
   
   DataInvertida = Format$(ctr!ctrDataEmissao, "yyyy-mm-dd")
   
   GridFatura.TextMatrix(Ind, 10) = ctr!chPessoa & DataInvertida
   GridFatura.TextMatrix(Ind, 0) = ctr!chPessoa
   
   If GridFatura.TextMatrix(Ind, 0) = GridFatura.TextMatrix((Ind - 1), 0) Then
      GridFatura.TextMatrix(Ind, 1) = Empty
   Else
      GridFatura.TextMatrix(Ind, 1) = ctr!chPessoa
   End If
   
   GridFatura.TextMatrix(Ind, 2) = ctr!chNotafiscal
   GridFatura.TextMatrix(Ind, 3) = ctr!ctrDataEmissao
   GridFatura.TextMatrix(Ind, 4) = ctr!chNumPedido
   GridFatura.TextMatrix(Ind, 5) = ctr!chNumPedidoComp
   GridFatura.TextMatrix(Ind, 6) = ctr!ctrDataVencito
   GridFatura.TextMatrix(Ind, 7) = Format$(ctr!ctrValorDaBoleta, "##,##0.00")
   
   If ctr!ctrStatus = 0 Then
      If ctr!ctrDataVencito > Date Then
         
         GridFatura.TextMatrix(Ind, 8) = "Pendente"
      Else
         GridFatura.TextMatrix(Ind, 8) = "Atrasado"
      End If
   Else
      GridFatura.TextMatrix(Ind, 8) = "Pago"
   End If
   
   AcumulaFatura = AcumulaFatura + ctr!ctrValorDaBoleta
   TotalFatura = TotalFatura + ctr!ctrValorDaBoleta

   Ind = Ind + 1
   
   ctr.MoveNext
Loop

Limite = Ind
ClienteAnterior = Empty
TotalFatura = 0
AcumulaFatura = 0

GridFatura.Col = 10
GridFatura.ColSel = 10
GridFatura.Sort = 7

For Ind = 1 To Limite - 1
    If GridFatura.TextMatrix(Ind, 0) = ClienteAnterior Then
       GridFatura.TextMatrix(Ind, 1) = Empty
       GridFatura.TextMatrix(Ind, 9) = Empty
       AcumulaFatura = AcumulaFatura + GridFatura.TextMatrix(Ind, 7)
       TotalFatura = TotalFatura + GridFatura.TextMatrix(Ind, 7)
    Else
       If Ind > 1 Then
          GridFatura.TextMatrix(Ind - 1, 9) = Format$(AcumulaFatura, "##,##0.00")
          AcumulaFatura = GridFatura.TextMatrix(Ind, 7)
          TotalFatura = TotalFatura + GridFatura.TextMatrix(Ind, 7)
          GridFatura.TextMatrix(Ind, 1) = GridFatura.TextMatrix(Ind, 0)
       Else
          AcumulaFatura = GridFatura.TextMatrix(Ind, 7)
          TotalFatura = TotalFatura + GridFatura.TextMatrix(Ind, 7)
          GridFatura.TextMatrix(Ind, 1) = GridFatura.TextMatrix(Ind, 0)
       End If
    End If
    ClienteAnterior = GridFatura.TextMatrix(Ind, 0)
Next

GridFatura.TextMatrix((Ind - 1), 9) = Format$(AcumulaFatura, "##,#0.00")
lblTotalFaturado = Format$(TotalFatura, "##,#0.00")

End Sub

Public Sub CriaDatasPesquisa()

mes = Format$(cmbMes, "00")
ano = cmbAno
Dia = Format$(1, "00")

DataHoje = Format$(Dia, "00") & "/" & Format$(mes, "00") & "/" & ano
DataInicioOperacao = DataHoje
DataHoje = DataHoje - 1
Dia = Day(DataHoje)
mes = Month(DataHoje)
ano = Year(DataHoje)

DataInicioInvertida = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")

DataHoje = DataHoje + 1
MesProximo = Month(DataHoje)

Do While Month(DataHoje) = MesProximo
   DataHoje = DataHoje + 1
   'MesProximo = Format$(Month(DataHoje), "00")
Loop

mes = Format$(cmbMesFim, "00")
ano = cmbAnoFim
Dia = Format$(1, "00")

DataHoje = Format$(Dia, "00") & "/" & Format$(mes, "00") & "/" & ano
DataFimOperacao = DataHoje
DataHoje = DataHoje - 1
Dia = Day(DataHoje)
mes = Month(DataHoje)
ano = Year(DataHoje)

DataFimInvertida = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")

DataHoje = DataHoje + 1
MesProximo = Month(DataHoje)

Do While Month(DataHoje) = MesProximo
   DataHoje = DataHoje + 1
   'MesProximo = Format$(Month(DataHoje), "00")
Loop

DataFinalInvertida = Format$(DataHoje, "yyyy-mm-dd")
DataHoje = Date

End Sub

Private Sub GridFatura_Click()
Dim i As Integer
Dim j As Integer
Dim resultado As String
   
   If linhaInicio <> Empty And linhaFim <> Empty And colunaFim <> Empty And colunaInicio <> Empty Then
   
      For i = 1 To linhaFim
         For j = colunaInicio To colunaFim
            resultado = resultado & GridFatura.TextMatrix(i, j)
         Next
      Next
   End If
End Sub

Private Sub GridFatura_KeyPress(KeyAscii As Integer)
   linhaInicio = GridFatura.Row
   colunaInicio = GridFatura.Col
End Sub

'Private Sub GridFatura_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   linhaInicio = GridFatura.Row
'   colunaInicio = GridFatura.Col
'End Sub

Private Sub GridFatura_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   linhaFim = GridFatura.Row
   colunaFim = GridFatura.Col
End Sub
