VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAviso 
   Caption         =   "frmAviso"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAviso 
      Caption         =   "Não mostrar mais este AVISO"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   6600
      Width           =   3855
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Fechar Aviso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid grdAviso 
      Height          =   4455
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Colaborador                                                                  |Data Exame                 "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Relação de Colaboradores com data dentro do prazo de aviso"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   10935
   End
   Begin VB.Label Label1 
      Caption         =   "ASO - AVISO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColaboradorAnterior As String
Dim pessoaAnterior As String
Dim DataAnterior As Date
Dim DataBase As String
Dim DataDias As Date
Dim DataInvertida As String
Dim DataHojeInvertida As String

Dim Dia As Integer
Dim mes As Integer
Dim ano As Integer
Dim DiaDb As Integer
Dim MesDb As Integer
Dim AnoDb As Integer

Dim Linha As Integer



Private Sub cmdSair_Click()

If Not optAviso = True Then
   Unload Me
End If

Call Rotina_AbrirBanco

usu.Open "Select * from usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Erro no acesso a usuario na rotina de atualização de mostrar aviso. Comunicar Analista responsável"), vbCritical
   End
End If

If optAviso = True Then
   usu!usuMostrarAviso = 0
   usu.Update
End If

Unload Me

End Sub

Private Sub Form_Load()

txtHoje = Date
If Not ColaboradorAnterior = Empty Then
   Exit Sub
End If
ColaboradorAnterior = Empty
DataAnterior = Empty
optAviso = False

ano = Year(Date)
mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

Call Rotina_AbrirBanco

usu.Open "Select * from usuario where  chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Usuário inexistente. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   End
End If

If Not usu!usuMostrarAviso = 1 Then
   Call FechaDB
   Exit Sub
   Unload Me
End If

Call LimpaGridAviso

asoa.Open "Select * from asoagenda where asoaStatus = ('" & 0 & "') and status = ('" & 1 & "')", db, 3, 3
If asoa.EOF Then
   Call FechaDB
   Exit Sub
End If

Linha = 1

asoa.MoveFirst

Do While Not asoa.EOF
   If Not asoa!chPessoa = pessoaAnterior Then
      pessoaAnterior = asoa!chPessoa
      If pes.State = 1 Then
         pes.Close: Set pes = Nothing
      End If
      pes.Open "Select * from pessoa where pesRazaoSocial = ('" & asoa!chPessoa & "')", db, 3, 3
      If pes.EOF Then
         MsgBox ("Registro em Agenda não encontrado em pessoa."), vbInformation
         Call FechaDB
         Exit Sub
      End If
   End If
   
   If Not pes!pesStatusPessoa = 3 Then
      If asoe.State = 1 Then
         asoe.Close: Set asoe = Nothing
      End If
      
      asoe.Open "Select * from asoexame where chNomeExame = ('" & asoa!chNomeExame & "')", db, 3, 3
      If Not asoe.EOF Then
         If asoe!exmUnidTempo = 0 Then
            DataDias = Date + asoe!exmPrazoAviso
            ano = Year(DataDias)
            mes = Month(DataDias)
            Dia = Day(DataDias)
            DataBase = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
         Else
            If asoe!exmUnidTempo = 1 Then
               ano = Year(Date)
               mes = Month(Date)
               mes = mes + asoe!exmPrazoAviso
               If mes > 12 Then
                  ano = Year(Date)
                  ano = ano + 1
                  mes = mes - 12
               End If
               Dia = Day(Date)
               DataBase = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
            Else
               ano = Year(Date)
               ano = ano + asoe!exmPrazoAviso
               mes = Month(Date)
               Dia = Day(Date)
               DataBase = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
            End If
         End If
      End If
      
      AnoDb = Year(asoa!asoaDataProxExame)
      MesDb = Month(asoa!asoaDataProxExame)
      DiaDb = Day(asoa!asoaDataProxExame)
   
      DataInvertida = AnoDb & "-" & Format(MesDb, "00") & "-" & Format$(DiaDb, "00")
   
      If (DataInvertida > DataHojeInvertida) Or ((DataInvertida < DataHojeInvertida) And asoa!asoaStatus = 0) Then
         If Not (DataInvertida > DataBase) Then
            If Not (asoa!chPessoa = ColaboradorAnterior And DataAnterior = asoa!asoaDataProxExame) Then
               grdAviso.Rows = Linha + 1
               If asoa!chPessoa = ColaboradorAnterior Then
                  grdAviso.TextMatrix(Linha, 0) = Empty
               Else
                  grdAviso.TextMatrix(Linha, 0) = asoa!chPessoa
               End If
               grdAviso.TextMatrix(Linha, 1) = asoa!asoaDataProxExame
               Linha = Linha + 1
               ColaboradorAnterior = asoa!chPessoa
               DataAnterior = asoa!asoaDataProxExame
            End If
         End If
      End If
   End If
   
asoa.MoveNext

Loop
   
If grdAviso.TextMatrix(1, 0) = Empty Then
   grdAviso.Rows = Linha + 1
   grdAviso.TextMatrix(Linha, 0) = "Não há exames nos próximos 20 dias"
End If
End Sub
Public Sub LimpaGridAviso()
grdAviso.Rows = 2
Linha = 1
    grdAviso.TextMatrix(Linha, 0) = Empty
    grdAviso.TextMatrix(Linha, 1) = Empty
End Sub
