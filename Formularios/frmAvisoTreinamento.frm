VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAvisoTreinamento 
   Caption         =   "frmAvisoTreinamento"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Width           =   2775
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
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1815
   End
   Begin VB.OptionButton optAviso 
      Caption         =   "Não mostrar mais este AVISO"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid grdAviso 
      Height          =   4455
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Colaborador                                                                                 |Vencito. em  "
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
   Begin VB.Label Label1 
      Caption         =   "TREINAMENTO - AVISO"
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
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   8415
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
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   10935
   End
End
Attribute VB_Name = "frmAvisoTreinamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColaboradorAnterior As String
Dim DataAnterior As Date
Dim DataBase As String
Dim DataDias As Date
Dim DataInvertida As String
Dim DataHojeInvertida As String

Dim Dia As Integer
Dim Mes As Integer
Dim Ano As Integer
Dim DiaDb As Integer
Dim MesDb As Integer
Dim AnoDb As Integer

Dim Linha As Integer
Dim Passei As Integer


Private Sub cmdSair_Click()

If Not optAviso = True Then
   Unload Me
End If

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Erro no acesso a Usuario na rotina de atualização de tREINAMENTO. Mostrar aviso. Comunicar Analista responsável"), vbCritical
   End
End If

If optAviso = True Then
   usu!usuAvisoTreinamento = 0
   usu.Update
End If

Passei = 1

Unload Me

End Sub

Private Sub Form_Load()

If Passei = 1 Then
   Exit Sub
End If


txtHoje = Date
ColaboradorAnterior = Empty
DataAnterior = Empty
optAviso = False

Ano = Year(Date)
Mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where  chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Usuário inexistente. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   End
End If

If Not usu!usuAvisoTreinamento = 1 Then
   Call FechaDB
   Exit Sub
   Unload Me
End If

Call LimpaGridAviso

agcto.Open "Select * from TreinamentoAgenda where agctoStatus = ('" & 0 & "')", db, 3, 3
If agcto.EOF Then
   Call FechaDB
   Exit Sub
End If

Linha = 1
      

agcto.MoveFirst

Do While Not agcto.EOF
   If pes.State = 1 Then
      pes.Close: Set pes = Nothing
   End If
   pes.Open "Select * from Pessoa where pesRazaoSocial = ('" & agcto!chPessoa & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Registro em Agenda de Treinamento não encontrado em Pessoa."), vbInformation
      Call FechaDB
      Exit Sub
   End If
      
   If Not pes!pesStatusPessoa = 3 Then
      
      If cto.State = 1 Then
         cto.Close: Set cto = Nothing
      End If
      
      cto.Open "Select * from Treinamento where chNomeCurso = ('" & agcto!chNomeCurso & "')", db, 3, 3
      If Not cto.EOF Then
         DataDias = Date + cto!ctoAvisoEm
         Ano = Year(DataDias)
         Mes = Month(DataDias)
         Dia = Day(DataDias)
         DataBase = Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
   
         AnoDb = Year(agcto!agctoDataProxCurso)
         MesDb = Month(agcto!agctoDataProxCurso)
         DiaDb = Day(agcto!agctoDataProxCurso)
      
         DataInvertida = AnoDb & "-" & Format(MesDb, "00") & "-" & Format$(DiaDb, "00")
      
         If Not agcto!agctoStatus > 0 Then
            If Not (DataInvertida > DataBase) Then
               If Not (agcto!chPessoa = ColaboradorAnterior And DataAnterior = agcto!agctoDataProxCurso) Then
                  If pes.State = 1 Then
                     pes.Close: Set pes = Nothing
                  End If
         
                  pes.Open "Select * from Pessoa where pesRazaoSocial = ('" & agcto!chPessoa & "')", db, 3, 3
                  If Not pes.EOF Then
                     If Not pes!pesStatusPessoa = 3 Then
                        grdAviso.Rows = Linha + 1
                        If agcto!chPessoa = ColaboradorAnterior Then
                           grdAviso.TextMatrix(Linha, 0) = Empty
                        Else
                           grdAviso.TextMatrix(Linha, 0) = agcto!chPessoa
                        End If
                        grdAviso.TextMatrix(Linha, 1) = agcto!agctoDataProxCurso
                        Linha = Linha + 1
                        ColaboradorAnterior = agcto!chPessoa
                        DataAnterior = agcto!agctoDataProxCurso
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   
agcto.MoveNext

Loop
 
End Sub
Public Sub LimpaGridAviso()
grdAviso.Rows = 2
Linha = 1
    grdAviso.TextMatrix(Linha, 0) = Empty
    grdAviso.TextMatrix(Linha, 1) = Empty
End Sub


