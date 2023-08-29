VERSION 5.00
Begin VB.Form frmEquipamentoTipo 
   Caption         =   "frmEquipamentoTipo"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   18360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Height          =   7095
      Left            =   120
      TabIndex        =   17
      Top             =   550
      Width           =   3375
      Begin VB.ListBox lstEquipamentos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5160
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Relação de Equipamentos Cadastrados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   15
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   3600
      TabIndex        =   1
      Top             =   1080
      Width           =   14655
      Begin VB.Frame Frame6 
         Caption         =   "Aviso "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4080
         TabIndex        =   29
         Top             =   3360
         Width           =   10095
         Begin VB.TextBox txtDiasAntecedencia 
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
            Height          =   405
            Left            =   4680
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Dias"
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
            Left            =   6120
            TabIndex        =   31
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Avisar com antecedência de"
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
            Left            =   3600
            TabIndex        =   30
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.TextBox dtHoje 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12600
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comandos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   4080
         TabIndex        =   9
         Top             =   4440
         Width           =   10095
         Begin VB.CommandButton cmdNovo 
            BackColor       =   &H00FFFF00&
            Caption         =   "Novo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H0000FFFF&
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
            Height          =   855
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdExcluir 
            BackColor       =   &H000000FF&
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
            Height          =   855
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton cmdSalvar 
            BackColor       =   &H0000FF00&
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
            Height          =   855
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Manutenção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   4080
         TabIndex        =   8
         Top             =   1560
         Width           =   10575
         Begin VB.Frame Frame3 
            Caption         =   "Calibração"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   360
            TabIndex        =   32
            Top             =   360
            Width           =   3135
            Begin VB.OptionButton optSim 
               Caption         =   "Sim"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   1560
               TabIndex        =   34
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton optNao 
               Caption         =   "Não"
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
               Left            =   1560
               TabIndex        =   33
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Periodicidade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   4680
            TabIndex        =   21
            Top             =   360
            Width           =   5415
            Begin VB.TextBox txtLimiteHorasCalibracao 
               Alignment       =   2  'Center
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
               Left            =   3720
               TabIndex        =   27
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtQtdTempoCalibracao 
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
               Left            =   2760
               TabIndex        =   24
               Top             =   720
               Width           =   495
            End
            Begin VB.ComboBox cmbUnidTempoCalibracao 
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
               Left            =   1080
               TabIndex        =   23
               Text            =   "Combo1"
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label16 
               Caption         =   "Limite Hrs Uso"
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
               Left            =   3360
               TabIndex        =   26
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label Label15 
               Caption         =   "Qtd"
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
               Left            =   2760
               TabIndex        =   25
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label14 
               Caption         =   "Unid. Tempo"
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
               Left            =   1080
               TabIndex        =   22
               Top             =   360
               Width           =   1575
            End
         End
      End
      Begin VB.TextBox txtDescCompleta 
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
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   14175
      End
      Begin VB.TextBox txtDescResumida 
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
         Left            =   5280
         TabIndex        =   5
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox txtTipoEquipamento 
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
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   135
         Left            =   7320
         TabIndex        =   28
         Top             =   3600
         Width           =   15
      End
      Begin VB.Image Image2 
         Height          =   4815
         Left            =   360
         Stretch         =   -1  'True
         Top             =   1755
         Width           =   3600
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Hoje"
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
         Left            =   12600
         TabIndex        =   16
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Descrição Completa"
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
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Desc. Resumida"
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
         Left            =   5280
         TabIndex        =   4
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Equipamento"
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
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Registro e Atualização de Tipos de Equipamentos"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frmEquipamentoTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Resp As String
Dim TipoDeEquipamento As String
Dim FSys As New Scripting.FileSystemObject
Dim Foto_Pedida As String
Dim Tem_Foto As String
Dim Caminho As String
Dim ImagemEqpto As String


Private Sub cmdExcluir_Click()
Call Rotina_AbrirBanco

teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & TipoDeEquipamento & "')", db, 3, 3
If teq.EOF Then
   MsgBox ("Tipo de equipamento não cadastrado."), vbInformation
   Call FechaDB
   Exit Sub
End If

Resp = MsgBox("Exclusão solicitada. Confirma???", vbExclamation + vbYesNo)

If Resp = vbYes Then
   teq.Delete
   MsgBox ("Exclusão efetuada com sucesso."), vbInformation
End If

Call FechaDB

Call CarregaLstEquipamento

Call LimpaForm

End Sub

Private Sub cmdNovo_Click()

Call LimpaForm

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub cmdSalvar_Click()

If txtTipoEquipamento = Empty Then
   MsgBox ("Cadastramento inválido. Codigo do tipo de equipamento não informado."), vbInformation
   Call FechaDB
   Exit Sub
End If

If txtDescResumida = Empty Then
   MsgBox ("Cadastramento inválido. Descrição Resumida não informada."), vbInformation
   Call FechaDB
   Exit Sub
End If

If txtDescCompleta = Empty Then
   MsgBox ("Cadastramento inválido. Descrição completa não informada."), vbInformation
   Call FechaDB
   Exit Sub
End If

'If cmbUnidTempoPrevent = Empty Then
'   MsgBox ("Cadastramento inválido. Unidade de tempo preventiva não informado."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If
'
'If txtQtdTempoPrevent = Empty Then
'   MsgBox ("Cadastramento inválido. Quantidade de tempo preventiva não informado."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If
'
'If txtLimiteHorasPrevent = Empty Then
'   MsgBox ("Cadastramento inválido. Limite de horas preventiva não informado."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If
'
'If cmbUnidTempoPeriodica = Empty Then
'   MsgBox ("Cadastramento inválido. Unidade de tempo periódica não informado."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If
'
'If txtQtdTempoPeriodica = Empty Then
'   MsgBox ("Cadastramento inválido. Quantidade de tempo periódica não informado."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If
'
'If txtLimiteHorasPeriodica = Empty Then
'   MsgBox ("Cadastramento inválido. Limite de horas periódica não informado."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If

If txtDiasAntecedencia = Empty Then
   MsgBox ("Cadastramento inválido. Não informado o número de dias de antecedência de Aviso."), vbInformation
   Call FechaDB
   Exit Sub
End If
If optSim = True Then
   If cmbUnidTempoCalibracao = Empty Then
      MsgBox ("Cadastramento inválido. Unidade de tempo Calibração não informado."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

If optSim = True Then
   If txtQtdTempoCalibracao = Empty Then
      MsgBox ("Cadastramento inválido. Quantidade de tempo Calibração não informado."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

If optSim = True Then
   If txtLimiteHorasCalibracao = Empty Then
      MsgBox ("Cadastramento inválido. Limite de horas Calibração não informado."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

Call Rotina_AbrirBanco

teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & txtTipoEquipamento & "')", db, 3, 3
If teq.EOF Then
   teq.AddNew
End If

teq!chTipoDeEquipamento = txtTipoEquipamento
teq!teqNomeEquipamentoCurto = txtDescResumida
teq!teqNomeEquipamentoLongo = txtDescCompleta
'teq!teqUnidTempoPreventiva = cmbUnidTempoPrevent
'teq!teqQtdTempoPreventiva = txtQtdTempoPrevent
'teq!teqLimiteHorasPreventiva = txtLimiteHorasPrevent
'teq!teqUnidTempoPeriodica = cmbUnidTempoPeriodica
'teq!teqQtdTempoPeriodica = txtQtdTempoPeriodica
'teq!teqLimiteHorasPeriodica = txtLimiteHorasPeriodica

If optSim = False And optNao = True Then
   teq!teqCalibracao = 0
Else
   teq!teqCalibracao = 1
End If

teq!teqUnidTempoCalibracao = cmbUnidTempoCalibracao
teq!teqQtdTempoCalibracao = txtQtdTempoCalibracao
teq!teqLimiteHorasCalibracao = txtLimiteHorasCalibracao

teq!teqDiasAntecedencia = txtDiasAntecedencia

teq.Update

MsgBox ("Tipo de equipamento salvo com sucesso"), vbInformation

Call FechaDB

Call CarregaLstEquipamento

End Sub

Private Sub Form_Load()


'CurDir 'Mostra o Diretório do VB6
'Caminho = CurDir("C")  'Mostra o diretório C:\
'Caminho ("OneDrive") 'Mostra o diretório D:\


dtHoje = Date

'cmbUnidTempoPrevent.AddItem "Dia"
'cmbUnidTempoPrevent.AddItem "Mês"
'cmbUnidTempoPrevent.AddItem "Ano"

'cmbUnidTempoPeriodica.AddItem "Dia"
'cmbUnidTempoPeriodica.AddItem "Mês"
'cmbUnidTempoPeriodica.AddItem "Ano"

cmbUnidTempoCalibracao.AddItem "Dia"
cmbUnidTempoCalibracao.AddItem "Mês"
cmbUnidTempoCalibracao.AddItem "Ano"

optSim = False
optSim.ForeColor = &H0&
optNao = True
optNao.ForeColor = &HFF&

Call CarregaLstEquipamento

Caminho = "C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Imagens\"
'Caminho = "C:\Meus Documentos\Logo"

Extensao = ".jpg"

End Sub

Public Sub CarregaLstEquipamento()

lstEquipamentos.Clear

Call Rotina_AbrirBanco

teq.Open "Select * from EquipamentoTipo", db, 3, 3
If Not teq.EOF Then
   teq.MoveFirst
   Do While Not teq.EOF
      lstEquipamentos.AddItem teq!chTipoDeEquipamento
      teq.MoveNext
   Loop
End If

End Sub



Private Sub lstEquipamentos_Click()

TipoDeEquipamento = lstEquipamentos.List(lstEquipamentos.ListIndex)

Call CarregaEquipamentoTipo

End Sub

Private Sub optNao_Click()
optNao = True
optSim = False
End Sub

Private Sub txtTipoEquipamento_LostFocus()

If Not txtTipoEquipamento = Empty Then
   TipoDeEquipamento = txtTipoEquipamento
   Call CarregaEquipamentoTipo
End If
End Sub

Public Sub CarregaEquipamentoTipo()

Call Rotina_AbrirBanco

teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & TipoDeEquipamento & "')", db, 3, 3
If teq.EOF Then
   'MsgBox ("Tipo de equipamento não cadastrado."), vbInformation
   Call FechaDB
   Exit Sub
End If

txtTipoEquipamento = teq!chTipoDeEquipamento
txtDescResumida = teq!teqNomeEquipamentoCurto
txtDescCompleta = teq!teqNomeEquipamentoLongo
'cmbUnidTempoPrevent = teq!teqUnidTempoPreventiva
'txtQtdTempoPrevent = teq!teqQtdTempoPreventiva
'txtLimiteHorasPrevent = teq!teqLimiteHorasPreventiva
'cmbUnidTempoPeriodica = teq!teqUnidTempoPeriodica
'txtQtdTempoPeriodica = teq!teqQtdTempoPeriodica
'txtLimiteHorasPeriodica = teq!teqLimiteHorasPeriodica
If teq!teqCalibracao = 0 Then
   optNao = True
   optNao.ForeColor = &HFF&
   optSim = False
   optSim.ForeColor = &H0&
   'cmbUnidTempoCalibracao = Empty
   'txtQtdTempoCalibracao = Empty
   'txtLimiteHorasCalibracao = Empty
Else
   optNao = False
   optNao.ForeColor = &H0&
   optSim = True
   optSim.ForeColor = &HFF&
End If

cmbUnidTempoCalibracao = teq!teqUnidTempoCalibracao
txtQtdTempoCalibracao = teq!teqQtdTempoCalibracao
txtLimiteHorasCalibracao = teq!teqLimiteHorasCalibracao


txtDiasAntecedencia = teq!teqDiasAntecedencia

ImagemEqpto = teq!chTipoDeEquipamento

usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Usuario não encontrado. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

Caminho = usu!usuEnderecoOneDrive

Foto_Pedida = Caminho & "Sistema\Imagens\" & ImagemEqpto & Extensao

' If txtEnderecoFoto = Empty Then
'    Tem_Foto = "Não"
' Else
'    Tem_Foto = "Sim"
    If FSys.FileExists(Foto_Pedida) Then
       Image2.Visible = True
       Image2.Picture = LoadPicture(Foto_Pedida)
    Else
       Image2.Visible = False
       MsgBox "Imagem não Disponível"
    End If
' End If

'Foto_Pedida = Caminho & txtEnderecoFoto & Extensao
'If txtEnderecoFoto = Empty Then
'   Tem_Foto = "Não"
'Else
'   Tem_Foto = "Sim"
'   Image2.Picture = LoadPicture(Foto_Pedida)
'End If

Call FechaDB


End Sub


Public Sub LimpaForm()

txtTipoEquipamento = Empty
txtDescResumida = Empty
txtDescCompleta = Empty

txtDiasAntecedencia = Empty

txtTipoEquipamento.SetFocus

End Sub
