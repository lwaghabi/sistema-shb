VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaServicosPO 
   Caption         =   "frmConsultaServicosPO"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   14595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cmbClasse 
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
      Left            =   3960
      TabIndex        =   5
      Top             =   2040
      Width           =   4335
   End
   Begin VB.ComboBox cmbGrupo 
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
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid tblServicos 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Serviços                                                               |Grupo|Classe|Fornecedor|PO|Data Pedido|Qtd"
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
   Begin VB.Label Label3 
      Caption         =   "Classe"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Grupo"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Consulta Posição de Ordem De Compra de Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmConsultaServicosPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim queryBase As String

Private Sub cmbGrupo_LostFocus()
   Call carregaClasse
End Sub

Private Sub cmdFiltrar_Click()
   Call filtraRegistros
End Sub

Private Sub Form_Load()
 Call carregaGrupo
 queryBase = "SELECT ss.descricao,spd.grupo,spd.classe,sp.fornecedor,sp.id,sp.dataPedido,spd.quantidade FROM servpo sp INNER JOIN servpodetalhe spd ON sp.id=spd.id INNER JOIN servservico ss ON ss.grupo=spd.grupo AND spd.classe=ss.classe AND spd.codServ=ss.codServ WHERE sp.status=1"
 Call carregaTabela(queryBase)
End Sub

Public Sub carregaTabela(query As String)
   On Error GoTo Erro
   Call Rotina_AbrirBanco

   rs.Open query, db, 3, 3
   
   tblServicos.Rows = 1

   If Not rs.EOF Then
   
      Do While Not rs.EOF
      
         tblServicos.AddItem rs!Descricao & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!fornecedor & vbTab & rs!Id & vbTab & rs!DataPedido & vbTab & rs!quantidade
         rs.MoveNext
      
      Loop
   
   End If
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar tabela: " & Err.Description), vbInformation
End Sub


Public Sub carregaGrupo()
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select * from servgrupoclasse where classe = '000'", db, 3, 3
      If Prod.EOF Then
         MsgBox ("ERRO: Arquivo vazio."), vbCritical
         Call FechaDB
         Exit Sub
      End If
      
      Prod.MoveFirst
      
      Do While Not Prod.EOF
         cmbGrupo.AddItem Prod!Descricao
         Prod.MoveNext
      Loop
   Prod.Close
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar grupos" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaClasse()
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   neg.Open "Select * from servgrupoclasse where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe != '000' ", db, 3, 3
   If neg.EOF Then
      MsgBox "Erro: Não existem classes nesse grupo", vbCritical
      FechaDB
      Exit Sub
   End If
   neg.MoveFirst
   
   cmbClasse.Clear
   
   Do While Not neg.EOF
      cmbClasse.AddItem neg!Descricao
      neg.MoveNext
   Loop
   
   neg.Close
   
Exit Sub
Erro: MsgBox ("Erro ao carregar classes" & Err.Description), vbInformation
End Sub

Public Sub filtraRegistros()
   Dim bloco1 As String
   Dim bloco2 As String
   Dim queryFinal As String
   
   If cmbGrupo <> Empty Then
      bloco1 = " AND spd.grupo = '" & Format(cmbGrupo.ListIndex + 1, "00") & "'"
   End If
   
   If cmbClasse <> Empty Then
      bloco2 = " AND spd.classe = '" & Format(cmbClasse.ListIndex + 1, "000") & "'"
   End If
      
   queryFinal = queryBase & bloco1 & bloco2
      
   Call carregaTabela(queryFinal)
End Sub
