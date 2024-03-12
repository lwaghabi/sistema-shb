VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaPO 
   Caption         =   "frmConsultaPO"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20175
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   20175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
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
      Left            =   17400
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSair 
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
      Height          =   975
      Left            =   11520
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
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
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
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
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid tblRegistros 
      Height          =   5775
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      FormatString    =   $"frmConsultaPO.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
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
      Left            =   17400
      TabIndex        =   9
      Top             =   240
      Width           =   1695
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
      Left            =   4440
      TabIndex        =   7
      Top             =   1440
      Width           =   3975
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
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Posição Atual de Ordem de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
End
Attribute VB_Name = "frmConsultaPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbGrupo_LostFocus()
   Call Rotina_AbrirBanco
   
   On Error GoTo Erro:
   
   If cmbGrupo.ListIndex > 0 Then
      
      pes.Open "Select descricao from supgrupoclasse where grupo = ('" & Format$((cmbGrupo.ListIndex), "00") & "') and classe != 0", db, 3, 3
   
      If pes.EOF Then
   
         MsgBox ("Não existem classes para esse grupo")
         FechaDB
         Exit Sub
      
      End If
   
      pes.MoveFirst
      cmbClasse.Clear
      cmbClasse.AddItem "TODAS"
   
      Do While Not pes.EOF
   
         cmbClasse.AddItem pes!Descricao
         pes.MoveNext
   
      Loop
   
      pes.Close
      
   Else
   
      cmbClasse.Clear
   
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar classes: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub cmdFiltrar_Click()
   Call geraTabela
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
   On Error GoTo Erro:
   
   txtHoje = Date
   
   Call Rotina_AbrirBanco
   
   rs.Open "Select descricao from supgrupoclasse where classe = 0", db, 3, 3

   If rs.EOF Then

      MsgBox ("Não existem grupo registrados")
      FechaDB
      Exit Sub
   
   End If

   rs.MoveFirst
   cmbGrupo.AddItem "TODOS"
   Do While Not rs.EOF

      cmbGrupo.AddItem rs!Descricao
      rs.MoveNext

   Loop
   rs.Close
   FechaDB
   
   Call geraTabela
   
Exit Sub
Erro: MsgBox ("Erro ao carregar tela: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub geraTabela()
   Dim sql As String
   
   On Error GoTo Erro
   
   tblRegistros.Rows = 1
   
   Call Rotina_AbrirBanco
   
   sql = "SELECT sp.nomeProd,sp.grupo,sp.classe,spdc.fornecedor,spdc.id,spdc.dataPedido, "
   sql = sql & "spdc.dataPrevistaDeEntrega,spd.qtdPedida,(spd.qtdPedida-spd.qtdAtendida) qtdEmAberto, "
   sql = sql & "se.qtdEmEstoque,(se.qtdEmEstoque + spd.qtdPedida) novoEstoque "
   sql = sql & "FROM suppedidodecompra spdc "
   sql = sql & "INNER JOIN suppedidodetalhe spd ON spdc.id=spd.id "
   sql = sql & "INNER JOIN supestoque se ON se.grupo = spd.grupo AND se.classe = spd.classe AND se.codProd = spd.codProd "
   sql = sql & "INNER JOIN supproduto sp ON sp.grupo = spd.grupo AND sp.classe = spd.classe AND sp.codProd = spd.codProd "
   sql = sql & "WHERE spdc.status = 1"
   
   sql = geraQuery(sql)
   
   'MsgBox (sql), vbInformation
   
   rs.Open sql, db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Nenhum registro foi encontrado!"), vbInformation
      FechaDB
      Exit Sub
   End If
   
   rs.MoveFirst
   
   Do While Not rs.EOF
   
      tblRegistros.AddItem rs!nomeProd & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!fornecedor & vbTab & rs!Id & vbTab & rs!DataPedido & vbTab & rs!dataPrevistaDeEntrega & vbTab & rs!qtdPedida & vbTab & rs!qtdEmAberto & vbTab & rs!qtdEmEstoque & vbTab & rs!novoEstoque
      rs.MoveNext
      
   Loop
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao gerar tabela: " & Err.Description), vbInformation
FechaDB
End Sub

Public Function geraQuery(sql As String) As String
   Dim query As String
   Dim cond1 As String
   Dim cond2 As String
   
   query = sql
   
   If cmbGrupo.ListIndex > 0 Then
      cond1 = Format(cmbGrupo.ListIndex, "00")
   Else
      cond1 = ""
   End If
   
   If cmbGrupo.ListIndex > 0 And cmbClasse.ListIndex > 0 Then
      cond2 = Format(cmbClasse.ListIndex, "000")
   Else
      cond2 = ""
   End If
   
   If cond1 <> "" Then
      query = query & " AND " & cond1
   End If
   
   If cond2 <> "" Then
      query = query & " AND " & cond2
   End If
   
   geraQuery = query
   
End Function
