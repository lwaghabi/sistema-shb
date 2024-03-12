VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaRequisicoes 
   Caption         =   "frmConsultaRequisicoes"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStatus 
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
      Left            =   12600
      TabIndex        =   15
      Top             =   2040
      Width           =   2535
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
      Height          =   855
      Left            =   18480
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox cmbLocalEnvio 
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
      Left            =   10080
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
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
      Height          =   855
      Left            =   15600
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
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
      Left            =   5400
      TabIndex        =   9
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
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   4695
   End
   Begin MSComCtl2.DTPicker dtDataFim 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   240582657
      CurrentDate     =   45215
   End
   Begin MSComCtl2.DTPicker dtDataInicio 
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   390201345
      CurrentDate     =   45215
   End
   Begin MSFlexGridLib.MSFlexGrid tblRequisicoes 
      Height          =   5535
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      FormatString    =   $"frmConsultaRequisicoes.frx":0000
   End
   Begin VB.Label Label7 
      Caption         =   "Status"
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
      Left            =   12600
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Local de Envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label4 
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
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Até"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "De"
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
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Requisições de Baixa"
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "frmConsultaRequisicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFiltrar_Click()
   
   Screen.MousePointer = vbHourglass
      
   On Error GoTo Erro:
   
   Dim cond1 As String
   Dim cond2 As String
   Dim cond3 As String
   Dim cond4 As String
   Dim cond5 As String
   Dim cond6 As String
   Dim query As String
   
   query = "SELECT sr.id,sp.grupo,sp.classe,sp.nomeProd,srd.quantidadeAtendida,sr.dataReq,sr.status,sr.chPessoa,sr.unidadeOperacional FROM suprequisicao sr INNER JOIN suprequisicaodetalhe srd ON sr.id=srd.id INNER JOIN supproduto sp ON sp.grupo=srd.grupo AND sp.classe=srd.classe AND sp.codProd=srd.codProd"
   
   tblRequisicoes.Rows = 1
   
   If dtDataInicio <> Empty Then
      cond1 = "sr.dataReq > " & "'" & Format(dtDataInicio, "yyyy-MM-dd") & "'"
   Else
      cond1 = ""
   End If
   If dtDataFim <> Empty Then
      cond2 = "sr.dataReq < " & "'" & Format(dtDataFim, "yyyy-MM-dd") & "'"
   Else
      cond2 = ""
   End If
   If cmbGrupo.ListIndex > 0 Then
      cond3 = "sp.grupo =  " & "'" & Format(cmbGrupo.ListIndex, "00") & "'"
   Else
      cond3 = ""
   End If
   If cmbClasse.ListIndex > 0 And cmbGrupo.ListIndex > 0 Then
      cond4 = "sp.classe = " & "'" & Format(cmbClasse.ListIndex, "000") & "'"
   Else
      cond4 = ""
   End If
   If cmbLocalEnvio <> Empty Then
      cond5 = "sr.unidadeOperacional = " & "'" & cmbLocalEnvio & "'"
   Else
      cond5 = ""
   End If
   If cmbStatus.ListIndex > 0 Then
      cond6 = "sr.status = " & "'" & cmbStatus.ListIndex - 1 & "'"
   Else
      cond6 = ""
   End If
   
   If cond1 <> "" Or cond2 <> "" Or cond3 <> "" Or cond4 <> "" Or cond5 <> "" Or cond6 <> "" Then
      query = query & " WHERE"
   End If
   
   query = query & " " & cond1
   
   If (cond2 <> "" Or cond3 <> "" Or cond4 <> "" Or cond5 <> "" Or cond6 <> "") And cond1 <> "" Then
      query = query & " AND "
   End If
   
   query = query & " " & cond2
      
   If (cond3 <> "" Or cond4 <> "" Or cond5 <> "" Or cond6 <> "") And cond2 <> "" Then
      query = query & " AND "
   End If
   
   query = query & " " & cond3
   
   If (cond4 <> "" Or cond5 <> "" Or cond6 <> "") And cond3 <> "" Then
      query = query & " AND "
   End If
   
   query = query & " " & cond4
   
   If (cond5 <> "" Or cond6 <> "") And cond4 <> "" Then
      query = query & " AND "
   End If
   
   query = query & " " & cond5
   
   If (cond6 <> "") And cond5 <> "" Then
      query = query & " AND "
   End If
   
   query = query & " " & cond6
      
   Call Rotina_AbrirBanco
   
   
   rs.Open query, db, 3, 3
   
   If Not rs.EOF Then
   
      rs.MoveFirst
      
      Do While Not rs.EOF
         If rs!Status = 0 Then
            tblRequisicoes.AddItem rs!Id & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!nomeProd & vbTab & rs!quantidadeAtendida & vbTab & rs!dataReq & vbTab & "PENDENTE ESTOQUE" & vbTab & rs!chPessoa & vbTab & rs!unidadeOperacional
         Else
            tblRequisicoes.AddItem rs!Id & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!nomeProd & vbTab & rs!quantidadeAtendida & vbTab & rs!dataReq & vbTab & "ATENDIDA" & vbTab & rs!chPessoa & vbTab & rs!unidadeOperacional
         End If
         rs.MoveNext
      
      Loop
   
   Else
      
      MsgBox ("Não existem registros com essas especificações!"), vbInformation
      
   End If
   
   rs.Close
   
   FechaDB
   
   Screen.MousePointer = vbDefault
   
Exit Sub
Erro: MsgBox ("Erro ao filtrar requisições: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub cmbGrupo_LostFocus()
   Call Rotina_AbrirBanco
   
   On Error GoTo Erro:
   
   If cmbGrupo.ListIndex > 0 Then
         
      pes.Open "Select descricao from supgrupoclasse where grupo = ('" & Format$(cmbGrupo.ListIndex, "00") & "') and classe > 0", db, 3, 3
   
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
         
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar classes: " & Err.Description), vbInformation
FechaDB
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   dtDataInicio = Date
   dtDataFim = Date
   
   cmbStatus.AddItem "TODOS"
   cmbStatus.AddItem "PENDENTE ESTOQUE"
   cmbStatus.AddItem "ATENDIDA"
   
   On Error GoTo Erro:
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
   
   rs.Open "SELECT chUnidadeOperacional FROM unidadeoperacional", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não existem unidades operacionais cadastradas"), vbInformation
      FechaDB
      Exit Sub
      
   End If
   
   rs.MoveFirst
   cmbLocalEnvio.Clear
   cmbLocalEnvio.AddItem "BASE"
   
   Do While Not rs.EOF
   
      cmbLocalEnvio.AddItem rs!chUnidadeOperacional
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   
   rs.Open "SELECT sr.id,sp.grupo,sp.classe,sp.nomeProd,srd.quantidadeAtendida,sr.dataReq,sr.status,sr.chPessoa,sr.unidadeOperacional FROM suprequisicao sr INNER JOIN suprequisicaodetalhe srd ON sr.id=srd.id INNER JOIN supproduto sp ON sp.grupo=srd.grupo AND sp.classe=srd.classe AND sp.codProd=srd.codProd", db, 3, 3
   
   If Not rs.EOF Then
   
   rs.MoveFirst
   tblRequisicoes.Rows = 1
   
   Do While Not rs.EOF
   
      If rs!Status = 0 Then
         tblRequisicoes.AddItem rs!Id & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!nomeProd & vbTab & rs!quantidadeAtendida & vbTab & rs!dataReq & vbTab & "PENDENTE ESTOQUE" & vbTab & rs!chPessoa & vbTab & rs!unidadeOperacional
      Else
         tblRequisicoes.AddItem rs!Id & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!nomeProd & vbTab & rs!quantidadeAtendida & vbTab & rs!dataReq & vbTab & "ATENDIDA" & vbTab & rs!chPessoa & vbTab & rs!unidadeOperacional
      End If
      rs.MoveNext
   
   Loop
   
   End If
   
   rs.Close
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao abrir consulta: " & Err.Description), vbInformation
FechaDB
End Sub
