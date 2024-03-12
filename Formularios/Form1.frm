VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEstoqueMovimentMat 
   Caption         =   "frmEstoqueMovimentMat"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   615
      Left            =   11280
      TabIndex        =   8
      Top             =   1680
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
      Height          =   735
      Left            =   11280
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid tblRegistros 
      Height          =   7335
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   20415
      _ExtentX        =   36010
      _ExtentY        =   12938
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      FormatString    =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11055
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
         Left            =   4920
         TabIndex        =   4
         Top             =   720
         Width           =   5655
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
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label2 
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
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label txtValorTotalEmEstoque 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13750
      TabIndex        =   10
      Top             =   2650
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Valor Total em Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13750
      TabIndex        =   9
      Top             =   2100
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Posição Atualizada de Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmEstoqueMovimentMat"
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
   Screen.MousePointer = vbHourglass
   Dim sql As String
      
   On Error GoTo Erro:
   
   If cmbGrupo.ListIndex > 0 And cmbClasse.ListIndex = -1 Then
      MsgBox ("Selecione uma classe para continuar"), vbInformation
      Exit Sub
   End If
   
   sql = "SELECT * FROM (SELECT sp.nomeProd,sp.grupo,sp.classe,sp.codProd,se.qtdEmEstoque,DATEDIFF(CURDATE(),se.dataUltimaRequisicao) AS diasUltimaRequisicao, "
   sql = sql & "MAX(se.dataUltimaRequisicao) AS dataUltimaRequisicao,srd.quantidadeAtendida,se.valorMedioEstoque,(se.qtdEmEstoque*se.valorMedioEstoque) AS valorTotalEstoque, "
   sql = sql & "spdc.fornecedor,MAX(spdc.dataPedido) AS dataPedido,spd.valorUnitario,TIMESTAMPDIFF(MONTH,CURDATE(),se.dataUltimaRequisicao) AS smeo FROM supproduto sp "
   sql = sql & "LEFT JOIN supestoque se ON  sp.grupo=se.grupo AND sp.classe=se.classe AND sp.codProd=se.codProd "
   sql = sql & "LEFT JOIN suprequisicaodetalhe srd ON srd.grupo=sp.grupo AND srd.classe=sp.classe AND srd.codProd=sp.codProd AND se.dataUltimaRequisicao = srd.dataProcessamento "
   sql = sql & "LEFT JOIN suppedidodetalhe spd ON spd.grupo=sp.grupo AND spd.classe=sp.classe AND spd.codProd=sp.codProd "
   sql = sql & "LEFT JOIN suppedidodecompra spdc ON spdc.id=spd.id "
   
   If cmbGrupo.ListIndex > 0 And cmbClasse.ListIndex = 0 Then
      sql = sql & "WHERE sp.grupo=('" & Format(cmbGrupo.ListIndex, "00") & "') "
   ElseIf cmbGrupo.ListIndex > 0 And cmbClasse.ListIndex > 0 Then
      sql = sql & "WHERE sp.grupo=('" & Format(cmbGrupo.ListIndex, "00") & "') and sp.classe =('" & Format(cmbClasse.ListIndex, "000") & "') "
   End If
   
   sql = sql & "GROUP BY sp.nomeProd,spdc.fornecedor,spdc.dataPedido "
   sql = sql & "ORDER BY sp.grupo,sp.classe,sp.nomeProd) AS tabela GROUP BY tabela.nomeProd"
   
   Call geraTabela(sql)

   Screen.MousePointer = vbDefault
Exit Sub
Erro: MsgBox ("Erro ao filtrar requisições, erro na linha: " & Err.Source & ", descrição do erro: " & Err.Description), vbInformation
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   MsgBox (calculaGiroEstoque("03", "001", "00005", CDate("2023-11-01"), CDate("2024-01-31")))
End Sub

Private Sub Form_Load()
   Dim Grupo As String
   Dim Classe As String
   Dim sql As String
   Dim smeo As String
   Dim ValorTotal As Currency
   Dim dataUltimaRequisicao As String
   
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
   
sql = "SELECT * FROM (SELECT sp.nomeProd,sp.grupo,sp.classe,sp.codProd,se.qtdEmEstoque,DATEDIFF(CURDATE(),se.dataUltimaRequisicao) AS diasUltimaRequisicao, "
sql = sql & "MAX(se.dataUltimaRequisicao) AS dataUltimaRequisicao,srd.quantidadeAtendida,se.valorMedioEstoque,(se.qtdEmEstoque*se.valorMedioEstoque) AS valorTotalEstoque, "
sql = sql & "spdc.fornecedor,spdc.dataPedido,spd.valorUnitario,TIMESTAMPDIFF(MONTH,CURDATE(),se.dataUltimaRequisicao) AS smeo FROM supproduto sp "
sql = sql & "LEFT JOIN supestoque se ON  sp.grupo=se.grupo AND sp.classe=se.classe AND sp.codProd=se.codProd "
sql = sql & "LEFT JOIN suprequisicaodetalhe srd ON srd.grupo=sp.grupo AND srd.classe=sp.classe AND srd.codProd=sp.codProd AND se.dataUltimaRequisicao = srd.dataProcessamento "
sql = sql & "LEFT JOIN suppedidodetalhe spd ON spd.grupo=sp.grupo AND spd.classe=sp.classe AND spd.codProd=sp.codProd "
sql = sql & "LEFT JOIN suppedidodecompra spdc ON spdc.id=spd.id "
sql = sql & "GROUP BY sp.nomeProd,spdc.fornecedor,spdc.dataPedido "
sql = sql & "ORDER BY sp.grupo,sp.classe,sp.nomeProd,spdc.dataPedido DESC) as tabela GROUP BY tabela.nomeProd"
   
   rs.Open sql, db, 3, 3
   If Not rs.EOF Then
   
   rs.MoveFirst
   tblRegistros.Rows = 1
   ValorTotal = 0
      
   Do While Not rs.EOF
      Prod.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & rs!Grupo & "') and classe = '000'", db, 3, 3
      Grupo = Prod!Descricao
      Prod.Close
      Prod.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & rs!Grupo & "') and classe = ('" & rs!Classe & "')", db, 3, 3
      Classe = Prod!Descricao
      Prod.Close
      
      If rs!smeo > 3 And rs!smeo <= 6 Then
         smeo = "SM"
      ElseIf rs!smeo > 6 Then
         smeo = "EO"
      Else
         smeo = ""
      End If
      
      If Not IsNull(rs!valorTotalEstoque) Then
         ValorTotal = ValorTotal + rs!valorTotalEstoque
      End If
      
      
      tblRegistros.AddItem rs!nomeProd & vbTab & Grupo & vbTab & Classe & vbTab & rs!qtdEmEstoque & vbTab & rs!diasUltimaRequisicao & vbTab & rs!dataUltimaRequisicao & vbTab & rs!quantidadeAtendida & vbTab & Format(rs!valorMedioEstoque, "##,##0.00") & vbTab & Format(rs!valorTotalEstoque, "##,##0.00") & vbTab & rs!fornecedor & vbTab & rs!DataPedido & vbTab & Format(rs!valorUnitario, "##,##0.00") & vbTab & Format(calculaGiroEstoque(rs!Grupo, rs!Classe, rs!codProd, "2023-11-01", Date), "##,####0.0000") & vbTab & smeo
      rs.MoveNext
   
   Loop
   
   End If
   
   rs.Close
   
   txtValorTotalEmEstoque = Format(ValorTotal, "##,##0.00")
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao abrir consulta: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub geraTabela(query As String)
   On Error GoTo Erro
   
   Dim Grupo As String
   Dim Classe As String
   Dim smeo As String
   Dim ValorTotal As Currency
   
   Call Rotina_AbrirBanco
   
   tblRegistros.Rows = 1
   
   rs.Open query, db, 3, 3
   If Not rs.EOF Then
   
   rs.MoveFirst
   ValorTotal = 0
      
   Do While Not rs.EOF
      Prod.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & rs!Grupo & "') and classe = '000'", db, 3, 3
      Grupo = Prod!Descricao
      Prod.Close
      Prod.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & rs!Grupo & "') and classe = ('" & rs!Classe & "')", db, 3, 3
      Classe = Prod!Descricao
      Prod.Close
      
      If rs!smeo > 3 And rs!smeo <= 6 Then
         smeo = "SM"
      ElseIf rs!smeo > 6 Then
         smeo = "EO"
      Else
         smeo = ""
      End If
      
      If Not IsNull(rs!valorTotalEstoque) Then
         ValorTotal = ValorTotal + rs!valorTotalEstoque
      End If
      
      tblRegistros.AddItem rs!nomeProd & vbTab & Grupo & vbTab & Classe & vbTab & rs!qtdEmEstoque & vbTab & rs!diasUltimaRequisicao & vbTab & rs!dataUltimaRequisicao & vbTab & rs!quantidadeAtendida & vbTab & Format(rs!valorMedioEstoque, "##,##0.00") & vbTab & Format(rs!valorTotalEstoque, "##,##0.00") & vbTab & rs!fornecedor & vbTab & rs!DataPedido & vbTab & Format(rs!valorUnitario, "##,##0.00") & vbTab & Format(calculaGiroEstoque(rs!Grupo, rs!Classe, rs!codProd, "2023-11-01", Date), "##,####0.0000") & vbTab & smeo
      rs.MoveNext
   
   Loop
   
   End If
   
   rs.Close
   
   FechaDB
   
   txtValorTotalEmEstoque = Format(ValorTotal, "##,##0.00")
Exit Sub
Erro: MsgBox ("Erro ao gerar tabela: " & Err.Description), vbInformation
FechaDB
End Sub

Public Function calculaGiroEstoque(Grupo As String, Classe As String, codProd As String, data1 As Date, data2 As Date) As Double
   Dim valorEstoque As Integer
   Dim ultAtual As Date
   Dim totalEstoque As Integer
   Dim totalDias As Integer
   Dim estoqueMedio As Double
   Dim giroDeEstoque As Double
   Dim totalReq As Integer
   
   totalDias = calculaDiasMes(Month(data1), Year(data1))

   neg.Open "SELECT * FROM supmovimentacaoestoque WHERE grupo=('" & Grupo & "') and classe=('" & Classe & "') and codProd=('" & codProd & "') and dataMovimentacao > '" & Format(data1, "yyyy-MM-dd") & "' and dataMovimentacao < '" & Format(data2, "yyyy-MM-dd") & "' ORDER BY dataMovimentacao,tipoMovimentacao", db, 3, 3
   Prod.Open "SELECT * FROM supestoquehist WHERE grupo=('" & Grupo & "') and classe=('" & Classe & "') and codProd=('" & codProd & "') and seloDeTempo = '" & Format(data1, "yyyy-MM") & "-" & "01" & "'", db, 3, 3

   If neg.EOF Or Prod.EOF Then
      'MsgBox ("Não foi possível calcular o giro de estoque devido a falta de informação"), vbCritical
      neg.Close
      Prod.Close
      Exit Function
   End If
   
   neg.MoveFirst
   valorEstoque = Prod!qtdEmEstoque
   ultData = Prod!seloDeTempo

   Do While Not neg.EOF

      totalEstoque = totalEstoque + (valorEstoque * (neg!dataMovimentacao - ultData))

      If neg!tipoMovimentacao = "E" Then

         valorEstoque = valorEstoque + neg!qtdMovimentado

      Else

         valorEstoque = valorEstoque - neg!qtdMovimentado

      End If

      ultData = neg!dataMovimentacao
      
      neg.MoveNext

   Loop
   
   estoqueMedio = totalEstoque / totalDias
   
   pes.Open "SELECT SUM(qtdEntregue) as totalRequisitado FROM suprequisicaodetalhe WHERE grupo=('" & Grupo & "') and classe=('" & Classe & "') and codProd=('" & codProd & "') and dataProcessamento > '" & Format(data1, "yyyy-MM-dd") & "' and dataProcessamento < '" & Format(data2, "yyyy-MM-dd") & "' ", db, 3, 3

   If IsNull(pes!totalRequisitado) Then
      totalReq = 0
   Else
      totalReq = pes!totalRequisitado
   End If
   
   giroDeEstoque = totalReq / estoqueMedio
   
   pes.Close
   Prod.Close
   neg.Close

   calculaGiroEstoque = giroDeEstoque

Exit Function
Erro: MsgBox ("Erro ao calcular giro de estoque: " & Err.Description), vbInformation
End Function
