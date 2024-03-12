VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAcordoDeServico 
   Caption         =   "frmAcordoDeServico"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   13980
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid tblResumoOp 
      Height          =   1335
      Left            =   600
      TabIndex        =   28
      Top             =   7800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "Descrição                                                        |Qtd|Consumido|PO"
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11880
      TabIndex        =   27
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdEncerrar 
      Caption         =   "Encerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10320
      TabIndex        =   26
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8760
      TabIndex        =   25
      Top             =   7680
      Width           =   1335
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
      Left            =   4800
      TabIndex        =   9
      Top             =   2280
      Width           =   5415
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
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   13335
      Begin VB.TextBox txtTotalAcordo 
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
         Height          =   375
         Left            =   9840
         TabIndex        =   23
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton cmdRetira 
         Caption         =   "Retirar Da Lista"
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
         Left            =   11040
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdInsere 
         Caption         =   "Inserir Na Lista"
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
         Left            =   9840
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid tblServicos 
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         FormatString    =   "Descrição                                                                      |Qtd   |P.U    |Valor Total   |Saldo"
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
      Begin VB.TextBox txtPU 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8760
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtQtd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7920
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cmbDescricao 
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
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   7815
      End
      Begin VB.Label Label11 
         Caption         =   "Total Acordo"
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
         Left            =   8040
         TabIndex        =   24
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "P.U"
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
         Left            =   8760
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Qtd."
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
         Left            =   7920
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Descrição"
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
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   7815
      End
   End
   Begin MSComCtl2.DTPicker dtDataFim 
      Height          =   495
      Left            =   10920
      TabIndex        =   6
      Top             =   1080
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
      Format          =   239927297
      CurrentDate     =   45244
   End
   Begin MSComCtl2.DTPicker dtDataInicio 
      Height          =   495
      Left            =   8640
      TabIndex        =   5
      Top             =   1080
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
      Format          =   239927297
      CurrentDate     =   45244
   End
   Begin VB.ComboBox cmbFornecedor 
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
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
   Begin VB.ComboBox cmbAcordo 
      Appearance      =   0  'Flat
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
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label8 
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
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   1800
      Width           =   5415
   End
   Begin VB.Label Label7 
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
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Data Fim"
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
      Left            =   10920
      TabIndex        =   14
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Data Início"
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
      Left            =   8640
      TabIndex        =   13
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Fornecedores"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Acordo de Serviço"
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
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Registro de Acordo de Serviço"
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
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmAcordoDeServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAcordo_LostFocus()
   Call carregaTela
End Sub

Private Sub cmbClasse_LostFocus()
   Call Rotina_AbrirBanco
   Call carregaServico
   FechaDB
End Sub

Private Sub cmbGrupo_LostFocus()
   Call Rotina_AbrirBanco
   Call carregaClasse
   FechaDB
End Sub

Private Sub cmdEncerrar_Click()
   Call encerraAcordo
End Sub

Private Sub cmdInsere_Click()
   Call adicionaItem
   Call atualizaTotal
End Sub

Private Sub cmdRetira_Click()
   Call retiraItem
   Call atualizaTotal
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
   Call salvaAcordo
   Call limpaTela
End Sub

Private Sub Form_Load()
   Call carregaAcordo
   Call carregaFornecedores
   Call carregaDatas
   Call carregaGrupo
End Sub

Public Sub carregaAcordo()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM servacordocomercial where status = 1", db, 3, 3
      
      If rs.EOF Then
      
         cmbAcordo.AddItem 1
         cmbAcordo.ListIndex = 0
      
      Else
      
      rs.MoveFirst
      
         Do While Not rs.EOF
         
            cmbAcordo.AddItem rs!Id
            rs.MoveNext
         
         Loop
         
         pes.Open "SELECT MAX(id) as id FROM servacordocomercial", db, 3, 3
         
         cmbAcordo.AddItem pes!Id + 1
         
         pes.Close
         
         cmbAcordo.ListIndex = cmbAcordo.ListCount - 1
         
      End If
      
      
      rs.Close
      
Exit Sub
Erro: MsgBox ("Erro ao carregar contratos" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaFornecedores()
   On Error GoTo Erro
   pes.Open "Select chPessoa from pessoa where pesTipoPessoa = 2", db, 3, 3
   
      If pes.EOF Then
   
         MsgBox ("Não existem fornecedores registrados")
         FechaDB
         Exit Sub
      
      End If
   
      pes.MoveFirst
   
      Do While Not pes.EOF
   
         cmbFornecedor.AddItem pes!chPessoa
         pes.MoveNext
   
      Loop
      
      pes.Close

Exit Sub
Erro: MsgBox ("Erro ao carregar fornecedores" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaDatas()
   On Error GoTo Erro
   dtDataInicio = Date
   dtDataFim = Date
Exit Sub
Erro: MsgBox ("Erro ao carregar datas" & Err.Description), vbInformation
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

Public Sub carregaServico()
   cmbDescricao.Clear
   
   On Error GoTo Erro
   
   Prod.Open "Select descricao from servservico where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') order by codServ", db, 3, 3
   
   If Prod.EOF Then
   
      MsgBox ("Não há serviços cadastrados nessa categoria"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbDescricao.AddItem Prod!Descricao
      Prod.MoveNext
   
   Loop
   
Exit Sub
Erro: MsgBox ("Erro ao carregar serviços: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub adicionaItem()
   Dim i As Integer
   i = 1
   Do While i < tblServicos.Rows
      If cmbDescricao = tblServicos.TextMatrix(i, 0) Then
         tblServicos.TextMatrix(i, 0) = cmbDescricao
         tblServicos.TextMatrix(i, 1) = txtQtd
         tblServicos.TextMatrix(i, 2) = txtPU
         tblServicos.TextMatrix(i, 3) = Format(txtQtd * txtPU, "##,##0.00")
         Exit Sub
      End If
      i = i + 1
   Loop
   tblServicos.AddItem cmbDescricao & vbTab & txtQtd & vbTab & txtPU & vbTab & Format$(txtQtd * txtPU, "##,##0.00")
End Sub

Public Sub retiraItem()
   If tblServicos.Rows > 2 Then
      tblServicos.RemoveItem (tblServicos.Row)
   Else
      tblServicos.Rows = 1
   End If
End Sub

Public Sub atualizaTotal()
   Dim i As Integer
   Dim total As Currency
   
   i = 1
   total = 0
   
   Do While i < tblServicos.Rows
      total = total + tblServicos.TextMatrix(i, 3)
      i = i + 1
   Loop
   
   txtTotalAcordo = Format(total, "##,##0.00")
End Sub

Public Sub salvaAcordo()
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   db.BeginTrans
   
   rs.Open "Select * from servacordocomercial where id = ('" & cmbAcordo & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!Id = cmbAcordo
   rs!fornecedor = cmbFornecedor
   rs!dataInicio = dtDataInicio
   rs!dataFim = dtDataFim
   rs!ValorTotal = txtTotalAcordo
   rs!Status = 1
   rs!Grupo = Format$(cmbGrupo.ListIndex + 1, "00")
   rs!Classe = Format$(cmbClasse.ListIndex + 1, "000")
   rs.Update
   
   rs.Close
   
   i = 1
   
   Do While i < tblServicos.Rows
      rs.Open "SELECT codServ from servservico where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and descricao = ('" & tblServicos.TextMatrix(i, 0) & "')", db, 3, 3
      Prod.Open "SELECT * FROM servacordocomercialdetalhe WHERE id = ('" & cmbAcordo & "') and codServ=('" & rs!codServ & "')", db, 3, 3
      If Prod.EOF Then
      
         Prod.AddNew
      
      End If
      
      Prod!Id = cmbAcordo
      Prod!codServ = rs!codServ
      Prod!qtdTotal = tblServicos.TextMatrix(i, 1)
      Prod!precoUnit = tblServicos.TextMatrix(i, 2)
      Prod!ValorTotalServico = CDec(tblServicos.TextMatrix(i, 3))
      Prod.Update
            
      Prod.Close
      rs.Close
      i = i + 1
   Loop
   MsgBox ("Salvo com Sucesso"), vbInformation
   
   
   db.CommitTrans
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao salvar acordo: " & Err.Description), vbInformation
db.RollbackTrans
End Sub

Public Sub encerraAcordo()
   Call Rotina_AbrirBanco
   On Error GoTo Erro
   db.BeginTrans
   db.Execute ("UPDATE servacordocomercial SET status = 0,dataEncerramento=('" & Format$(Date, "yyyy-MM-dd") & "') WHERE id = ('" & cmbAcordo & "')")
   MsgBox ("Acordo encerrado"), vbInformation
   db.CommitTrans
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao encerrar acordo"), vbCritical
db.RollbackTrans
FechaDB
End Sub

Public Sub carregaTela()
   On Error GoTo Erro
   
   Dim Grupo As Integer
   Dim Classe As Integer
   
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM servacordocomercial WHERE id=('" & cmbAcordo & "')", db, 3, 3
   
   If Not rs.EOF Then
   
      cmbFornecedor = rs!fornecedor
      dtDataInicio = rs!dataInicio
      dtDataFim = rs!dataFim
      txtTotalAcordo = Format$(rs!ValorTotal, "##,##0.00")
      Grupo = rs!Grupo
      Classe = rs!Classe
      cmbGrupo.ListIndex = Grupo - 1
      Call carregaClasse
      cmbClasse.ListIndex = Classe - 1
      Call carregaServico
   
   rs.Close
   
   Call carregaTabela
   
   Call carregaResumo
   
   End If
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar tela: " & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaTabela()
   On Error GoTo Erro
   
   rs.Open "SELECT * FROM servacordocomercialdetalhe inner join servacordocomercial on servacordocomercialdetalhe.id = servacordocomercial.id where servacordocomercial.id = ('" & cmbAcordo & "')", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Acordo não possui servicos cadastrados"), vbInformation
      FechaDB
      Exit Sub
      
   End If
   
   tblServicos.Rows = 1
   
   Do While Not rs.EOF
      pes.Open "Select descricao from servservico where grupo=('" & rs!Grupo & "') and classe=('" & rs!Classe & "') and codServ=('" & rs!codServ & "')", db, 3, 3
      tblServicos.AddItem pes!Descricao & vbTab & rs!qtdTotal & vbTab & Format$(rs!precoUnit, "##,##0.00") & vbTab & Format$(rs!ValorTotalServico, "##,##0.00") & vbTab & rs!qtdTotal - rs!QtdEntregue, tblServicos.Rows
      pes.Close
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   
Exit Sub
Erro: MsgBox ("Erro ao carregar tabela: " & Err.Description), vbInformation
End Sub

Public Sub limpaTela()
   cmbFornecedor = Empty
   tblServicos.Rows = 1
   tblResumoOp.Rows = 1
   cmbClasse.Clear
   cmbGrupo = Empty
   txtPU = Empty
   txtQtd = Empty
   txtTotalAcordo = Empty
   cmbDescricao.Clear
   dtDataInicio = Date
   dtDataFim = Date
End Sub

Public Sub carregaResumo()
   On Error GoTo Erro
   
   rs.Open "SELECT ss.descricao,sacd.qtdTotal,spd.quantidade,spd.quantidadeAtendida,spd.id,sp.status FROM servacordocomercialdetalhe sacd inner join servacordocomercial sac on sacd.id = sac.id inner join servpodetalhe spd on sac.grupo = spd.grupo and sac.classe=spd.classe and sacd.codServ=spd.codServ and sac.id=spd.acordo inner join servpo sp on spd.id=sp.id inner join servservico ss on ss.grupo=spd.grupo and ss.classe = spd.classe and spd.codServ=ss.codServ where sac.id = ('" & cmbAcordo & "') and spd.quantidade>0", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Acordo não possui servicos cadastrados"), vbInformation
      FechaDB
      Exit Sub
      
   End If
   
   tblResumoOp.Rows = 1
   
   Do While Not rs.EOF
      If rs!Status = 3 Then
         tblResumoOp.AddItem rs!Descricao & vbTab & rs!qtdTotal & vbTab & rs!quantidadeAtendida & vbTab & rs!Id
      Else
         tblResumoOp.AddItem rs!Descricao & vbTab & rs!qtdTotal & vbTab & rs!quantidade & vbTab & rs!Id
      End If
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   
Exit Sub
Erro: MsgBox ("Erro ao carregar resumo: " & Err.Description), vbInformation
End Sub

Private Sub tblServicos_Click()
   cmbDescricao = tblServicos.TextMatrix(tblServicos.Row, 0)
   txtQtd = tblServicos.TextMatrix(tblServicos.Row, 1)
   txtPU = tblServicos.TextMatrix(tblServicos.Row, 2)
End Sub
