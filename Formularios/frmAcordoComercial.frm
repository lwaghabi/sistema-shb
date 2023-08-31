VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAcordoComercial 
   Caption         =   "frmAcordoComercial"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3240
      TabIndex        =   25
      Top             =   2640
      Width           =   2295
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
      TabIndex        =   24
      Top             =   2640
      Width           =   2655
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
      Height          =   855
      Left            =   8160
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
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
      Height          =   855
      Left            =   9720
      TabIndex        =   11
      Top             =   7800
      Width           =   1215
   End
   Begin VB.ComboBox cmbIdentificador 
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
      Top             =   1560
      Width           =   2415
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
      Height          =   855
      Left            =   11280
      TabIndex        =   12
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Frame frameDetalhe 
      Caption         =   "Produtos"
      Height          =   4455
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   12255
      Begin VB.CommandButton cmdRetirarDaLista 
         Caption         =   "Retirar da Lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8760
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdIncluirNaLista 
         Caption         =   "Incluir na Lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8760
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtValotUnit 
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
         Left            =   5160
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtQtdTotalProd 
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
         Left            =   4080
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid tblProdutosAcordo 
         Height          =   2295
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   "Descrição                        | Qtd     | P. Unit| Valor Total"
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
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Total"
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
         Left            =   6120
         TabIndex        =   22
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "PU"
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
         Left            =   5400
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Qtd Total Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   3480
      End
   End
   Begin MSComCtl2.DTPicker dtDataFim 
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   1560
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
      Format          =   415367169
      CurrentDate     =   45139
   End
   Begin MSComCtl2.DTPicker dtDataInicio 
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   1560
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
      Format          =   415367169
      CurrentDate     =   45139
   End
   Begin VB.ComboBox cmbFornecedores 
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
      Left            =   3240
      TabIndex        =   2
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label13 
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
      Left            =   3240
      TabIndex        =   27
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label12 
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
      Left            =   480
      TabIndex        =   26
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "Número do Acordo"
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
      Left            =   600
      TabIndex        =   21
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label4 
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
      Left            =   10560
      TabIndex        =   16
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   8400
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   3240
      TabIndex        =   14
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Registro de Acordo Comercial"
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
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "frmAcordoComercial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Linha As Integer

Private Sub cmbClasse_LostFocus()
   Call Rotina_AbrirBanco
   
   pes.Open "Select nomeProd from supproduto where grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') and classe = ('" & Format$((cmbClasse.ListIndex + 1), "000") & "') order by codProd", db, 3, 3

   If pes.EOF Then

      MsgBox ("Não existem produtos para essa classe")
      FechaDB
      cmdSair.SetFocus
      Exit Sub
   
   End If

   pes.MoveFirst
   cmbDescricao.Clear

   Do While Not pes.EOF

      cmbDescricao.AddItem pes!nomeProd
      pes.MoveNext

   Loop

   pes.Close
   
   FechaDB
End Sub

Private Sub cmbGrupo_LostFocus()
   Call carregaClasse
End Sub

Private Sub cmbIdentificador_LostFocus()

   Call Rotina_AbrirBanco
   Dim classe As Integer
   Dim grupo As Integer
   
   rs.Open "Select * from supacordocomercial where id = ('" & cmbIdentificador & "')", db, 3, 3
   
   If rs.EOF Then
      cmbFornecedores = Empty
      dtDataInicio = Date
      dtDataFim = Date
      tblProdutosAcordo.Rows = 1
      FechaDB
      Exit Sub
   
   End If
   
   cmbFornecedores = rs!Fornecedor
   dtDataInicio = rs!dataInicio
   dtDataFim = rs!dataFim
   txtTotal = Format$(rs!ValorTotal, "##,##0.00")
   grupo = rs!grupo
   classe = rs!classe
   cmbGrupo.ListIndex = grupo - 1
   Call carregaClasse
   cmbClasse.ListIndex = classe - 1
   
   rs.Close
   
   rs.Open "SELECT * FROM supacordocomercialdetalhe inner join supacordocomercial on supacordocomercialdetalhe.id = supacordocomercial.id where supacordocomercial.id = ('" & cmbIdentificador & "')", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Acordo não possui produtos cadastrados"), vbInformation
      FechaDB
      Exit Sub
      
   End If
   
   tblProdutosAcordo.Rows = 1
   
   Do While Not rs.EOF
      pes.Open "Select nomeProd from supproduto where grupo=('" & rs!grupo & "') and classe=('" & rs!classe & "') and codProd=('" & rs!codProd & "')", db, 3, 3
      tblProdutosAcordo.AddItem pes!nomeProd & vbTab & rs!qtdTotal - rs!qtdEntregue & vbTab & Format$(rs!precoUnit, "##,##0.00") & vbTab & Format$(rs!ValorTotalProduto, "##,##0.00"), tblProdutosAcordo.Rows
      pes.Close
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   cmbClasse.SetFocus
   
   FechaDB
End Sub

Private Sub cmdEncerrar_Click()
   Call Rotina_AbrirBanco
   On Error GoTo Erro
   db.Execute ("UPDATE supAcordoComercial SET status = 0,dataEncerramento=('" & Format$(Date, "yyyy-MM-dd") & "') WHERE id = ('" & cmbIdentificador & "')")
   MsgBox ("Acordo encerrado"), vbInformation
   
   FechaDB
Exit Sub

Erro: MsgBox ("Erro ao encerrar acordo"), vbCritical

End Sub

Private Sub cmdIncluirNaLista_Click()
   Dim i As Integer
   
   i = 0
   
   Do While i < tblProdutosAcordo.Rows
      If tblProdutosAcordo.TextMatrix(i, 0) = cmbDescricao Then
         tblProdutosAcordo.TextMatrix(i, 1) = txtQtdTotalProd
         tblProdutosAcordo.TextMatrix(i, 2) = Format$(txtValotUnit, "##,##0.00")
         tblProdutosAcordo.TextMatrix(i, 3) = Format$(txtQtdTotalProd * txtValotUnit, "##,##0.00")
         Call calculaTotal
         Exit Sub
      End If
      i = i + 1
   Loop
   
   
   If cmbGrupo <> Empty And cmbClasse <> Empty And cmbDescricao <> Empty Then
      On Error GoTo Erro
      tblProdutosAcordo.AddItem cmbDescricao & vbTab & txtQtdTotalProd & vbTab & txtValotUnit & vbTab & Format$(txtQtdTotalProd * txtValotUnit, "##,##0.00"), tblProdutosAcordo.Rows
   End If
   Call calculaTotal
Exit Sub
Erro: MsgBox ("Verificar valores informados"), vbInformation

End Sub

Private Sub cmdRetirarDaLista_Click()
   If tblProdutosAcordo.Rows = 2 Then
      tblProdutosAcordo.Rows = 1
   Else
      On Error GoTo Erro:
      tblProdutosAcordo.RemoveItem (tblProdutosAcordo.Row)
   End If
   Call calculaTotal
Exit Sub
Erro: MsgBox ("Erro ao excluir da lista"), vbInformation
End Sub

Private Sub cmdSalvar_Click()
   Call Rotina_AbrirBanco
   Dim i As Integer
   
   On Error GoTo Erro
   
   rs.Open "Select * from supAcordoComercial where id = ('" & cmbIdentificador & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!id = cmbIdentificador
   rs!Fornecedor = cmbFornecedores
   rs!dataInicio = dtDataInicio
   rs!dataFim = dtDataFim
   rs!ValorTotal = txtTotal
   rs!Status = 1
   rs!grupo = Format$(cmbGrupo.ListIndex + 1, "00")
   rs!classe = Format$(cmbClasse.ListIndex + 1, "000")
   rs.Update
   
   rs.Close
   
   i = 1
   
   db.BeginTrans
   
   db.Execute ("DELETE FROM supAcordoComercialDetalhe WHERE id=('" & cmbIdentificador & "')")

   Do While i < tblProdutosAcordo.Rows
      rs.Open "SELECT codProd from supProduto where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and nomeProd = ('" & tblProdutosAcordo.TextMatrix(i, 0) & "')", db, 3, 3
      Prod.Open "SELECT * FROM supAcordoComercialDetalhe WHERE id = ('" & cmbIdentificador & "') and codProd=('" & rs!codProd & "')", db, 3, 3
      If Prod.EOF Then
      
         Prod.AddNew
      
      End If
      
      Prod!id = cmbIdentificador
      Prod!codProd = rs!codProd
      Prod!qtdTotal = tblProdutosAcordo.TextMatrix(i, 1)
      Prod!precoUnit = tblProdutosAcordo.TextMatrix(i, 2)
      Prod!ValorTotalProduto = tblProdutosAcordo.TextMatrix(i, 3)
      Prod.Update
            
      Prod.Close
      rs.Close
      i = i + 1
   Loop
   MsgBox ("Salvo com Sucesso"), vbInformation
   
   db.CommitTrans
   
FechaDB
Exit Sub

Erro: MsgBox ("Erro ao salvar"), vbCritical

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub dtDataFim_LostFocus()
   If dtDataInicio >= dtDataFim Then
      MsgBox ("Data inválida"), vbInformation
      cmdSair.SetFocus
   End If
End Sub

Private Sub Form_Load()

   dtDataFim = Date
   dtDataInicio = Date
   
   Call Rotina_AbrirBanco
   
   tblProdutosAcordo.Rows = 1
   
      rs.Open "SELECT * FROM supacordocomercial where status = 1", db, 3, 3
      
      If rs.EOF Then
      
         cmbIdentificador.AddItem 1
         cmbIdentificador.ListIndex = 0
      
      Else
      
      rs.MoveFirst
      
         Do While Not rs.EOF
         
            cmbIdentificador.AddItem rs!id
            rs.MoveNext
         
         Loop
         
         pes.Open "SELECT MAX(id) as id FROM supacordocomercial", db, 3, 3
         
         cmbIdentificador.AddItem pes!id + 1
         
         pes.Close
         
         cmbIdentificador.ListIndex = cmbIdentificador.ListCount - 1
         
      End If
      
      
      rs.Close
   
      pes.Open "Select chPessoa from Pessoa where pesTipoPessoa=2", db, 3, 3
   
      If pes.EOF Then
   
         MsgBox ("Não existem fornecedores registrados")
         FechaDB
         Exit Sub
      
      End If
   
      pes.MoveFirst
   
      Do While Not pes.EOF
   
         cmbFornecedores.AddItem pes!chPessoa
         pes.MoveNext
   
      Loop
      
      pes.Close
         
      rs.Open "Select descricao from supgrupoclasse where classe = 0", db, 3, 3
   
      If rs.EOF Then
   
         MsgBox ("Não existem grupo registrados")
         FechaDB
         Exit Sub
      
      End If
   
      rs.MoveFirst
   
      Do While Not rs.EOF
   
         cmbGrupo.AddItem rs!Descricao
         rs.MoveNext
   
      Loop
   
      rs.Close
   
   FechaDB
End Sub

Public Sub calculaTotal()
   Dim i As Integer
   Dim total As Currency
   i = 1
   total = 0
   Do While i < tblProdutosAcordo.Rows
      total = total + tblProdutosAcordo.TextMatrix(i, 3)
      i = i + 1
   Loop
   txtTotal = Format$(total, "##,##0.00")
End Sub

Private Sub tblProdutosAcordo_Click()
   cmbDescricao = tblProdutosAcordo.TextMatrix(tblProdutosAcordo.Row, 0)
   txtQtdTotalProd = tblProdutosAcordo.TextMatrix(tblProdutosAcordo.Row, 1)
   txtValotUnit = tblProdutosAcordo.TextMatrix(tblProdutosAcordo.Row, 2)
End Sub

Public Sub carregaClasse()
   
   Call Rotina_AbrirBanco
   
   
   pes.Open "Select descricao from supgrupoclasse where grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') and classe != 0", db, 3, 3

   If pes.EOF Then

      MsgBox ("Não existem classes para esse grupo")
      FechaDB
      Exit Sub
   
   End If

   pes.MoveFirst
   cmbClasse.Clear

   Do While Not pes.EOF

      cmbClasse.AddItem pes!Descricao
      pes.MoveNext

   Loop

   pes.Close
   
End Sub
