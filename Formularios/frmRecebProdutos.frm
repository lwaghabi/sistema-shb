VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecebProdutos 
   Caption         =   "frmRecebProdutos"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   17340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotalRecebido 
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
      Height          =   495
      Left            =   15360
      TabIndex        =   27
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9480
      TabIndex        =   23
      Top             =   6840
      Width           =   7455
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
         Height          =   615
         Left            =   4200
         TabIndex        =   25
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton txtProcessarEstoque 
         Caption         =   "Processar Estoque"
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
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nota Fiscal Recebida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   18
      Top             =   1560
      Width           =   16695
      Begin VB.CommandButton cmdIncluiRecebido 
         Caption         =   "Inclui Recebido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14760
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   9015
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
         Height          =   495
         Left            =   10440
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtValorUnit 
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
         Left            =   11520
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtValorTotalRec 
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
         Left            =   12960
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label10 
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
         Left            =   10440
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Valor Unit"
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
         Left            =   11520
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Valor Total"
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
         Left            =   12960
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid tblEquipamentos 
      Height          =   2535
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      FormatString    =   $"frmRecebProdutos.frx":0000
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
   Begin VB.TextBox txtValorTotal 
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
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtDataEntrega 
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   960
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
      Format          =   391708673
      CurrentDate     =   45145
   End
   Begin VB.TextBox txtNotaFiscal 
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
      Left            =   6000
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtFornecedor 
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
      Left            =   2880
      TabIndex        =   15
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtNumPO 
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
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "Valor Total Calculado"
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
      Left            =   12480
      TabIndex        =   28
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label16 
      Caption         =   "Recebimento de Nota Fiscal de Fornecedor"
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
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esperado"
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
      Height          =   495
      Left            =   10260
      TabIndex        =   17
      Top             =   3120
      Width           =   3210
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recebido"
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
      Height          =   495
      Left            =   13455
      TabIndex        =   16
      Top             =   3120
      Width           =   3210
   End
   Begin VB.Label Label5 
      Caption         =   "Valor Total da Nota Fiscal"
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
      TabIndex        =   14
      Top             =   690
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Data Entrega"
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
      Left            =   8160
      TabIndex        =   13
      Top             =   690
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Nota Fiscal"
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
      Left            =   6000
      TabIndex        =   12
      Top             =   690
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Fornecedor"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   690
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Número da PO"
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
      TabIndex        =   1
      Top             =   690
      Width           =   1935
   End
End
Attribute VB_Name = "frmRecebProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flagVerificacao As Boolean


Private Sub cmdIncluiRecebido_Click()
   Dim i As Integer
   Dim AcumulaRecebido As Currency
   
   tblEquipamentos.TextMatrix(tblEquipamentos.Row, 5) = txtQtd
   tblEquipamentos.TextMatrix(tblEquipamentos.Row, 6) = Format$(txtValorUnit, "##,##0.00")
   tblEquipamentos.TextMatrix(tblEquipamentos.Row, 7) = Format$(txtValorTotalRec, "##,##0.00")
   
   i = 1
   AcumulaRecebido = 0
   
   Do While i < CInt(tblEquipamentos.Rows)
   
      If tblEquipamentos.TextMatrix(i, 7) <> Empty Then
         AcumulaRecebido = Format$(AcumulaRecebido + tblEquipamentos.TextMatrix(i, 7), "##,##0.00")
      End If
      i = i + 1
   
   Loop

   txtTotalRecebido = Format$(AcumulaRecebido, "##,##0.00")
   
   txtDescricao = Empty
   txtQtd = Empty
   txtValorUnit = Empty
   txtValorTotalRec = Empty
   If CInt(tblEquipamentos.TextMatrix(tblEquipamentos.Row, 2)) < CInt(tblEquipamentos.TextMatrix(tblEquipamentos.Row, 5)) Then
         flagVerificacao = False
         tblEquipamentos.Col = 2
         tblEquipamentos.CellBackColor = vbYellow
         tblEquipamentos.Col = 5
         tblEquipamentos.CellBackColor = vbYellow
         MsgBox ("Quantidade incompatível com a PO. Nota Fiscal será recusada."), vbInformation
      
      ElseIf CInt(tblEquipamentos.TextMatrix(tblEquipamentos.Row, 2)) > CInt(tblEquipamentos.TextMatrix(tblEquipamentos.Row, 5)) Then
         tblEquipamentos.Col = 2
         tblEquipamentos.CellBackColor = vbYellow
         tblEquipamentos.Col = 5
         tblEquipamentos.CellBackColor = vbYellow
      
      Else
         
         tblEquipamentos.Col = 2
         tblEquipamentos.CellBackColor = vbWhite
         tblEquipamentos.Col = 5
         tblEquipamentos.CellBackColor = vbWhite
      End If
      
      If CInt(tblEquipamentos.TextMatrix(tblEquipamentos.Row, 3)) <> CInt(tblEquipamentos.TextMatrix(tblEquipamentos.Row, 6)) Then
         flagVerificacao = False
         tblEquipamentos.Col = 3
         tblEquipamentos.CellBackColor = vbYellow
         tblEquipamentos.Col = 6
         tblEquipamentos.CellBackColor = vbYellow
         MsgBox ("Valor incompatível com a PO. Nota Fiscal será recusada."), vbInformation
      
      Else
      
         tblEquipamentos.Col = 3
         tblEquipamentos.CellBackColor = vbWhite
         tblEquipamentos.Col = 6
         tblEquipamentos.CellBackColor = vbWhite
      
      End If
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub



Private Sub Form_Load()
   dtDataEntrega = Date
End Sub

Private Sub tblEquipamentos_Click()

   txtDescricao = tblEquipamentos.TextMatrix(tblEquipamentos.Row, 0)
   txtQtd = tblEquipamentos.TextMatrix(tblEquipamentos.Row, 5)
   txtValorUnit = tblEquipamentos.TextMatrix(tblEquipamentos.Row, 3)
   txtValorTotalRec = tblEquipamentos.TextMatrix(tblEquipamentos.Row, 7)
   
End Sub
Private Sub txtNumPO_LostFocus()
              
   flagVerificacao = True
   
   If txtNumPO = Empty Then
   
      MsgBox ("Número da PO deve ser informado"), vbInformation
      Exit Sub
   
   End If
   
   Call Rotina_AbrirBanco
   
   rs.Open "Select status from suppedidodecompra where id=('" & txtNumPO & "')", db, 3, 3
   
      If Not rs.EOF Then
         If rs!Status = 2 Then
      
            MsgBox ("PO encerrada!"), vbInformation
            FechaDB
            Exit Sub
         End If
      End If
      
   rs.Close
   
   rs.Open "SELECT *,suppedidodetalhe.status as statusItem FROM suppedidodetalhe inner join supproduto on suppedidodetalhe.grupo = supproduto.grupo and suppedidodetalhe.classe = supproduto.classe and suppedidodetalhe.codProd = supproduto.codProd WHERE id = ('" & txtNumPO & "')", db, 3, 3
   
      If rs.EOF Then
      
         MsgBox ("Não existem produtos para essa PO"), vbInformation
         FechaDB
         Exit Sub
      
      End If
      
      tblEquipamentos.Rows = 1
      rs.MoveFirst
      
      Do While Not rs.EOF
         
         If rs!statusItem = 0 Then
            tblEquipamentos.AddItem rs!nomeProd & vbTab & rs!unidade & vbTab & rs!qtdPedida - rs!qtdAtendida & vbTab & Format$(rs!valorUnitario, "##,##0.00") & vbTab & Format$(rs!ValorTotal - (rs!valorUnitario * rs!qtdAtendida), "##,##0.00")
         End If
         
         rs.MoveNext
      
      Loop
   
   rs.Close
   
   rs.Open "SELECT fornecedor FROM suppedidodecompra WHERE id = ('" & txtNumPO & "')", db, 3, 3
   
      txtFornecedor = rs!fornecedor
   
   rs.Close
   
   
   FechaDB
End Sub

Private Sub txtProcessarEstoque_Click()
   On Error GoTo Erro
   Dim i As Integer
   Call Rotina_AbrirBanco
   rs.Open "Select status from suppedidodecompra where id=('" & txtNumPO & "')", db, 3, 3
   If rs!Status < 2 Then
      db.BeginTrans
      rs.Close
      Call verificaLista
      If flagVerificacao And txtTotalRecebido = Format$(txtValorTotal, "##,##0.00") Then
         i = 1
         rs.Open "SELECT * FROM supproduto INNER JOIN suppedidodetalhe ON suppedidodetalhe.grupo = supproduto.grupo AND suppedidodetalhe.classe = supproduto.classe AND suppedidodetalhe.codProd = supproduto.codProd WHERE id = ('" & txtNumPO & "') and suppedidodetalhe.status = 0", db, 3, 3
         Do While Not rs.EOF
            
            Prod.Open "SELECT * FROM supestoque WHERE grupo = ('" & rs!Grupo & "') AND classe = ('" & rs!Classe & "') AND codProd = ('" & rs!codProd & "')", db, 3, 3
            
            pes.Open "SELECT * from suppedidodetalhe where id=('" & txtNumPO & "') and grupo = ('" & rs!Grupo & "') and classe = ('" & rs!Classe & "') and codProd = ('" & rs!codProd & "')", db, 3, 3
            pes!qtdAtendida = pes!qtdAtendida + tblEquipamentos.TextMatrix(i, 5)
            If pes!qtdAtendida = pes!qtdPedida Then
               
               pes!Status = 1
            
            End If
            If pes!acordo <> "NÃO" Then
               
               neg.Open "SELECT * FROM supacordocomercialdetalhe WHERE id = ('" & pes!acordo & "') and codProd = ('" & rs!codProd & "')", db, 3, 3
               
               If Not neg.EOF Then
                  neg!QtdEntregue = neg!QtdEntregue + tblEquipamentos.TextMatrix(i, 5)
               
                  neg.Update
               End If
               
               neg.Close
            
            End If
            pes.Update
            pes.Close
            
            If Prod.EOF Then
            
            Prod.AddNew
            
            End If
            
            Prod!Grupo = rs!Grupo
            Prod!Classe = rs!Classe
            Prod!codProd = rs!codProd
            Prod!qtdEmEstoque = Prod!qtdEmEstoque + tblEquipamentos.TextMatrix(i, 5)
            Prod!dataUltimaAtualizacao = Date
            Prod.Update
            Prod.Close
            
            Call RegistraMov(rs!Grupo, rs!Classe, rs!codProd, CInt(tblEquipamentos.TextMatrix(i, 5)), "E")
            i = i + 1
            
            rs.MoveNext
         Loop
         
         rs.Close
         
         rs.Open "Select * from suppedidodecompra where id=('" & txtNumPO & "')", db, 3, 3
         Prod.Open "SELECT COUNT(status) as atendidos from suppedidodetalhe WHERE id=('" & txtNumPO & "') and status = 1", db, 3, 3
         pes.Open "SELECT COUNT(status) as atendidos from suppedidodetalhe WHERE id=('" & txtNumPO & "')"
            If Prod!atendidos = pes!atendidos Then
               rs!Status = 2
               rs.Update
            End If
         Prod.Close
         pes.Close
         
         db.Execute ("UPDATE supestoque se,(SELECT (SUM(valorUnitario*qtdAtendida)/SUM(qtdAtendida)) AS mediaPonderada,grupo,classe,codProd FROM suppedidodetalhe GROUP BY grupo,classe,codProd) md SET se.valorMedioEstoque = md.mediaPonderada WHERE se.grupo=md.grupo AND se.classe=md.classe AND se.codProd=md.codProd")
         
         MsgBox ("Processado com sucesso."), vbInformation
         
         If rs!formaDePagamento <> "Antecipado" Then
            Call gerarfinanceiro
            MsgBox ("Financeiro gerado com sucesso"), vbInformation
         End If
         'rs.Close
         db.CommitTrans
         FechaDB
      Else
      
         MsgBox ("Nota fiscal com inconsistência. Será recusada"), vbInformation
      
      End If
   
   Else
      rs.Close
      MsgBox ("Pedido já foi processado")
   
   End If
Exit Sub
Erro: MsgBox ("Erro ao processar estoque :" & Err.Description), vbInformation
db.RollbackTrans
End Sub
Public Sub verificaLista()
   Dim i As Integer
   Dim teste As Integer
   i = 1
   teste = 0
   Do While i < CInt(tblEquipamentos.Rows)
      If CInt(tblEquipamentos.TextMatrix(i, 2)) < CInt(tblEquipamentos.TextMatrix(i, 5)) Or tblEquipamentos.TextMatrix(i, 3) <> tblEquipamentos.TextMatrix(i, 6) Then
         teste = 1
      End If
      i = i + 1
   Loop
   
   If teste = 0 Then
      
      flagVerificacao = True
   
   Else
      
      flagVerificacao = False
   
   End If
   
End Sub

Public Sub gerarfinanceiro()
   Dim i As Integer
   

   
   If CInt(txtValorTotal) > 0 Then
   
      If rs.State = 1 Then
         rs.Close: Set rs = Nothing
      End If
   
      rs.Open "Select * from notafiscalentrada where chPessoa=('" & txtFornecedor & "') and chNotaFiscalEntrada=('" & txtNotaFiscal & "')", db, 3, 3
      
      If rs.EOF Then
      
         rs.AddNew
      
      End If
      
      rs!chPessoa = txtFornecedor
      rs!chNotaFiscalEntrada = txtNotaFiscal
      rs!nfeFinalidadePagto = 2
      rs!nfeDataEmissao = Date
      rs!nfedataLanc = Date
      rs!nfeValorDaNota = txtValorTotal
      rs!nfeValorFrete = 0
      rs!nfePagtoFrete = 0
      rs!nfeValorICMS = 0
      rs!nfeValorIPI = 0
      rs!nfeNF_Boleto = 3
      Prod.Open "SELECT indice from tipolancamento where chTipoDocumento = (SELECT metodoPagamento from suppedidodecompra where id=('" & txtNumPO & "'))", db, 3, 3
      If Not Prod.EOF Then
         rs!nfeTipoLancamento = Prod!indice
      End If
      Prod.Close
      rs!nfeStatus = 0
      rs.Update
      
      rs.Close
      
      i = 1
      
      Do While i < CInt(tblEquipamentos.Rows)
         
         If CInt(tblEquipamentos.TextMatrix(i, 5)) > 0 Then
         
            rs.Open "Select * from notafiscaldetprod where chPessoa=('" & txtFornecedor & "') and chNotaFiscalEntrada=('" & txtNotaFiscal & "') and chCodProduto=('" & tblEquipamentos.TextMatrix(i, 0) & "')", db, 3, 3
            
            If rs.EOF Then
            
               rs.AddNew
            
            End If
            rs!chPessoa = txtFornecedor
            rs!chNotaFiscalEntrada = txtNotaFiscal
            rs!chCodProduto = tblEquipamentos.TextMatrix(i, 0)
            'rs!chFatura = 1
            rs!nfdCentroDeCusto = "2"
            
            Prod.Open "Select GrupoCentroDeCusto,SubGrupoCentroDeCusto from supproduto where nomeProd = ('" & tblEquipamentos.TextMatrix(i, 0) & "')", db, 3, 3
               rs!nfdGrupoCentroDeCusto = Prod!GrupoCentroDeCusto
               rs!nfdSubGrupoCentroDeCusto = Prod!SubGrupoCentroDeCusto
               pes.Open "SELECT DescricaoCentroDeCusto from centrodecusto where chCentroDeCusto=2 and chGrupoCentroDeCusto=('" & Prod!GrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto= '00' ", db, 3, 3
               rs!chProdutoFabrica = pes!DescricaoCentroDeCusto
               pes.Close
            Prod.Close
            rs!nfdQtd = tblEquipamentos.TextMatrix(i, 5)
            rs!nfdPU = tblEquipamentos.TextMatrix(i, 6)
            rs!nfdValorDaCompra = tblEquipamentos.TextMatrix(i, 7)
            
            'rs!nfdValorDaParcela
            rs!nfdStatusPagto = 0
            rs.Update
            i = i + 1
            rs.Close
         End If
      Loop
      
      i = 1
      
   '   Do While i < CInt(tblFaturas.Rows)
   '
   '      rs.Open "SELECT * FROM notafiscaldesdobramento WHERE chPessoa = ('" & txtFornecedor & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & Format$(tblFaturas.TextMatrix(i, 1), "yyyy-MM-dd") & "')", db, 3, 3
   '
   '      If rs.EOF Then
   '
   '         rs.AddNew
   '
   '      End If
   '
   '      rs!chPessoa = txtFornecedor
   '      rs!chNotaFiscalEntrada = txtNotaFiscal
   '      rs!chDataVencimento = Format$(tblFaturas.TextMatrix(i, 1), "yyyy-MM-dd")
   '      rs!nfdDataVencoriginal = Format$(tblFaturas.TextMatrix(i, 1), "yyyy-MM-dd")
   '      rs!nfdFaturaNumero = tblFaturas.TextMatrix(i, 0)
   '      rs!nfdValorDaFatura = tblFaturas.TextMatrix(i, 2)
   '      rs!nfdStatus = 0
   '      rs!nfdStatusPagto = 0
   '      rs!nfdOrdemBoleto = 0
   '      rs.Update
   '
   '      rs.Close
   '      i = i + 1
   '   Loop
   End If
End Sub



Private Sub txtValorUnit_LostFocus()
   If txtQtd <> Empty And txtValorUnit <> Empty Then
      txtValorTotalRec = Format$(txtQtd * txtValorUnit, "##,##0.00")
   End If
End Sub
