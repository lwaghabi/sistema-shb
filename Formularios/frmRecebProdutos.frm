VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecebProdutos 
   Caption         =   "frmRecebProdutos"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   14505
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
      Left            =   13200
      TabIndex        =   27
      Top             =   5280
      Width           =   1215
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
      Height          =   3375
      Left            =   9240
      TabIndex        =   23
      Top             =   6120
      Width           =   3255
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
         Left            =   480
         TabIndex        =   25
         Top             =   2160
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
      Left            =   2160
      TabIndex        =   18
      Top             =   1560
      Width           =   9735
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
         Left            =   8160
         TabIndex        =   10
         Top             =   480
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
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   2775
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
         Left            =   3840
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
         Left            =   4920
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
         Left            =   6360
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
         Left            =   840
         TabIndex        =   22
         Top             =   360
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
         Left            =   3840
         TabIndex        =   21
         Top             =   360
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
         Left            =   4920
         TabIndex        =   20
         Top             =   360
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
         Left            =   6360
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid tblEquipamentos 
      Height          =   2415
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      FormatString    =   "Descrição                                                   |Unid|Qtd|Valor Unit|Valor Total|Qtd|Valor Unit|Valor Total"
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
      Format          =   113377281
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
      Caption         =   "Total Recebido"
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
      Left            =   13200
      TabIndex        =   28
      Top             =   4560
      Width           =   1215
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
      Left            =   6300
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
      Left            =   9500
      TabIndex        =   16
      Top             =   3120
      Width           =   3210
   End
   Begin VB.Label Label5 
      Caption         =   "Valor Total "
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
      Left            =   10440
      TabIndex        =   14
      Top             =   690
      Width           =   2295
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

Private Sub tblEquipamentos_Click()

   txtDescricao = tblEquipamentos.TextMatrix(tblEquipamentos.Row, 0)
   txtQtd = tblEquipamentos.TextMatrix(tblEquipamentos.Row, 5)
   txtValorUnit = tblEquipamentos.TextMatrix(tblEquipamentos.Row, 6)
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
   
      If Not rs.EOF And rs!Status = 1 Then
      
         MsgBox ("PO encerrada!"), vbInformation
         FechaDB
         Exit Sub
      
      End If
      
   rs.Close
   
   rs.Open "SELECT *,supPedidoDetalhe.status as statusItem FROM supPedidoDetalhe inner join supProduto on supPedidoDetalhe.grupo = supProduto.grupo and supPedidoDetalhe.classe = supProduto.classe and supPedidoDetalhe.codProd = supProduto.codProd WHERE id = ('" & txtNumPO & "')", db, 3, 3
   
      If rs.EOF Then
      
         MsgBox ("Não existem produtos para essa PO"), vbInformation
         FechaDB
         Exit Sub
      
      End If
      
      tblEquipamentos.Rows = 1
      rs.MoveFirst
      
      Do While Not rs.EOF
         
         If rs!statusItem = 0 Then
            tblEquipamentos.AddItem rs!nomeProd & vbTab & rs!Unidade & vbTab & rs!qtdPedida - rs!qtdAtendida & vbTab & Format$(rs!valorUnitario, "##,##0.00") & vbTab & Format$(rs!ValorTotal - (rs!valorUnitario * rs!qtdAtendida), "##,##0.00")
         End If
         
         rs.MoveNext
      
      Loop
   
   rs.Close
   
   rs.Open "SELECT fornecedor FROM suppedidodecompra WHERE id = ('" & txtNumPO & "')", db, 3, 3
   
      txtFornecedor = rs!Fornecedor
   
   rs.Close
   
   
   FechaDB
End Sub

Private Sub txtProcessarEstoque_Click()
   Dim i As Integer
   Call Rotina_AbrirBanco
   rs.Open "Select status from supPedidoDeCompra where id=('" & txtNumPO & "')", db, 3, 3
   If rs!Status = 0 Then
      rs.Close
      Call verificaLista
      If flagVerificacao And txtTotalRecebido = Format$(txtValorTotal, "##,##0.00") Then
         i = 1
         rs.Open "SELECT * FROM supProduto INNER JOIN suppedidodetalhe ON supPedidoDetalhe.grupo = supProduto.grupo AND supPedidoDetalhe.classe = supProduto.classe AND supPedidoDetalhe.codProd = supProduto.codProd WHERE id = ('" & txtNumPO & "') and suppedidodetalhe.status = 0", db, 3, 3
         Do While Not rs.EOF
            
            Prod.Open "SELECT * FROM supestoque WHERE grupo = ('" & rs!grupo & "') AND classe = ('" & rs!classe & "') AND codProd = ('" & rs!codProd & "')", db, 3, 3
            
            pes.Open "SELECT * from suppedidodetalhe where id=('" & txtNumPO & "') and grupo = ('" & rs!grupo & "') and classe = ('" & rs!classe & "') and codProd = ('" & rs!codProd & "')", db, 3, 3
            pes!qtdAtendida = pes!qtdAtendida + tblEquipamentos.TextMatrix(i, 5)
            If pes!qtdAtendida = pes!qtdPedida Then
               
               pes!Status = 1
            
            End If
            If pes!acordo <> "NÃO" Then
               
               neg.Open "SELECT * FROM supAcordoComercialDetalhe WHERE id = ('" & pes!acordo & "') and grupo = ('" & rs!grupo & "') and classe = ('" & rs!classe & "') and codProd = ('" & rs!codProd & "')", db, 3, 3
               
               neg!qtdEntregue = neg!qtdEntregue + tblEquipamentos.TextMatrix(i, 5)
               
               neg.Update
               neg.Close
            
            End If
            pes.Update
            pes.Close
            
            If Prod.EOF Then
            
            Prod.AddNew
            
            End If
            
            Prod!grupo = rs!grupo
            Prod!classe = rs!classe
            Prod!codProd = rs!codProd
            Prod!qtdEmEstoque = Prod!qtdEmEstoque + tblEquipamentos.TextMatrix(i, 5)
            Prod!dataUltimaAtualizacao = Date
            Prod.Update
            Prod.Close
            i = i + 1
            
            rs.MoveNext
         Loop
         
         rs.Close
         
         rs.Open "Select * from supPedidoDeCompra where id=('" & txtNumPO & "')", db, 3, 3
         Prod.Open "SELECT COUNT(status) as atendidos from supPedidoDetalhe WHERE id=('" & txtNumPO & "') and status = 1", db, 3, 3
         pes.Open "SELECT COUNT(status) as atendidos from supPedidoDetalhe WHERE id=('" & txtNumPO & "')"
            If Prod!atendidos = pes!atendidos Then
               rs!Status = 1
               rs.Update
            End If
         rs.Close
         Prod.Close
         pes.Close
         MsgBox ("Processado com sucesso."), vbInformation
         
         
            
         FechaDB
         
         Call gerarfinanceiro
         
      Else
      
         MsgBox ("Nota fiscal com inconsistência. Será recusada"), vbInformation
      
      End If
   
   Else
      rs.Close
      MsgBox ("Pedido já foi processado")
   
   End If
End Sub

Private Sub txtValorTotal_LostFocus()

   Call Rotina_AbrirBanco
   
   rs.Open "SELECT Total FROM suppedidodecompra WHERE id = ('" & txtNumPO & "')", db, 3, 3
   
      If rs.EOF Then
      
         MsgBox ("Número de PO informado não existe"), vbInformation
         FechaDB
         Exit Sub
      
      End If
   
      If txtValorTotal <> rs!total Then
      
         MsgBox ("Valor Total da nota incompatível com o valor total da PO"), vbCritical
      
      End If
   
   rs.Close
   
   FechaDB

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
   Call Rotina_AbrirBanco
   Dim i As Integer
   
   rs.Open "Select * from NotaFiscalEntrada where chPessoa=('" & txtFornecedor & "') and chNotaFiscalEntrada=('" & txtNotaFiscal & "')", db, 3, 3
   
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
   Prod.Open "SELECT indice from tipoLancamento where chTipoDocumento = (SELECT metodoPagamento from supPedidoDeCompra where id=('" & txtNumPO & "'))", db, 3, 3
   rs!nfeTipoLancamento = Prod!Indice
   Prod.Close
   rs!nfeStatus = 0
   rs.Update
   
   rs.Close
   
   i = 1
   
   Do While i < CInt(tblEquipamentos.Rows)
      
      rs.Open "Select * from notaFiscalDetProd where chPessoa=('" & txtFornecedor & "') and chNotaFiscalEntrada=('" & txtNotaFiscal & "') and chCodProduto=('" & tblEquipamentos.TextMatrix(i, 0) & "')", db, 3, 3
      
      If rs.EOF Then
      
         rs.AddNew
      
      End If
      rs!chPessoa = txtFornecedor
      rs!chNotaFiscalEntrada = txtNotaFiscal
      rs!chCodProduto = tblEquipamentos.TextMatrix(i, 0)
      'rs!chFatura = 1
      rs!nfdCentroDeCusto = "2"
      
      Prod.Open "Select GrupoCentroDeCusto,SubGrupoCentroDeCusto from supProduto where nomeProd = ('" & tblEquipamentos.TextMatrix(i, 0) & "')", db, 3, 3
         rs!nfdGrupoCentroDeCusto = Prod!GrupoCentroDeCusto
         rs!nfdSubGrupoCentroDeCusto = Prod!SubGrupoCentroDeCusto
         pes.Open "SELECT DescricaoCentroDeCusto from centroDeCusto where chCentroDeCusto=2 and chGrupoCentroDeCusto=('" & Prod!GrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto= '00' ", db, 3, 3
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
   Loop
   
   i = 1
   
'   Do While i < CInt(tblFaturas.Rows)
'
'      rs.Open "SELECT * FROM NotaFiscalDesdobramento WHERE chPessoa = ('" & txtFornecedor & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & Format$(tblFaturas.TextMatrix(i, 1), "yyyy-MM-dd") & "')", db, 3, 3
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
   FechaDB
End Sub
