VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCentroDeCusto 
   Caption         =   "frmCentroDeCusto"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   17175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   6495
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   16935
      Begin VB.Frame Frame4 
         Caption         =   "PARA: Grupo e Subgrupo Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   8400
         TabIndex        =   17
         Top             =   3000
         Width           =   7335
         Begin VB.ComboBox cmbSubGrupoDest 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   4
            Top             =   1560
            Width           =   6975
         End
         Begin VB.ComboBox cmbGrupoDest 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   6975
         End
         Begin VB.Label Label4 
            Caption         =   "SubGrupo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "DE: Grupo e SubGrupo Selecionados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8400
         TabIndex        =   11
         Top             =   1560
         Width           =   7335
         Begin VB.TextBox txtSubGrupoSel 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   20
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtDescSelec 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2520
            TabIndex        =   16
            Top             =   720
            Width           =   4695
         End
         Begin VB.TextBox txtGrupoSelec 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Descrição"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   15
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label txtSubGrupoSelec 
            Caption         =   "SubGrupo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdProd 
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         FormatString    =   "Ítem no SubGrupo                                                  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sub Grupo Centro de Custo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8400
         TabIndex        =   10
         Top             =   360
         Width           =   7215
         Begin VB.ComboBox cmbSubGrupo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   1
            Top             =   480
            Width           =   6855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Grupo de Centro de Custo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7935
         Begin VB.ComboBox cmbGrupo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   7455
         End
      End
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H00FFFF80&
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5520
         Width           =   1695
      End
   End
   Begin VB.Label lblAtualizaçãoDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atualização de Centros de Custos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   7140
   End
End
Attribute VB_Name = "frmCentroDeCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Limite As Integer
Dim IndAux As Integer
Dim IndLinha As Integer
Dim ChaveGrupo As String
Dim ChaveSubGrupo As String
Dim ChaveProduto As String
Dim DescricaoDest As String
Dim Verifica As String
Dim Encontrei As Integer
Dim ContaFornecedor As Integer
Dim ContaDespesa As Integer
Dim TemPessoa As Integer

Private Sub cmbGrupo_LostFocus()
Verifica = Mid$(cmbGrupo, 1, 2)
ChaveGrupo = Verifica

Call Rotina_AbrirBanco

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ChaveGrupo & "') and chSubGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("ERRO: Carga de subgrupo em Cad de centro de custo."), vbCritical
   Call FechaDB
   Exit Sub
End If

'txtCodGrupoCentroDeCusto = grdGrupo.TextMatrix(IndLinha, 0)
'txtDescGrupoCentroDeCusto = grdGrupo.TextMatrix(IndLinha, 1)

cmbSubGrupo.Clear

Do While Not Prod.EOF
   cmbSubGrupo.AddItem Prod!chSubGrupoCentroDeCusto & " - " & Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop

End Sub


Private Sub cmbGrupoDest_LostFocus()

Verifica = Mid$(cmbGrupoDest, 1, 2)
ChaveGrupo = Verifica

Call Rotina_AbrirBanco

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ChaveGrupo & "') and chSubGrupoCentroDeCusto > ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("ERRO: Carga de subgrupo em Cad de centro de custo."), vbCritical
   Call FechaDB
   Exit Sub
End If

'txtCodGrupoCentroDeCusto = grdGrupo.TextMatrix(IndLinha, 0)
'txtDescGrupoCentroDeCusto = grdGrupo.TextMatrix(IndLinha, 1)

cmbSubGrupoDest.Clear

Do While Not Prod.EOF
   cmbSubGrupoDest.AddItem Prod!chSubGrupoCentroDeCusto & " - " & Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop

End Sub

Private Sub cmbSubGrupo_LostFocus()

Verifica = Mid$(cmbSubGrupo, 1, 2)
ChaveSubGrupo = Verifica



Verifica = Mid$(cmbGrupo, 1, 2)
ChaveGrupo = Verifica

Call RotinaEncheGrid

End Sub

'Private Sub cmdExcluir_Click()
'
'If txtDescGrupoCentroDeCusto = Empty Then
'   MsgBox ("Solicitação inválida. Centro de custo não informado"), vbCritical
'   Call FechaDB
'   Exit Sub
'End If
'
'Call Rotina_AbrirBanco
'
'ProdTerc.Open "Select * from produtoterceiros where chTipoProduto = ('" & txtDescGrupoCentroDeCusto & "')", db, 3, 3
'If ProdTerc.EOF Then
'   MsgBox ("Centro de custo inexistente"), vbCritical
'   Call FechaDB
'   Exit Sub
'End If
'
'ProdTerc.Delete
'
'Call FechaDB
'
'Call CargaGridCentroDeCusto
'
'txtDescGrupoCentroDeCusto = Empty
'
'End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
Call Rotina_AbrirBanco

If cmbGrupoDest = Empty Then
   MsgBox ("Código do Grupo de Custo para salvar não foi informado"), vbInformation
   Call FechaDB
   Exit Sub
End If
'
If cmbSubGrupoDest = Empty Then
   MsgBox ("Descrição do Subgrupo de Custo destino não foi informado"), vbInformation
   Call FechaDB
   Exit Sub
End If

Verifica = Mid$(cmbGrupoDest, 1, 2)
ChaveGrupo = Verifica

Verifica = Mid$(cmbSubGrupoDest, 1, 2)
ChaveSubGrupo = Verifica

ChaveProduto = grdProd.TextMatrix(IndLinha, 0)

If ccc.State = 1 Then
   ccc.Close: Set ccc = Nothing
End If

ccc.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ChaveGrupo & "') and chSubGrupoCentroDeCusto = ('" & ChaveSubGrupo & "')", db, 3, 3
If ccc.EOF Then
   MsgBox ("Centro de custo destino não existe."), vbCritical
   Call FechaDB
   Exit Sub
Else
   DescricaoDest = ccc!DescricaoCentroDeCusto
End If

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Prod.Open "Select * from produtofornecedor where chTipoProduto = ('" & ChaveProduto & "')", db, 3, 3
If Prod.EOF Then
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   Prod.Open "Select * from produtofornecedor where chProdutoFabrica= ('" & ChaveProduto & " ')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("ERRO: Centro de custo de origem não existe. Comunicar ao analista responsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If
End If

Prod.MoveFirst

Do While Not Prod.EOF
   Prod!pinGrupoCentroDeCusto = ChaveGrupo
   Prod!pinSubGrupoCentroDeCusto = ChaveSubGrupo
   Prod!pinDescricaoCentroDeCusto = DescricaoDest
   Prod.Update
   Prod.MoveNext
Loop

nfd.Open "Select * from notafiscaldetprod where chCodProduto = ('" & ChaveProduto & "')", db, 3, 3
If nfd.EOF Then
   MsgBox ("ERRO: Não encontrado o detalhe da noota fiscal em alteração de centro de custo"), vbInformation
   Call FechaDB
   Exit Sub
End If

nfd.MoveFirst

Do While Not nfd.EOF
   nfd!nfdGrupoCentroDeCusto = ChaveGrupo
   nfd!nfdSubGrupoCentroDeCusto = ChaveSubGrupo
   nfd.Update
   nfd.MoveNext
Loop

Call FechaDB

MsgBox ("Centro de custo atualizado com sucesso"), vbInformation

txtGrupoSelec = Empty
txtSubGrupoSel = Empty
txtDescSelec = Empty
cmbGrupoDest = Empty
cmbSubGrupoDest = Empty
ChaveGrupo = cmbGrupo
ChaveSubGrupo = cmbSubGrupo

Call RotinaEncheGrid

End Sub

Private Sub Form_Load()

'CargaGridCentroDeCusto

Call Rotina_AbrirBanco

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto > ('" & "00" & "') and chSubGrupoCentroDeCusto = ('" & "00" & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("ERRO: Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

Prod.MoveFirst

cmbGrupo.Clear

Do While Not Prod.EOF
   cmbGrupo.AddItem Prod!chGrupoCentroDeCusto & " - " & Prod!DescricaoCentroDeCusto
   cmbGrupoDest.AddItem Prod!chGrupoCentroDeCusto & " - " & Prod!DescricaoCentroDeCusto
   Prod.MoveNext
Loop
   
Call FechaDB

End Sub

Private Sub txtDescGrupoCentroDeCusto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub grdProd_Click()

Call Rotina_AbrirBanco

IndLinha = grdProd.Row

If grdProd.TextMatrix(IndLinha, 0) = Empty Then
   MsgBox ("Clicar somente em linha com conteúdo."), vbInformation
   Call FechaDB
   Exit Sub
End If

Verifica = Mid$(cmbGrupo, 1, 2)
ChaveGrupo = Verifica

Verifica = Mid$(cmbSubGrupo, 1, 2)
ChaveSubGrupo = Verifica

ChaveProduto = grdProd.TextMatrix(IndLinha, 0)

Prod.Open "Select * from centrodecusto where chCentroDeCusto = ('" & "2" & "') and chGrupoCentroDeCusto = ('" & ChaveGrupo & "') and chSubGrupoCentroDeCusto = ('" & ChaveSubGrupo & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("ERRO: Carga de subgrupo em gridcusto de centro de custo."), vbCritical
   Call FechaDB
   Exit Sub
End If

txtGrupoSelec = Prod!chGrupoCentroDeCusto
txtSubGrupoSel = Prod!chSubGrupoCentroDeCusto
txtDescSelec = Prod!DescricaoCentroDeCusto

'MsgBox ("Produto - ") & ChaveProduto

End Sub

Public Sub RotinaEncheGrid()

grdProd.Rows = 1

Call Rotina_AbrirBanco

Verifica = Mid$(cmbGrupo, 1, 2)
ChaveGrupo = Verifica

Verifica = Mid$(cmbSubGrupo, 1, 2)
ChaveSubGrupo = Verifica


ContaFornecedor = 0
ContaDespesa = 0

Prod.Open "Select * from produtoentrada where pinCentroDeCusto = ('" & "2" & "') and pinGrupoCentroDeCusto = ('" & ChaveGrupo & "') and pinSubGrupoCentroDeCusto = ('" & ChaveSubGrupo & "')", db, 3, 3
If Not Prod.EOF Then

   Prod.MoveFirst
   
   grdProd.Rows = 1
   
   IndLinha = 1
   
   ContaFornecedor = 0
   
   Do While Not Prod.EOF
      Encontrei = 0
      If IndLinha > 1 Then
         Limite = IndLinha
         IndAux = 1
         Do While (IndAux < Limite)
            If Prod!chTipoProduto = grdProd.TextMatrix(IndAux, 0) Then
               Encontrei = 1
            End If
         IndAux = IndAux + 1
         Loop
      End If
      
      If Encontrei = 0 Then
         grdProd.Rows = IndLinha + 1
         grdProd.TextMatrix(IndLinha, 0) = Prod!chTipoProduto
         IndLinha = IndLinha + 1
         ContaFornecedor = ContaFornecedor + 1
      End If
      
      Prod.MoveNext
   Loop
End If

' Inicio da Rotina em ProdFornec

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

ContaDespesa = 0

Prod.Open "Select * from produtofornecedor where pinCentroDeCusto = ('" & "2" & "') and pinGrupoCentroDeCusto = ('" & ChaveGrupo & "') and pinSubGrupoCentroDeCusto = ('" & ChaveSubGrupo & "')", db, 3, 3
If Not Prod.EOF Then

   Prod.MoveFirst
   
   If ContaFornecedor = 0 Then
      grdProd.Rows = 2
      IndLinha = 1
      IndAux = 1
   Else
      grdProd.Rows = IndLinha
      Limite = IndLinha
      IndAux = 1
   End If
   'IndLinha = 1
   
   Do While Not Prod.EOF
      Encontrei = 0
      If IndLinha > 1 Then
         Limite = IndLinha
         IndAux = 1
         Do While (IndAux < Limite)
            If (ChaveGrupo = "03" And ChaveSubGrupo = "01") Or (ChaveGrupo = "03" And ChaveSubGrupo = "02") Or (ChaveGrupo = "03" And ChaveSubGrupo = "03") Then
               If Prod!chTipoProduto = grdProd.TextMatrix(IndAux, 0) Then
                  Encontrei = 1
               End If
            Else
               If Prod!chProdutoFabrica = grdProd.TextMatrix(IndAux, 0) Then
                  Encontrei = 1
                  IndAux = Limite
               End If
         End If
         
         IndAux = IndAux + 1
         Loop
      End If
      
      If Encontrei = 0 Then
         grdProd.Rows = IndLinha + 1
         If (ChaveGrupo = "03" And ChaveSubGrupo = "01") Or (ChaveGrupo = "03" And ChaveSubGrupo = "02") Or (ChaveGrupo = "03" And ChaveSubGrupo = "03") Then
            grdProd.TextMatrix(IndLinha, 0) = Prod!chTipoProduto
         Else
            grdProd.TextMatrix(IndLinha, 0) = Prod!chProdutoFabrica
         End If
         IndLinha = IndLinha + 1
         ContaDespesa = ContaDespesa + 1
      End If
      
      Prod.MoveNext
   Loop
End If

grdProd.Col = 0
grdProd.ColSel = 0
grdProd.Row = 0
grdProd.RowSel = 0
grdProd.Sort = 7

'MsgBox ("Total fornecedor - ") & ContaFornecedor
'MsgBox ("Total Despesa    - ") & ContaDespesa

End Sub
