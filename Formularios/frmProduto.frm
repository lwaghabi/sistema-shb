VERSION 5.00
Begin VB.Form frmProduto 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6885
   ClientLeft      =   2355
   ClientTop       =   1215
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   13110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      Height          =   6855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12975
      Begin VB.Frame CfrmList 
         Caption         =   "Relação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   3975
         Begin VB.ListBox lstProduto 
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
            TabIndex        =   22
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Grupo Locacação"
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
         Index           =   0
         Left            =   4920
         TabIndex        =   18
         Top             =   2280
         Width           =   7455
         Begin VB.ComboBox cmbUnidadeOperacional 
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
            Left            =   4200
            TabIndex        =   3
            Top             =   600
            Width           =   3015
         End
         Begin VB.ComboBox cmbLocacao 
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
            Left            =   120
            TabIndex        =   2
            Text            =   "cmbLocacao"
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label lblLabel2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Locação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4200
            TabIndex        =   20
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label lblLabel1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente Locador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1650
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Unidade"
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
         Left            =   7080
         TabIndex        =   17
         Top             =   3480
         Width           =   2535
         Begin VB.ComboBox cmbUnidade 
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
            ItemData        =   "frmProduto.frx":0000
            Left            =   720
            List            =   "frmProduto.frx":0002
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Descrição Resumida do Produto"
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
         Left            =   7200
         TabIndex        =   15
         Top             =   1080
         Width           =   5535
         Begin VB.TextBox txtNomeProd 
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
            Left            =   120
            TabIndex        =   1
            Text            =   " "
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Código do Produto"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   1080
         Width           =   2535
         Begin VB.TextBox txtMneumonico 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
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
            Left            =   120
            TabIndex        =   0
            Text            =   " "
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Descrição Completa do Produto"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   4560
         Width           =   8295
         Begin VB.TextBox txtDescCompleta 
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
            Left            =   240
            TabIndex        =   5
            Text            =   " "
            Top             =   360
            Width           =   7815
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Operações"
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
         Left            =   6600
         TabIndex        =   12
         Top             =   5760
         Width           =   6015
         Begin VB.CommandButton cmdFechar 
            BackColor       =   &H0000FFFF&
            Caption         =   "Fechar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdExcluir 
            BackColor       =   &H000000FF&
            Caption         =   "Excluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdAlterar 
            BackColor       =   &H000080FF&
            Caption         =   "Alterar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdIncluir 
            BackColor       =   &H0000FF00&
            Caption         =   "Incluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdNovo 
            BackColor       =   &H00FFFF00&
            Caption         =   "&Novo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Registro e Atualização de Produtos"
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
         TabIndex        =   16
         Top             =   240
         Width           =   7455
      End
   End
End
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Incluir As Byte
Dim msg As String
Dim baseaberta As Byte
Dim Mneumonico As String
Dim operacao As String
Dim rotinicial As Byte
Dim NaoAchei As Byte
Dim TipoOperacao As Byte '1 = inclusão, 2 = Alteração e 3 = Exclusão
Dim fim As Byte
Dim Ind As Integer
Dim PrimeiraVez As Byte
Dim Conteudo As String



Private Sub cmbLocacao_LostFocus()

cmbUnidadeOperacional.Clear

Call Rotina_AbrirBanco

uoper.Open "Select * from UnidadeOperacional Where chPessoa = ('" & cmbLocacao & "')", db, 3, 3
If uoper.EOF Then
   MsgBox ("Cadastrar as Unidades Operacionais deste Cliente SHB."), vbCritical
   Call FechaDB
   Exit Sub
End If

Do While Not uoper.EOF
   cmbUnidadeOperacional.AddItem uoper!chUnidadeOperacional
   uoper.MoveNext
Loop

Call FechaDB

End Sub



Private Sub cmdAlterar_Click()
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select * from Produto where chProduto = ('" & txtMneumonico & "')", db, 3, 3
   If Prod.EOF Then
      Prod.AddNew
   End If
      
   db.BeginTrans
   
   Prod!prdNomeProd = txtNomeProd
   Prod!prdDescCompleta = txtDescCompleta
   'Prod!prdAtividade = cmbAtividade
   If cmbUnidadeOperacional = Empty Then
         Prod!prdUnidadeOperacional = "CONTRATO"
      Else
         Prod!prdUnidadeOperacional = cmbUnidadeOperacional
   End If
   Prod!prdunidade = cmbUnidade.ListIndex
   Prod!prdgrupo = cmbLocacao
   Prod!prdLocadora = cmbLocacao
   Prod!prdUnidadeOperacional = cmbUnidadeOperacional
   Prod!prdtipo = 2

   Prod.Update
   
  db.CommitTrans

   MsgBox ("Atualização de Produtos realizada com sucesso."), vbInformation

   
  cmdIncluir.Enabled = False
  cmdAlterar.Enabled = True
  cmdExcluir.Enabled = True
  cmdFechar.Enabled = True
   
Call FechaDB

End Sub

Private Sub cmdExcluir_Click()
Dim resp As String

resp = MsgBox("Voce esta prestes a deletar este registro. Confirma???", vbYesNo)
If resp = vbYes Then

   Call Rotina_AbrirBanco

   Prod.Open "Select * from Produto where chProduto = ('" & txtMneumonico & "')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("Produto solicitado para exclusão inexistente"), vbCritical
      Call FechaDB
      Exit Sub
   End If

   db.BeginTrans
  
   Prod.Delete
   
  db.CommitTrans
   
   MsgBox ("Registro deletado com sucesso em Produto"), vbInformation
     
   Call Limpa_Tela
   
   Call FechaDB

   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   cmdIncluir.Enabled = True
   cmdFechar.Enabled = True

   Call cargaLista

End If

End Sub

Private Sub cmdFechar_Click()
       Unload Me
End Sub

Private Sub cmdIncluir_Click()

On Error Resume Next

Call Rotina_AbrirBanco

Prod.Open "Select * from Produto where chProduto = ('" & txtMneumonico & "')", db, 3, 3
If Prod.EOF Then
   Prod.AddNew
   
   
   db.BeginTrans
  
      Prod!chProduto = txtMneumonico
      Prod!prdNomeProd = txtNomeProd
      'Prod!prdAtividade = cmbAtividade
      Prod!prdgrupo = cmbLocacao
      Prod!prdLocadora = cmbLocacao
      If cmbUnidadeOperacional = Empty Then
         Prod!prdUnidadeOperacional = "CONTRATO"
      Else
         Prod!prdUnidadeOperacional = cmbUnidadeOperacional
      End If
      Prod!prdOrdemApresentacao = 0
      Prod!prdunidade = cmbUnidade.ListIndex
      Prod!prdDescCompleta = txtDescCompleta
      Prod!prdtipo = 2


      Prod!prdAnoLancamento = Year(Date)
              
      Prod.Update
   
  db.CommitTrans
  
  MsgBox ("Inclusão de produto realizada com sucesso."), vbInformation
    
  cmdAlterar.Enabled = True
  cmdExcluir.Enabled = True
  cmdIncluir.Enabled = True
  cmdFechar.Enabled = True
End If

Call FechaDB

Call cargaLista
    
End Sub

Private Sub cmdNavega_Click(Index As Integer)


   If Prod.State = 0 Then
      Call Rotina_AbrirBanco
      Prod.Open "select * from Produto", db, 3, 3
      If Prod.EOF Then
         MsgBox ("Não há registros em Produto."), vbCritical
         Call FechaDB
         Exit Sub
      End If
   End If


Select Case Index

   Case 0
        Prod.MoveFirst
        Prod.MoveNext
   Case 1
        Prod.MoveNext
        If Prod.EOF Then
           Prod.MoveLast
           Prod.MovePrevious
        End If
        
   Case 2
        Prod.MovePrevious
        If Prod.BOF Then
           Prod.MoveFirst
           Prod.MoveNext
        End If
   Case 3
        Prod.MoveLast
        Prod.MovePrevious
End Select

   If Prod.BOF = True Then
      Prod.MoveFirst
   End If
   
   If Prod.EOF = True Then
      Prod.MoveLast
   End If
   
   Call Limpa_Tela
      
   txtMneumonico = Prod!chProduto
   txtNomeProd = Prod!prdNomeProd
   cmbUnidade.ListIndex = Prod!prdunidade
   'cmbLocacao.ListIndex = Prod!prdOrdemApresentacao
   txtDescCompleta = Prod!prdDescCompleta
   'cmbAtividade = Prod!prdAtividade
   cmbLocacao = Prod!prdLocadora
   If cmbUnidadeOperacional = Empty Then
      If uoper.State = 1 Then
         uoper.Close: Set uoper = Nothing
      End If
   
      uoper.Open "Select * from UnidadeOperacional Where chPessoa = ('" & cmbLocacao & "')", db, 3, 3
      If uoper.EOF Then
         MsgBox ("Cadastrar as Unidades Operacionais deste Cliente SHB."), vbCritical
         'Call FechaDB
         Exit Sub
      End If

      Do While Not uoper.EOF
         cmbUnidadeOperacional.AddItem uoper!chUnidadeOperacional
         uoper.MoveNext
      Loop
   End If
   
   If IsNull(Prod!prdUnidadeOperacional) Then
      cmbUnidadeOperacional = Empty
   Else
      cmbUnidadeOperacional = Prod!prdUnidadeOperacional
   End If

     
   cmdIncluir.Enabled = False
   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   cmdFechar.Enabled = True
   txtMneumonico.Enabled = False
   txtNomeProd.SetFocus
   
End Sub

Private Sub cmdNovo_Click()
   txtMneumonico.Enabled = True
   
   Call Limpa_Tela
   
   txtMneumonico.SetFocus
   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   cmdIncluir.Enabled = True
   cmdFechar.Enabled = True
   
End Sub

Private Sub Form_Load()
      cmdIncluir.Enabled = False
      cmdAlterar.Enabled = False
      cmdExcluir.Enabled = False
      cmdNovo.Enabled = True
      cmdFechar.Enabled = True
      
      cmbUnidade.AddItem "M2"
      cmbUnidade.AddItem "Un"
      cmbUnidade.AddItem "Hr"
      cmbUnidade.ListIndex = 0
      
      cmbLocacao.Clear
      'cmbAtividade.Clear
      
      cmbLocacao.AddItem "LIVRE"
      cmbLocacao.ListIndex = 0
      
      'cmbAtividade.AddItem "DOBRA"
      'cmbAtividade.AddItem "EMBARCADO"
      'cmbAtividade.AddItem "HORA EXTRA"
      'cmbAtividade.AddItem "HOTEL"
      'cmbAtividade.AddItem "QUEBRA FOLGA"
      Call cargaLista
      
      Call Rotina_AbrirBanco
      
      pes.Open "Select * from Pessoa", db, 3
            
      pes.MoveFirst
      If pes.EOF Then
         MsgBox ("Dataset Pessoa sem registro. Informar ao administrador do sistema"), vbCritical
         Exit Sub
      Else
         Do While fim = 0
            If pes!pestipopessoa = 0 Then
               cmbLocacao.AddItem pes!chPessoa
            End If
            pes.MoveNext
            If pes.EOF Then
               fim = 1
            End If
         Loop
      End If
If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
PrimeiraVez = 1
Call FechaDB
      
End Sub

Private Sub Text1_Change()

End Sub







'Private Sub txtComissao_LostFocus()
'On Error Resume Next
'
'If txtComissao = Empty Then
'   msg = "Comissão não Informada"
'   MsgBox (msg), vbInformation, , "SIC rot Erro "
'   txtComissao.SetFocus
'Else
'
'   If Incluir = 1 Then
'      Incluir = 0
'      cmdIncluir.Enabled = True
'      cmdAlterar.Enabled = False
'      cmdExcluir.Enabled = False
'
'   Else
'      cmdIncluir.Enabled = False
'      cmdExcluir.Enabled = True
'      cmdAlterar.Enabled = True
'      cmdFechar.Enabled = True
'
'   End If
'
'End If

'End Sub

Private Sub lstProduto_Click()
Call Rotina_AbrirBanco
   Prod.Open "Select * from Produto where chProduto = ('" & lstProduto & "')", db, 3, 3
   
   txtMneumonico = Prod!chProduto
   txtNomeProd = Prod!prdNomeProd
   'cmbAtividade = Prod!prdAtividade
   cmbLocacao = Prod!prdLocadora
   cmbUnidadeOperacional = Prod!prdUnidadeOperacional
   cmbUnidade.ListIndex = Prod!prdunidade
   txtDescCompleta = Prod!prdDescCompleta
   
FechaDB
End Sub
Private Sub txtDescCompleta_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub txtMneumonico_LostFocus()
On Error Resume Next

Call Rotina_AbrirBanco

Prod.Open "Select * from Produto where chProduto = ('" & txtMneumonico & "')", db, 3, 3
If Prod.EOF Then
   Incluir = 1
Else
   Incluir = 0
   
   txtMneumonico = Prod!chProduto
   txtNomeProd = Prod!prdNomeProd
   'cmbAtividade = Prod!prdAtividade
   cmbLocacao = Prod!prdLocadora
   cmbUnidadeOperacional = Prod!prdUnidadeOperacional
   cmbUnidade.ListIndex = Prod!prdunidade
   txtDescCompleta = Prod!prdDescCompleta

   fim = 0
    
    
  cmdIncluir.Enabled = False
  cmdAlterar.Enabled = True
  cmdExcluir.Enabled = True
  cmdFechar.Enabled = True
  
End If

Call FechaDB
    
End Sub

Private Sub txtNomeProd_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Public Sub Limpa_Tela()

   txtNomeProd = Empty
   txtDescCompleta = Empty
   cmbUnidade.ListIndex = 0
   txtMneumonico = Empty
   cmbLocacao = Empty
   cmbUnidadeOperacional = Empty
End Sub

Public Sub cargaLista()
   Call Rotina_AbrirBanco
   
      rs.Open "SELECT chProduto FROM Produto WHERE prdOrdemApresentacao != 1", db, 3, 3
      
         If rs.EOF Then
            MsgBox "Erro ao listar Produtos", vbCritical
            FechaDB
            Exit Sub
         End If
         
         rs.MoveFirst
         lstProduto.Clear
         
         Do While Not rs.EOF
            lstProduto.AddItem rs!chProduto
            rs.MoveNext
         Loop
      rs.Close
   
   FechaDB
End Sub
