VERSION 5.00
Begin VB.Form frmEndereco 
   Caption         =   "frmEndereco"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComplemento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7920
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtNumero 
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
      Height          =   435
      Left            =   6720
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcluir 
      BackColor       =   &H008080FF&
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      BackColor       =   &H0000FF00&
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtCEP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtUF 
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
      Height          =   435
      Left            =   9120
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtCidade 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5160
      TabIndex        =   9
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox txtBairro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox txtRua 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   6015
   End
   Begin VB.TextBox txtApelido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Complemento"
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
      TabIndex        =   19
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Número"
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
      Left            =   6720
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "CEP"
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
      Left            =   3120
      TabIndex        =   17
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "UF"
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
      Left            =   9120
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Cidade"
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
      Left            =   5160
      TabIndex        =   15
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Bairro"
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
      TabIndex        =   14
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Rua"
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
      TabIndex        =   13
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Apelido"
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
      TabIndex        =   12
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Registro e Atualização de Endereços"
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
      Left            =   720
      TabIndex        =   11
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "frmEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Integer
Dim PauseTime As Integer
Dim Start As Integer
Dim Timer As Integer

Private Sub cmdExcluir_Click()
   Call Rotina_AbrirBanco
      db.Execute ("Delete from supendereco where apelido = ('" & txtApelido & "')")
      MsgBox ("Endereço excluido com sucesso!"), vbInformation
      db.Close
   FechaDB
   limpaCampos
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
   Call Rotina_AbrirBanco
   
   rs.Open "Select * from supendereco where apelido = ('" & txtApelido & "')", db, 3, 3

      If rs.EOF Then
      
         rs.AddNew
      
      End If
   
      rs!apelido = txtApelido
      rs!rua = txtRua
      rs!numero = txtNumero
      rs!complemento = txtComplemento
      rs!bairro = txtBairro
      rs!cidade = txtCidade
      rs!uf = txtUF
      rs!cep = txtCEP
      rs.Update
      MsgBox ("Salvo com sucesso!"), vbInformation
   
   rs.Close
   
   FechaDB
   
   limpaCampos
End Sub

Private Sub txtApelido_LostFocus()
   Call Rotina_AbrirBanco
   
   rs.Open "Select * from supendereco where apelido =('" & txtApelido & "')", db, 3, 3
      If rs.EOF Then
      
         flag = 0
         Exit Sub
         
      Else
         flag = 1
         txtApelido = txtApelido
         txtCEP = rs!cep
         txtRua = rs!rua
         txtNumero = rs!numero
         txtComplemento = rs!complemento
         txtBairro = rs!bairro
         txtCidade = rs!cidade
         txtUF = rs!uf
         cmdSalvar.SetFocus
      
      End If
   rs.Close
   
   FechaDB
End Sub

Private Sub txtCEP_LostFocus()
   If flag = 0 Then
    Dim http As New MSXML2.ServerXMLHTTP60
    Dim doc As DOMDocument60
    Dim valor As MSXML2.IXMLDOMElement
    Dim success As Boolean
    Dim url As String
   
    txtBairro = Empty
    txtCidade = Empty
    txtRua = Empty
    txtUF = Empty
    txtNumero = Empty
    txtComplemento = Empty
    
    url = "https://viacep.com.br/ws/" & txtCEP & "/xml/"
    
    http.Open "GET", url
    http.Send
    If http.Status <> 200 Then
       MsgBox ("Erro ao realizar consulta"), vbInformation
    Else
       Dim cep As String
       Dim rua As String
       Dim bairro As String
       Dim cidade As String
       Dim uf As String
   
       Set valor = http.responseXML.getElementsByTagName("xmlcep")(0)
       
       If valor.FirstChild.Text = "true" Then
       
          MsgBox ("CEP não encontrado"), vbInformation
       
       Else
          
          cep = valor.getElementsByTagName("cep")(0).Text
          rua = valor.getElementsByTagName("logradouro")(0).Text
          bairro = valor.getElementsByTagName("bairro")(0).Text
          cidade = valor.getElementsByTagName("localidade")(0).Text
          uf = valor.getElementsByTagName("uf")(0).Text
          
       End If
       
       txtBairro = bairro
       txtRua = rua
       txtCidade = cidade
       txtUF = uf
       
    End If
   Else
      flag = 0
   End If
End Sub

Public Sub limpaCampos()
   txtApelido = Empty
   txtCEP = Empty
   txtBairro = Empty
   txtCidade = Empty
   txtRua = Empty
   txtUF = Empty
   txtNumero = Empty
   txtComplemento = Empty
    
End Sub
