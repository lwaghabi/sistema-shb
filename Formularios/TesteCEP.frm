VERSION 5.00
Begin VB.Form TesteCEP 
   Caption         =   "TesteCEP"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
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
      Left            =   8040
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
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
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1440
      Width           =   4455
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
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "TesteCEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBuscar_Click()

    Dim obj As MSXML2.ServerXMLHTTP
    Dim objLerXml As MSXML2.DOMDocument
    
    
    'Dim ws As New WebService
    Dim strURL As String
    Dim strResponse As String
    Dim strRua As String
    
    Set obj = New MSXML2.ServerXMLHTTP
    Set strURL = New MSXML2.DOMDocument

    ' Validar CEP
    If Len(Trim(txtCEP.Text)) = 0 Then
        MsgBox "Informe o CEP.", vbExclamation, "Erro"
        Exit Sub
    End If

    ' Montar URL da API
    strURL = "https://viacep.com.br/ws/" & txtCEP.Text & "/json/"

    ' Acessar API
    ws.Open strURL
    strResponse = ws.responseText
    ws.Close

    ' Extrair rua da resposta
    'strRua = Mid(strResponse, InStr(strResponse, `"logradouro":`), InStr(strResponse, `","bairro":`) - InStr(strResponse, `"logradouro":`))
    'strRua = Mid(strResponse, InStr(strResponse, `"logradouro":`)
    'strRua = Replace(strRua, `\"`, "")

    ' Exibir rua
    txtRua.Text = strRua
End Sub


