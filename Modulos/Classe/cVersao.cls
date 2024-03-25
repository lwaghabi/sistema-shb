VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim blnGerarvsServidor As Boolean

Function IsVersaoOK() As Boolean
10    On Local Error GoTo Erro
20    strVSLocal = "1.3"
'30    AtualizaDados '<<----Só usar para atualizar a versão

40    mdiSHB.Caption = "SHB vs: " & strVSLocal
50    If Not (CriarArquivo) Then Exit Function

60    If strVSLocal <> GetVsServidor Then
70       MsgBox "O sistema precisa ser atualizado!" & vbNewLine & _
                "Você será redirecionado para nosso atendimento online.", vbCritical
               
'80       If IsDeveloper Then Exit Function
'90       AtendimentoOnline "&text=Ol%C3%A1%2C%20estou%20usando%20o%20sistema%20" & _
'                           "SistemaRH%20e%20n%C3%A3o%20consigo%20atualizar%20para" & _
'                           "%20vers%C3%A3o%20mais%20recente."
         'End
100   End If
110   Exit Function
Erro:
MsgBox Err.Description
End Function

Private Sub AtualizaDados()
10 On Local Error GoTo Erro
'20
      If Not (Compilando) Then Exit Sub
30    'If Not (ConfirmaAcao("Deseja atualizar a versão do servidor para: " & _
         strVSLocal & "?")) Then Exit Sub
40   Rotina_AbrirBanco
50       db.Execute ("UPDATE versaosistema SET vs='" & strVSLocal & "'")
60    FechaDB
70    CriarArquivo True
      'CriarArquivoExtra
80    'OKMsg
   End
90 Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Function GetVsOnline() As String
10 On Local Error GoTo Erro
20    Rotina_AbrirBanco
30       rs.Open "SELECT vs FROM versaosistema", db, 3, 3
40       GetVsOnline = "" & rs(0)
50    FechaDB
60 Exit Function
Erro:
69    GetVsOnline = strVSLocal
MsgBox Err.Description
End Function

Private Function GetVsServidor() As String
Dim fso As New Scripting.FileSystemObject
Dim ts As Scripting.TextStream
Dim fName As String

10 On Local Error GoTo Erro
20    fName = App.Path & "\vsServidor.txt"
30    If (Dir(fName, vbArchive) = "") Then _
          GetVsServidor = GetVsOnline: Exit Function

40    Set ts = fso.OpenTextFile(fName, ForReading)
50    GetVsServidor = ts.ReadAll
60    ts.Close
70 Exit Function
Erro:
79    GetVsServidor = strVSLocal
MsgBox Err.Description
End Function

Private Function CriarArquivo(Optional blnGerarvsServidor As Boolean) As Boolean
'Faça uma referência a biblioteca Microsoft Scripting Runtime ( SCRRUN.DLL )
Dim fso As New Scripting.FileSystemObject
Dim ts As Scripting.TextStream
Dim Textfile As String
Dim vsServidor As String

On Local Error GoTo Erro

If (blnGerarvsServidor) Then
   vsServidor = App.Path & "\vsServidor.txt"
   If (Dir(vsServidor, vbArchive) <> "") Then Kill vsServidor
   Set ts = fso.OpenTextFile(vsServidor, ForWriting, True)
   ts.Write strVSLocal
   GoTo Fim
End If

Textfile = App.Path & "\vsLocal.txt"

If (Dir(Textfile, vbArchive) <> "") Then Kill Textfile
'abre um arquivo para escrita , se não existir cria
Set ts = fso.OpenTextFile(Textfile, ForWriting, True)

ts.Write strVSLocal
CriarArquivo = True
'CriarArquivoExtra

Fim:
ts.Close
Set ts = Nothing
Exit Function
Erro:
Set ts = Nothing
MsgBox Err.Description
End Function

'Private Sub CriarArquivoExtra()
''Faça uma referência a biblioteca Microsoft Scripting Runtime ( SCRRUN.DLL )
'      Dim fso As New Scripting.FileSystemObject
'      Dim ts As Scripting.TextStream
'      Dim fName As String
'      Dim fServer As String
'      Dim fExtra As String
'10 On Local Error GoTo Erro
'20    If Not (IsDeveloper) Then Exit Sub
'30    fName = caminho & "vsLocal.txt"
'40    fExtra = caminho & "vs.txt"
'50    fServer = caminho & "vsServidor.txt"
'
'60    If (Dir(fName, vbArchive) <> "") Then Kill fName
'70    If (Dir(fExtra, vbArchive) <> "") Then Kill fExtra
'
'80    Set ts = fso.OpenTextFile(fName, ForWriting, True)
'90    ts.Write strVSLocal
'100   ts.Close
'
'110   Set ts = fso.OpenTextFile(fExtra, ForWriting, True)
'120   ts.Write strVSLocal
'130   ts.Close
'
'140   Set ts = fso.OpenTextFile(fServer, ForWriting, True)
'
'150   ts.Write strVSLocal
'160   ts.Close
'170   Set ts = Nothing
'
'180 Exit Sub
'Erro:
'MsgBox Err.Description
'End Sub



