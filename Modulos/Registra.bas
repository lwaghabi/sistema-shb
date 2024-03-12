Attribute VB_Name = "RegistraMovSup"
Public Sub RegistraMov(Grupo As String, Classe As String, codProd As String, qtd As Integer, operacao As String)
   On Error GoTo Erro
   
   db.Execute ("INSERT INTO supmovimentacaoestoque (grupo,classe,codProd,qtdMovimentado,tipoMovimentacao,dataMovimentacao) VALUES ('" & Grupo & "','" & Classe & "','" & codProd & "','" & qtd & "','" & operacao & "',NOW())")
   
Exit Sub
Erro: MsgBox ("Erro ao registrar movimentação: " & Err.Description), vbInformation
FechaDB
End Sub
