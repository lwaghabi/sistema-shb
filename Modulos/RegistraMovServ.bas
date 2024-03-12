Attribute VB_Name = "RegistraMovServ"
Public Sub RegistraMovSv(Grupo As String, Classe As String, codServ As String)
   On Error GoTo Erro
   
   db.Execute ("INSERT INTO servmovimentacaoservicos (grupo,classe,codServ,dataMovimentacao) VALUES ('" & Grupo & "','" & Classe & "','" & codServ & "',NOW())")
   
Exit Sub
Erro: MsgBox ("Erro ao registrar movimentação: " & Err.Description), vbInformation
FechaDB
End Sub
