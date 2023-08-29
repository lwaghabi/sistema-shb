Attribute VB_Name = "ExportarWord"

Public Function ExportarWord(responsavel, Cliente, medidaEquip, tamanhoTotal, qtdFunc, valorFunc, QtdDias, proposta, revisao)
        Dim CaminhoNew As String
                
        CaminhoNew = "C:\Meus Documentos\SISTEMA SHB" & "\docPadrao\"
        
        Dim wordObj As Word.Application
        Dim arqProp As Word.Document
        Dim conteudoDoc As Word.Selection
   
        Set wordObj = CreateObject("Word.Application")
        
        wordObj.Visible = True
        
        Set arqProp = wordObj.Documents.Open(CaminhoNew & "ModelWord.docx")
        Set conteudoDoc = arqProp.Application.Selection
        
        conteudoDoc.Find.Text = "#DESTINATARIO"
        conteudoDoc.Find.Replacement.Text = responsavel
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        conteudoDoc.Find.Text = "#CLIENTE_CONTRATO"
        conteudoDoc.Find.Replacement.Text = Cliente
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        

490     arqProp.SaveAs (CaminhoNew & "NovoWord.docx")
510     arqProp.Close
520     Set arqProp = Nothing
530     wordObj.Quit
540     Set wordObj = Nothing

Exit Function
Erro:
MsgBox "Ocorreu o seguinte erro:" + vbNewLine & _
        Err.Description + vbNewLine & "Na linha: " & Erl
End Function






