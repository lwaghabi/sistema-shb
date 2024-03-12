Attribute VB_Name = "RetornaQtdDiasMes"

Public Function calculaDiasMes(mes As Integer, ano As Integer) As Integer

   Select Case mes
      Case 2
         If (ano Mod 4 = 0) Then
            calculaDiasMes = 29
         Else
            calculaDiasMes = 28
         End If
      Case 4, 6, 9, 11
            calculaDiasMes = 30
      Case Else
            calculaDiasMes = 31
   End Select
End Function
