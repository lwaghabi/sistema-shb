Attribute VB_Name = "mRibbon"
Sub Inicio(f As Form, Ribbon As ACPRibbon, iml As ImageList)
Dim s As cRibbon: Set s = New cRibbon
s.Inicio f, Ribbon, iml: Set s = Nothing
End Sub
