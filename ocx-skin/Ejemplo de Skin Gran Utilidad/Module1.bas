Attribute VB_Name = "Module1"
Public Sub Aplicar_skin(ByVal Formulario As Form)
    Form1.Skin1.LoadSkin App.Path & "\Skins\winaqua.skn"
   Form1.Skin1.ApplySkin Formulario.hWnd
End Sub
