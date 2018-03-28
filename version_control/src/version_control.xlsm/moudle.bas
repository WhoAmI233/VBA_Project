Attribute VB_Name = "moudle"
Public Sub qiuguochao1()
    Dim proj_name As String
    proj_name = "vbaProject"

    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.importVbaCode vbaProject
End Sub
