Module modGlobal
    'Зарезервировать перемененые  
    Public app As Microsoft.Office.Interop.MSProject.Application = Nothing
    Public proj As Microsoft.Office.Interop.MSProject.Project = Nothing

    Public Sub load_project()
        'Инициализация приложения
        If app Is Nothing Then
            app = Globals.ThisAddIn.Application
        End If

        'Инициализация проекта
        If proj Is Nothing Then
            proj = app.ActiveProject
        End If

    End Sub
End Module
