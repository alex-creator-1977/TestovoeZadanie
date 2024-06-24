Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon

Public Class ribAddIn

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnSetResource_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSetResource.Click
        modGlobal.load_project()

        If MessageBox.Show("Все задачи с 'исполнитель' в названии получат выбранный ресурс. Продолжить с '" & cboResources.Text & "'?", "УЧАСТНИК", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            'опросить все задачи на предмет определённого шаблона в названии
            For Each task As Microsoft.Office.Interop.MSProject.Task In proj.Tasks

                'при положительном исходе добавить в ресурсы определённыхй ресурс

                If task.Name.Contains("исполнитель") Then

                    task.ResourceNames = cboResources.Text

                    'proj.save
                End If

            Next

        End If

        Try

        Catch ex As Exception
            If app Is Nothing Then
                'реакция и лог
            End If

            If proj Is Nothing Then
                'реакция и лог
            End If
        End Try
    End Sub

    Private Sub btnLoadResources_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLoadResources.Click
        Dim activproj As Microsoft.Office.Interop.MSProject.Project = Nothing
        Try

            ' 

            activproj = Globals.ThisAddIn.Application.ActiveProject

            For Each res As Microsoft.Office.Interop.MSProject.Resource In activproj.Resources


                Dim item = Globals.Factory.GetRibbonFactory.CreateRibbonDropDownItem
                item.Label = res.Name

                cboResources.Items.Add(item)

            Next
            If cboResources.Items.Count > 0 Then

                cboResources.Text = cboResources.Items(0).Label

            End If
            MsgBox("Найдено " & cboResources.Items.Count & " ресурс(ов)!")
        Catch ex As Exception
            If activproj Is Nothing Then
                'реакция и лог
            End If

        End Try
    End Sub
End Class
