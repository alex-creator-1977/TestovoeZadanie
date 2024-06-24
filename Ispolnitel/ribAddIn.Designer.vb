Partial Class ribAddIn
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()



    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    'Das Bearbeiten ist mit dem Komponenten-Designer möglich.
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btnSetResource = Me.Factory.CreateRibbonButton
        Me.btnLoadResources = Me.Factory.CreateRibbonButton
        Me.cboResources = Me.Factory.CreateRibbonComboBox
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.DialogLauncher = RibbonDialogLauncherImpl1
        Me.Group1.Items.Add(Me.btnLoadResources)
        Me.Group1.Items.Add(Me.cboResources)
        Me.Group1.Items.Add(Me.btnSetResource)
        Me.Group1.Label = "Ресурсы"
        Me.Group1.Name = "Group1"
        '
        'btnSetResource
        '
        Me.btnSetResource.Label = "назначить"
        Me.btnSetResource.Name = "btnSetResource"
        '
        'btnLoadResources
        '
        Me.btnLoadResources.Label = "закачать"
        Me.btnLoadResources.Name = "btnLoadResources"
        '
        'cboResources
        '
        Me.cboResources.Label = "доступные"
        Me.cboResources.Name = "cboResources"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Project.Project"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnSetResource As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLoadResources As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents cboResources As Microsoft.Office.Tools.Ribbon.RibbonComboBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property Ribbon1() As ribAddIn
        Get
            Return Me.GetRibbon(Of ribAddIn)()
        End Get
    End Property
End Class
