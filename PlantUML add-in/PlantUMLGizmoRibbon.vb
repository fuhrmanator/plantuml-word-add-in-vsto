Imports Microsoft.Office.Tools.Ribbon

Public Class PlantUMLGizmoRibbon

    Private Sub PlantUMLGizmoRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        DisplayTaskPane.Checked = Globals.PlantUMLGizmoAddIn.TaskPane.Visible
        AllowAnalytics.Checked = My.Settings.AllowAnalyticsSetting

    End Sub

    Private Sub DisplayTaskPane_Click(sender As Object, e As RibbonControlEventArgs) Handles DisplayTaskPane.Click
        Globals.PlantUMLGizmoAddIn.TaskPane.Visible = Not Globals.PlantUMLGizmoAddIn.TaskPane.Visible
        DisplayTaskPane.Checked = Globals.PlantUMLGizmoAddIn.TaskPane.Visible
        My.Settings.ShowTaskPaneSetting = DisplayTaskPane.Checked
        My.Settings.Save()

        'Track event
        Globals.PlantUMLGizmoAddIn.ga.trackApp("PlantUML-gizmo-word", Globals.PlantUMLGizmoAddIn.versionInfo, "", If(DisplayTaskPane.Checked, "displayed", "hidden"), "toggle-task-pane")
    End Sub

    Private Sub AllowAnalytics_Click(sender As Object, e As RibbonControlEventArgs) Handles AllowAnalytics.Click
        'Track event
        Globals.PlantUMLGizmoAddIn.ga.trackApp("PlantUML-gizmo-word", Globals.PlantUMLGizmoAddIn.versionInfo, "", If(My.Settings.AllowAnalyticsSetting, "disabled", "enabled"), "toggle-analytics")
        My.Settings.AllowAnalyticsSetting = Not My.Settings.AllowAnalyticsSetting
        My.Settings.Save()

    End Sub

End Class
