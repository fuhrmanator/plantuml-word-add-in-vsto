Imports Microsoft.Office.Tools.Ribbon

Public Class PlantUMLGizmoRibbon

    Private Sub PlantUMLGizmoRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        DisplayTaskPane.Checked = Globals.PlantUMLGizmoAddIn.TaskPane.Visible
    End Sub

    Private Sub DisplayTaskPane_Click(sender As Object, e As RibbonControlEventArgs) Handles DisplayTaskPane.Click
        Globals.PlantUMLGizmoAddIn.TaskPane.Visible = DisplayTaskPane.Checked
        'Track event
        Globals.PlantUMLGizmoAddIn.ga.trackApp("PlantUML-gizmo-word", Globals.PlantUMLGizmoAddIn.versionInfo, "", If(DisplayTaskPane.Checked, "displayed", "hidden"), "toggle-task-pane")
    End Sub
End Class
