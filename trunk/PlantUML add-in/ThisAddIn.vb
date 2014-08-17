Public Class ThisAddIn

    Private myUserControl1 As PlantUML_editor
    Private myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        myUserControl1 = New PlantUML_editor
        myCustomTaskPane = Me.CustomTaskPanes.Add(myUserControl1, "PlantUML Gizmo")
        myCustomTaskPane.Visible = True
        myCustomTaskPane.Width = 300
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
