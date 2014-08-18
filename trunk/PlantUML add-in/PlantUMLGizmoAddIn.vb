Imports Microsoft.Office.Tools

Public Class PlantUMLGizmoAddIn

    Private myUserControl1 As PlantUML_editor
    Private myCustomTaskPane As CustomTaskPane

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        myUserControl1 = New PlantUML_editor
        myCustomTaskPane = Me.CustomTaskPanes.Add(myUserControl1, "PlantUML Gizmo")
        myCustomTaskPane.Visible = True
        myCustomTaskPane.Width = 300
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Friend ReadOnly Property TaskPane As CustomTaskPane
        Get
            Return myCustomTaskPane
        End Get
    End Property
End Class
