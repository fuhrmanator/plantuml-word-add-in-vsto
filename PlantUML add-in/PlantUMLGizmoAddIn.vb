Imports Microsoft.Office.Tools
Imports System.Reflection

Public Class PlantUMLGizmoAddIn

    Private myUserControl1 As PlantUML_editor
    Private myCustomTaskPane As CustomTaskPane
    Private myGoogleTracker As GoogleTracker = New GoogleTracker("UA-2895169-10")


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        myUserControl1 = New PlantUML_editor
        myCustomTaskPane = Me.CustomTaskPanes.Add(myUserControl1, "PlantUML Gizmo")
        myCustomTaskPane.Visible = My.Settings.ShowTaskPaneSetting
        myCustomTaskPane.Width = 300
        If myCustomTaskPane.Visible Then ga.trackApp("PlantUML-gizmo-word", versionInfo, "", "", "task-pane")
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Friend ReadOnly Property TaskPane As CustomTaskPane
        Get
            Return myCustomTaskPane
        End Get
    End Property
    Friend ReadOnly Property ga As GoogleTracker
        Get
            Return myGoogleTracker
        End Get
    End Property
    Friend ReadOnly Property versionInfo As String
        Get
            Return getVersionInfo()
        End Get
    End Property

    Private Function getVersionInfo() As String
        Dim versionInfo As String = "unknown"
        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            versionInfo = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Else
            versionInfo = Assembly.GetExecutingAssembly.GetName.Version.ToString
        End If
        '' Initialize the version number
        'If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
        '    versionInfo = "unknown Network Deployed"

        '    'This is a ClickOnce Application
        '    If Not System.Diagnostics.Debugger.IsAttached Then
        '        versionInfo = System.Deployment.Application _
        '            .ApplicationDeployment.CurrentDeployment _
        '                .CurrentVersion.ToString()

        '    End If
        'End If
        System.Diagnostics.Debug.Print("Version info: " & versionInfo)

        Return versionInfo
        '        Return "1.0"
    End Function
End Class
