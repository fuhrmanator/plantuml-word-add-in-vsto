Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Linq
Imports System.Net
Imports System.Text
Imports System.Web
Public Class GoogleTracker

    ' See https://developers.google.com/analytics/devguides/collection/protocol/v1/devguide
    Private googleURL As String = "http://www.google-analytics.com/collect"
    Private googleVersion As String = "1"
    Private googleTrackingID As String = "UA-XXXX-Y"
    Private googleClientID = "555" ' Anonymous client ID. 

    Public Sub New(trackingID As String)
        googleTrackingID = trackingID
    End Sub

    Public Sub trackApp(appName As String, appVersion As String, appID As String, appInstallerID As String, screenName As String)
        Dim ht As NameValueCollection = baseValues()
        ht.Add("t", "screenview")
        ht.Add("an", appName)
        ht.Add("av", appVersion)
        If appID <> Nothing Then ht.Add("aid", appID)
        If appInstallerID <> Nothing Then ht.Add("aiid", appInstallerID)
        ht.Add("cd", screenName)

        postData(ht)

    End Sub
    Public Sub trackEvent(category As String, action As String, label As String, value As String)
        Dim ht As NameValueCollection = baseValues()
        ht.Add("t", "event")
        ht.Add("ec", category)
        ht.Add("ea", action)
        If label <> Nothing Then ht.Add("el", label)
        If value <> Nothing Then ht.Add("ev", value)

        postData(ht)

    End Sub
    Private Function baseValues() As NameValueCollection
        Dim ht As NameValueCollection = New NameValueCollection
        ht.Add("v", googleVersion)
        ht.Add("tid", googleTrackingID)
        ht.Add("cid", googleClientID)
        Return ht
    End Function
    Public Function postData(values As NameValueCollection) As Boolean
        Dim result As String
        Dim data As String = ""
        For Each key As String In values.Keys
            If data <> "" Then data &= "&"
            If values(key) <> Nothing Then data &= key.ToString() & "=" & HttpUtility.UrlEncode(values(key).ToString())
        Next
        Using client As New WebClient
            result = client.UploadString(googleURL, "POST", data)
        End Using
        Return True
    End Function

End Class
