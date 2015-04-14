Imports System.IO
Imports System.IO.Compression
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Word
Imports System.Timers

Public Class PlantUML_editor
    Private nextPreviewTime As Date
    Dim UPDATE_DELAY_MS As Double = 2000
    'Dim updatePending As Boolean
    Dim updateTimer As Timers.Timer
    Private Sub enableButtons(ByVal state As Boolean)
        InsertButton.Enabled = state
        EditSelectedImageButton.Enabled = state
    End Sub
    Private Sub InsertButtonClick(sender As Object, e As EventArgs) Handles InsertButton.Click
        Dim currentSelection As Word.Selection = Globals.PlantUMLGizmoAddIn.Application.ActiveWindow.Selection
        Dim replaced As Boolean = currentSelection.Type = WdSelectionType.wdSelectionInlineShape
        Dim range As Word.Range = currentSelection.Range

        'System.Diagnostics.Debug.Print("Replaced is " & replaced.ToString)

        enableButtons(False)
        'delete selected image (replace behavior)
        If replaced Then
            currentSelection.Delete(Word.WdUnits.wdCharacter, currentSelection.End - currentSelection.Start)
        End If

        Dim imageURL = getURLFromSource()
        Dim picture = range.InlineShapes.AddPicture(imageURL, True)
        Globals.PlantUMLGizmoAddIn.Application.ActiveDocument.Hyperlinks.Add(picture, imageURL)
        If (Not replaced) Then
            'move cursor after inserted picture
            currentSelection.MoveRight(Word.WdUnits.wdCharacter, 1, Word.WdMovementType.wdMove)
        Else
            'select inserted picture
            range.MoveEnd(Word.WdUnits.wdCharacter, 1)
            range.Select()
        End If

        'If replaced Then
        '    'Move cursor after selection
        '    Dim collapseEnd = Word.WdCollapseDirection.wdCollapseEnd
        '    range.Collapse(collapseEnd)
        'End If
        'Track event
        Globals.PlantUMLGizmoAddIn.ga.trackApp("PlantUML-gizmo-word", Globals.PlantUMLGizmoAddIn.versionInfo, "", If(replaced, "replaced", "inserted"), "insert")
        'Globals.PlantUMLGizmoAddIn.ga.trackApp("video", "insert", If(replaced, "replaced", "inserted"), "")
        enableButtons(True)

    End Sub

    Private Sub PreviewOnChange(sender As Object, e As EventArgs) Handles SourceCode.TextChanged
        ' Previews should only occur at a certain frequency so as not to overload the PlantUML server
        If (updateTimer Is Nothing) Then
            updateTimer = New Timers.Timer(UPDATE_DELAY_MS)
            AddHandler updateTimer.Elapsed, New ElapsedEventHandler(AddressOf UpdateElapsed)
            'System.Diagnostics.Debug.Print("Initialized timer")
        End If
        If (Not updateTimer.Enabled) Then
            'updatePending = True
            updateTimer.Enabled = True
        End If
    End Sub
    Private Sub UpdateElapsed()
        Dim imageURL As String
        imageURL = getURLFromSource()
        Preview.Navigate(imageURL)
        'updatePending = False
        'System.Diagnostics.Debug.Print("Updated image.")
        updateTimer.Enabled = False
    End Sub

    Private Sub EditSelectedImageButtonClick(sender As Object, e As EventArgs) Handles EditSelectedImageButton.Click
        enableButtons(False)
        ' find selected image
        Dim active As Window = Globals.PlantUMLGizmoAddIn.Application.ActiveWindow
        Dim isImageSelected As Boolean = False
        If active.Selection.Type = WdSelectionType.wdSelectionInlineShape Then
            isImageSelected = True
            Dim image = active.Selection.InlineShapes.Item(1)
            ' get image's hyperlink
            'Dim url As Uri
            Dim urlString As String '= url.AbsoluteUri
            urlString = image.Hyperlink.Address
            'url = New Uri("http://www.plantuml.com/plantuml/png/SoWkIImgAStDuV9FoafDBb6mgT7LLN0iAagizCaiBk622Liff1QM9kOKQsXomIM1WX3Pw5Y5r9pKtDIy4fV4aaGK1SMPLQb0FLmEgNafG5i0")
            ' find the part after '/png/'
            Dim location As Integer = urlString.LastIndexOf("/")
            Dim encodedText As String = urlString.Substring(location + 1)
            Dim decodedText As String = Decode64(encodedText)
            ' fix CRLF problems 
            SourceCode.Text = System.Text.RegularExpressions.Regex.Replace(decodedText, "(\r\n|\n|\r)", vbCrLf)
        End If
        'Track event
        Globals.PlantUMLGizmoAddIn.ga.trackApp("PlantUML-gizmo-word", Globals.PlantUMLGizmoAddIn.versionInfo, "", If(isImageSelected, "success", "noSelection"), "edit-selected")
        'Globals.PlantUMLGizmoAddIn.ga.trackEvent("video", "edit", If(isImageSelected, "success", "noSelection"), "")
        enableButtons(True)

    End Sub

    ' VB.net version of the JavaScript source from http://plantuml.sourceforge.net/codejavascript.html
    Private Function Encode64(sourceText As Byte()) As String
        Dim r As String = ""

        For i As Integer = 0 To sourceText.Length - 1 Step 3
            If i + 2 = sourceText.Length Then
                r &= append3bytes(sourceText(i), sourceText(i + 1), 0)
            ElseIf i + 1 = sourceText.Length Then
                r &= append3bytes(sourceText(i), 0, 0)
            Else
                r &= append3bytes(sourceText(i), sourceText(i + 1), sourceText(i + 2))
            End If
        Next
        Return r
    End Function

    Private Function append3bytes(b1 As Integer, b2 As Integer, b3 As Integer) As String
        Dim r As String = ""
        Dim c1, c2, c3, c4 As Integer
        c1 = b1 >> 2
        c2 = ((b1 And &H3) << 4) Or (b2 >> 4)
        c3 = ((b2 And &HF) << 2) Or (b3 >> 6)
        c4 = b3 And &H3F
        r = ""
        r &= encode6bit(c1 And &H3F)
        r &= encode6bit(c2 And &H3F)
        r &= encode6bit(c3 And &H3F)
        r &= encode6bit(c4 And &H3F)
        Return r
    End Function

    Private Function encode6bit(b As Integer) As String
        If b < 10 Then
            Return Chr(48 + b)
        End If
        b -= 10
        If b < 26 Then
            Return Chr(65 + b)
        End If
        b -= 26
        If (b < 26) Then
            Return Chr(97 + b)
        End If
        b -= 26
        If b = 0 Then
            Return "-"
        End If

        If b = 1 Then
            Return "_"
        End If

        Return "?"
    End Function
    Function Decode64(urlText As String) As String
        Dim pos As Integer = 0
        Dim ss As New List(Of Byte)
        Dim c1, c2, c3, c4 As Integer
        For i As Integer = 0 To urlText.Length - 1 Step 4
            c1 = decode6bit(urlText.Substring(i, 1))
            c2 = decode6bit(urlText.Substring(i + 1, 1))
            c3 = decode6bit(urlText.Substring(i + 2, 1))
            c4 = decode6bit(urlText.Substring(i + 3, 1))
            ss.Add((c1 << 2) Or (c2 >> 4))
            ss.Add(((c2 And &HF) << 4) Or (c3 >> 2))
            ss.Add(((c3 And &H3) << 6) Or c4)
        Next
        '        Return System.Web.HttpUtility.UrlDecode(System.Web.HttpUtility.UrlEncode(UnZipStr(ss.ToArray)))
        Return UnZipStr(ss.ToArray)
    End Function
    Private Function decode6bit(c As String) As Integer
        Dim r As Integer = 0
        If c >= "0" And c <= "9" Then
            r = Asc(c.Substring(0, 1)) - 48
        ElseIf c >= "A" And c <= "Z" Then
            r = Asc(c.Substring(0, 1)) - 65 + 10
        ElseIf c >= "a" And c <= "z" Then
            r = Asc(c.Substring(0, 1)) - 97 + 36
        ElseIf c = "-" Then
            r = 62
        ElseIf c = "_" Then
            r = 63
        End If
        Return r
    End Function

    Private Function ZipStr(stringToZip) As Byte()
        'Dim zippedString As String
        Using zipped As MemoryStream = New MemoryStream
            Using gzip As DeflateStream = New DeflateStream(zipped, CompressionMode.Compress)
                Using writer As StreamWriter = New StreamWriter(gzip, System.Text.Encoding.UTF8)
                    writer.Write(stringToZip)
                End Using
            End Using
            'zippedString = System.Text.UTF8Encoding.UTF8.GetString(zipped.ToArray)
            Return zipped.ToArray

        End Using
    End Function
    Private Function UnZipStr(bytesToUnZip As Byte()) As String
        Dim unzippedString As String
        Using unzipped As MemoryStream = New MemoryStream(bytesToUnZip)
            Using gunzip As DeflateStream = New DeflateStream(unzipped, CompressionMode.Decompress)
                Using reader As StreamReader = New StreamReader(gunzip, System.Text.Encoding.ASCII)
                    unzippedString = reader.ReadToEnd
                End Using
            End Using
            Return unzippedString

        End Using
    End Function
    Private Function ThrottlePreview() As Boolean
        Return DateTime.Now < nextPreviewTime
    End Function

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub PlantUML_editor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://profs.etsmtl.ca/cfuhrman")
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        System.Diagnostics.Process.Start("mailto:christopher.fuhrman@etsmtl.ca?subject=PlantUML%20Gizmo%20for%20Word")
    End Sub

    Private Function getURLFromSource() As String
        Return "http://www.plantuml.com/plantuml/png/" & Encode64(ZipStr(System.Web.HttpUtility.UrlDecode(System.Web.HttpUtility.UrlEncode(SourceCode.Text).Replace("+", "%20"))))
    End Function

End Class



