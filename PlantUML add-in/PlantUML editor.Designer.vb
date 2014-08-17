<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PlantUML_editor
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Preview = New System.Windows.Forms.WebBrowser()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.PreviewButton = New System.Windows.Forms.Button()
        Me.InsertButton = New System.Windows.Forms.Button()
        Me.EditSelectedImageButton = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SourceCode = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel()
        Me.FlowLayoutPanel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(76, 0)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(201, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-3, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select sample"
        '
        'Preview
        '
        Me.Preview.AllowWebBrowserDrop = False
        Me.Preview.IsWebBrowserContextMenuEnabled = False
        Me.Preview.Location = New System.Drawing.Point(3, 32)
        Me.Preview.MinimumSize = New System.Drawing.Size(20, 20)
        Me.Preview.Name = "Preview"
        Me.Preview.Size = New System.Drawing.Size(274, 375)
        Me.Preview.TabIndex = 3
        Me.Preview.Url = New System.Uri("http://www.plantuml.com/plantuml/png/SyfFKj2rKt3CoKnELR1Io4ZDoSa70000", System.UriKind.Absolute)
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.PreviewButton)
        Me.FlowLayoutPanel1.Controls.Add(Me.InsertButton)
        Me.FlowLayoutPanel1.Controls.Add(Me.EditSelectedImageButton)
        Me.FlowLayoutPanel1.Controls.Add(Me.Preview)
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 162)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(280, 407)
        Me.FlowLayoutPanel1.TabIndex = 4
        '
        'PreviewButton
        '
        Me.PreviewButton.Location = New System.Drawing.Point(3, 3)
        Me.PreviewButton.Name = "PreviewButton"
        Me.PreviewButton.Size = New System.Drawing.Size(55, 23)
        Me.PreviewButton.TabIndex = 0
        Me.PreviewButton.Text = "Preview"
        Me.PreviewButton.UseVisualStyleBackColor = True
        '
        'InsertButton
        '
        Me.InsertButton.Location = New System.Drawing.Point(64, 3)
        Me.InsertButton.Name = "InsertButton"
        Me.InsertButton.Size = New System.Drawing.Size(47, 23)
        Me.InsertButton.TabIndex = 1
        Me.InsertButton.Text = "Insert"
        Me.InsertButton.UseVisualStyleBackColor = True
        '
        'EditSelectedImageButton
        '
        Me.EditSelectedImageButton.Location = New System.Drawing.Point(117, 3)
        Me.EditSelectedImageButton.Name = "EditSelectedImageButton"
        Me.EditSelectedImageButton.Size = New System.Drawing.Size(81, 23)
        Me.EditSelectedImageButton.TabIndex = 2
        Me.EditSelectedImageButton.Text = "Edit Selected"
        Me.EditSelectedImageButton.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(-3, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "PlantUML source"
        '
        'SourceCode
        '
        Me.SourceCode.AcceptsReturn = True
        Me.SourceCode.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SourceCode.Location = New System.Drawing.Point(0, 41)
        Me.SourceCode.Multiline = True
        Me.SourceCode.Name = "SourceCode"
        Me.SourceCode.Size = New System.Drawing.Size(277, 115)
        Me.SourceCode.TabIndex = 6
        Me.SourceCode.Text = "Bob->Alice : hello"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(113, 575)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(153, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "PlantUML Gizmo for Word"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.PlantUML_add_in.My.Resources.Resources.ETS_rouge_devise_ecran_fond_transparent
        Me.PictureBox1.Location = New System.Drawing.Point(0, 575)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(111, 74)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(113, 588)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(140, 13)
        Me.LinkLabel1.TabIndex = 9
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "©2014 Christopher Fuhrman"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(113, 601)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(153, 13)
        Me.LinkLabel2.TabIndex = 10
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "christopher.fuhrman@etsmtl.ca"
        '
        'PlantUML_editor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SourceCode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Name = "PlantUML_editor"
        Me.Size = New System.Drawing.Size(280, 731)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Preview As System.Windows.Forms.WebBrowser
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents PreviewButton As System.Windows.Forms.Button
    Friend WithEvents InsertButton As System.Windows.Forms.Button
    Friend WithEvents EditSelectedImageButton As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SourceCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel

End Class
