<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnSaveAndLaunch = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnLoadDocument = New System.Windows.Forms.Button()
        Me.btnLoadReferenceDocument = New System.Windows.Forms.Button()
        Me.btnKeepInMemoryAsRefenceAndOpenNew = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnSaveAndLaunch)
        Me.GroupBox1.Controls.Add(Me.FlowLayoutPanel1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 41)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(507, 569)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Variables that can receive values:"
        '
        'btnSaveAndLaunch
        '
        Me.btnSaveAndLaunch.Location = New System.Drawing.Point(6, 540)
        Me.btnSaveAndLaunch.Name = "btnSaveAndLaunch"
        Me.btnSaveAndLaunch.Size = New System.Drawing.Size(323, 23)
        Me.btnSaveAndLaunch.TabIndex = 8
        Me.btnSaveAndLaunch.Text = "Create document based on current values  and launch Word."
        Me.btnSaveAndLaunch.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.AutoScroll = True
        Me.FlowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(6, 26)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(479, 508)
        Me.FlowLayoutPanel1.TabIndex = 4
        '
        'btnLoadDocument
        '
        Me.btnLoadDocument.Location = New System.Drawing.Point(12, 12)
        Me.btnLoadDocument.Name = "btnLoadDocument"
        Me.btnLoadDocument.Size = New System.Drawing.Size(191, 23)
        Me.btnLoadDocument.TabIndex = 9
        Me.btnLoadDocument.Text = "Open new template (close current)"
        Me.btnLoadDocument.UseVisualStyleBackColor = True
        '
        'btnLoadReferenceDocument
        '
        Me.btnLoadReferenceDocument.Location = New System.Drawing.Point(411, 12)
        Me.btnLoadReferenceDocument.Name = "btnLoadReferenceDocument"
        Me.btnLoadReferenceDocument.Size = New System.Drawing.Size(108, 23)
        Me.btnLoadReferenceDocument.TabIndex = 10
        Me.btnLoadReferenceDocument.Text = "Load reference"
        Me.btnLoadReferenceDocument.UseVisualStyleBackColor = True
        '
        'btnKeepInMemoryAsRefenceAndOpenNew
        '
        Me.btnKeepInMemoryAsRefenceAndOpenNew.Location = New System.Drawing.Point(221, 12)
        Me.btnKeepInMemoryAsRefenceAndOpenNew.Name = "btnKeepInMemoryAsRefenceAndOpenNew"
        Me.btnKeepInMemoryAsRefenceAndOpenNew.Size = New System.Drawing.Size(139, 23)
        Me.btnKeepInMemoryAsRefenceAndOpenNew.TabIndex = 11
        Me.btnKeepInMemoryAsRefenceAndOpenNew.Text = "Use as reference for new"
        Me.btnKeepInMemoryAsRefenceAndOpenNew.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(539, 627)
        Me.Controls.Add(Me.btnKeepInMemoryAsRefenceAndOpenNew)
        Me.Controls.Add(Me.btnLoadReferenceDocument)
        Me.Controls.Add(Me.btnLoadDocument)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.Text = "Yasminator"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSaveAndLaunch As System.Windows.Forms.Button
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnLoadDocument As System.Windows.Forms.Button
    Friend WithEvents btnLoadReferenceDocument As System.Windows.Forms.Button
    Friend WithEvents btnKeepInMemoryAsRefenceAndOpenNew As System.Windows.Forms.Button

End Class
