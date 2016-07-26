<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNameEntry
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
        Me.lblFirstName = New System.Windows.Forms.TextBox()
        Me.lblNameMessage = New System.Windows.Forms.Label()
        Me.btnConfirmName = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(30, 25)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(100, 20)
        Me.lblFirstName.TabIndex = 1
        '
        'lblNameMessage
        '
        Me.lblNameMessage.AutoSize = True
        Me.lblNameMessage.Location = New System.Drawing.Point(12, 9)
        Me.lblNameMessage.Name = "lblNameMessage"
        Me.lblNameMessage.Size = New System.Drawing.Size(141, 13)
        Me.lblNameMessage.TabIndex = 2
        Me.lblNameMessage.Text = "Please Enter your first name."
        '
        'btnConfirmName
        '
        Me.btnConfirmName.Location = New System.Drawing.Point(44, 51)
        Me.btnConfirmName.Name = "btnConfirmName"
        Me.btnConfirmName.Size = New System.Drawing.Size(75, 23)
        Me.btnConfirmName.TabIndex = 3
        Me.btnConfirmName.Text = "Ok"
        Me.btnConfirmName.UseVisualStyleBackColor = True
        '
        'frmNameEntry
        '
        Me.AcceptButton = Me.btnConfirmName
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(163, 91)
        Me.Controls.Add(Me.btnConfirmName)
        Me.Controls.Add(Me.lblNameMessage)
        Me.Controls.Add(Me.lblFirstName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNameEntry"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Enter your name"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFirstName As TextBox
    Friend WithEvents lblNameMessage As Label
    Friend WithEvents btnConfirmName As Button
End Class
