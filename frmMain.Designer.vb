<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.cmdNew = New System.Windows.Forms.Button()
        Me.cmdContinue = New System.Windows.Forms.Button()
        Me.cmdLoad = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.SuspendLayout()
        '
        'cmdNew
        '
        Me.cmdNew.Location = New System.Drawing.Point(15, 15)
        Me.cmdNew.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.Size = New System.Drawing.Size(315, 36)
        Me.cmdNew.TabIndex = 0
        Me.cmdNew.Text = "Create a new Mercenary Unit"
        Me.cmdNew.UseVisualStyleBackColor = True
        '
        'cmdContinue
        '
        Me.cmdContinue.Location = New System.Drawing.Point(15, 59)
        Me.cmdContinue.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdContinue.Name = "cmdContinue"
        Me.cmdContinue.Size = New System.Drawing.Size(315, 36)
        Me.cmdContinue.TabIndex = 1
        Me.cmdContinue.Text = "Continue my Mercenary Campaign"
        Me.cmdContinue.UseVisualStyleBackColor = True
        '
        'cmdLoad
        '
        Me.cmdLoad.Location = New System.Drawing.Point(15, 102)
        Me.cmdLoad.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdLoad.Name = "cmdLoad"
        Me.cmdLoad.Size = New System.Drawing.Size(315, 36)
        Me.cmdLoad.TabIndex = 2
        Me.cmdLoad.Text = "Load Mercenary Campaign"
        Me.cmdLoad.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(15, 146)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(315, 36)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "Abandon My Responsibilities!"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'ofd
        '
        Me.ofd.DefaultExt = "*.sav"
        Me.ofd.FileName = "MyMercUnit.sav"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(345, 201)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdLoad)
        Me.Controls.Add(Me.cmdContinue)
        Me.Controls.Add(Me.cmdNew)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Mercenary Campaign 1E"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdNew As Button
    Friend WithEvents cmdContinue As Button
    Friend WithEvents cmdLoad As Button
    Friend WithEvents cmdExit As Button
    Friend WithEvents ofd As OpenFileDialog
End Class
