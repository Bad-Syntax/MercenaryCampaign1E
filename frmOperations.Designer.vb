<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmOperations
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
        Me.labUnitStatus = New System.Windows.Forms.Label()
        Me.cmdOverhead = New System.Windows.Forms.Button()
        Me.labContract = New System.Windows.Forms.Label()
        Me.cmdAdvance = New System.Windows.Forms.Button()
        Me.cmdSalvage = New System.Windows.Forms.Button()
        Me.cmdFindContract = New System.Windows.Forms.Button()
        Me.cmdBreak = New System.Windows.Forms.Button()
        Me.cmdSupplies = New System.Windows.Forms.Button()
        Me.cmdPaySalaries = New System.Windows.Forms.Button()
        Me.cmdProvideMaintenance = New System.Windows.Forms.Button()
        Me.cmdProvideSupplies = New System.Windows.Forms.Button()
        Me.cmdPayOverhead = New System.Windows.Forms.Button()
        Me.labForces = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtEvents = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'labUnitStatus
        '
        Me.labUnitStatus.AutoSize = True
        Me.labUnitStatus.Location = New System.Drawing.Point(12, 137)
        Me.labUnitStatus.Name = "labUnitStatus"
        Me.labUnitStatus.Size = New System.Drawing.Size(101, 25)
        Me.labUnitStatus.TabIndex = 0
        Me.labUnitStatus.Text = "Unit Status:"
        '
        'cmdOverhead
        '
        Me.cmdOverhead.Location = New System.Drawing.Point(12, 52)
        Me.cmdOverhead.Name = "cmdOverhead"
        Me.cmdOverhead.Size = New System.Drawing.Size(173, 34)
        Me.cmdOverhead.TabIndex = 1
        Me.cmdOverhead.Text = "Set &Overhead"
        Me.cmdOverhead.UseVisualStyleBackColor = True
        '
        'labContract
        '
        Me.labContract.AutoSize = True
        Me.labContract.Location = New System.Drawing.Point(945, 92)
        Me.labContract.Name = "labContract"
        Me.labContract.Size = New System.Drawing.Size(146, 25)
        Me.labContract.TabIndex = 2
        Me.labContract.Text = "Current Contract:"
        '
        'cmdAdvance
        '
        Me.cmdAdvance.Enabled = False
        Me.cmdAdvance.Location = New System.Drawing.Point(473, 172)
        Me.cmdAdvance.Name = "cmdAdvance"
        Me.cmdAdvance.Size = New System.Drawing.Size(186, 34)
        Me.cmdAdvance.TabIndex = 9
        Me.cmdAdvance.Text = "&Advance"
        Me.cmdAdvance.UseVisualStyleBackColor = True
        '
        'cmdSalvage
        '
        Me.cmdSalvage.Location = New System.Drawing.Point(12, 92)
        Me.cmdSalvage.Name = "cmdSalvage"
        Me.cmdSalvage.Size = New System.Drawing.Size(173, 34)
        Me.cmdSalvage.TabIndex = 2
        Me.cmdSalvage.Text = "&Modify Unit"
        Me.cmdSalvage.UseVisualStyleBackColor = True
        '
        'cmdFindContract
        '
        Me.cmdFindContract.Location = New System.Drawing.Point(945, 52)
        Me.cmdFindContract.Name = "cmdFindContract"
        Me.cmdFindContract.Size = New System.Drawing.Size(146, 34)
        Me.cmdFindContract.TabIndex = 11
        Me.cmdFindContract.Text = "&Find Contract"
        Me.cmdFindContract.UseVisualStyleBackColor = True
        '
        'cmdBreak
        '
        Me.cmdBreak.Location = New System.Drawing.Point(945, 12)
        Me.cmdBreak.Name = "cmdBreak"
        Me.cmdBreak.Size = New System.Drawing.Size(146, 34)
        Me.cmdBreak.TabIndex = 10
        Me.cmdBreak.Text = "&Break Contract"
        Me.cmdBreak.UseVisualStyleBackColor = True
        Me.cmdBreak.Visible = False
        '
        'cmdSupplies
        '
        Me.cmdSupplies.Location = New System.Drawing.Point(12, 12)
        Me.cmdSupplies.Name = "cmdSupplies"
        Me.cmdSupplies.Size = New System.Drawing.Size(173, 34)
        Me.cmdSupplies.TabIndex = 0
        Me.cmdSupplies.Text = "Order &Supplies"
        Me.cmdSupplies.UseVisualStyleBackColor = True
        '
        'cmdPaySalaries
        '
        Me.cmdPaySalaries.Enabled = False
        Me.cmdPaySalaries.Location = New System.Drawing.Point(473, 52)
        Me.cmdPaySalaries.Name = "cmdPaySalaries"
        Me.cmdPaySalaries.Size = New System.Drawing.Size(186, 34)
        Me.cmdPaySalaries.TabIndex = 6
        Me.cmdPaySalaries.Text = "Pay Salaries"
        Me.cmdPaySalaries.UseVisualStyleBackColor = True
        '
        'cmdProvideMaintenance
        '
        Me.cmdProvideMaintenance.Enabled = False
        Me.cmdProvideMaintenance.Location = New System.Drawing.Point(473, 132)
        Me.cmdProvideMaintenance.Name = "cmdProvideMaintenance"
        Me.cmdProvideMaintenance.Size = New System.Drawing.Size(186, 34)
        Me.cmdProvideMaintenance.TabIndex = 8
        Me.cmdProvideMaintenance.Text = "Provide Maintenance"
        Me.cmdProvideMaintenance.UseVisualStyleBackColor = True
        '
        'cmdProvideSupplies
        '
        Me.cmdProvideSupplies.Location = New System.Drawing.Point(473, 12)
        Me.cmdProvideSupplies.Name = "cmdProvideSupplies"
        Me.cmdProvideSupplies.Size = New System.Drawing.Size(186, 34)
        Me.cmdProvideSupplies.TabIndex = 5
        Me.cmdProvideSupplies.Text = "Provide Supplies"
        Me.cmdProvideSupplies.UseVisualStyleBackColor = True
        '
        'cmdPayOverhead
        '
        Me.cmdPayOverhead.Enabled = False
        Me.cmdPayOverhead.Location = New System.Drawing.Point(473, 92)
        Me.cmdPayOverhead.Name = "cmdPayOverhead"
        Me.cmdPayOverhead.Size = New System.Drawing.Size(186, 34)
        Me.cmdPayOverhead.TabIndex = 7
        Me.cmdPayOverhead.Text = "Pay Overhead"
        Me.cmdPayOverhead.UseVisualStyleBackColor = True
        '
        'labForces
        '
        Me.labForces.AutoSize = True
        Me.labForces.Location = New System.Drawing.Point(473, 209)
        Me.labForces.Name = "labForces"
        Me.labForces.Size = New System.Drawing.Size(63, 25)
        Me.labForces.TabIndex = 12
        Me.labForces.Text = "Forces"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 756)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 25)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Events:"
        '
        'txtEvents
        '
        Me.txtEvents.Location = New System.Drawing.Point(12, 784)
        Me.txtEvents.Multiline = True
        Me.txtEvents.Name = "txtEvents"
        Me.txtEvents.ReadOnly = True
        Me.txtEvents.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtEvents.Size = New System.Drawing.Size(1234, 172)
        Me.txtEvents.TabIndex = 14
        '
        'frmOperations
        '
        Me.AcceptButton = Me.cmdAdvance
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1258, 968)
        Me.Controls.Add(Me.txtEvents)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.labForces)
        Me.Controls.Add(Me.cmdProvideSupplies)
        Me.Controls.Add(Me.cmdPayOverhead)
        Me.Controls.Add(Me.cmdProvideMaintenance)
        Me.Controls.Add(Me.cmdPaySalaries)
        Me.Controls.Add(Me.cmdSupplies)
        Me.Controls.Add(Me.cmdBreak)
        Me.Controls.Add(Me.cmdFindContract)
        Me.Controls.Add(Me.cmdSalvage)
        Me.Controls.Add(Me.cmdAdvance)
        Me.Controls.Add(Me.labContract)
        Me.Controls.Add(Me.cmdOverhead)
        Me.Controls.Add(Me.labUnitStatus)
        Me.Name = "frmOperations"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Monthly Mercenary Operations"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents labUnitStatus As Label
    Friend WithEvents cmdOverhead As Button
    Friend WithEvents labContract As Label
    Friend WithEvents cmdAdvance As Button
    Friend WithEvents cmdSalvage As Button
    Friend WithEvents cmdFindContract As Button
    Friend WithEvents cmdBreak As Button
    Friend WithEvents cmdSupplies As Button
    Friend WithEvents cmdPaySalaries As Button
    Friend WithEvents cmdProvideMaintenance As Button
    Friend WithEvents cmdProvideSupplies As Button
    Friend WithEvents cmdPayOverhead As Button
    Friend WithEvents labForces As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtEvents As TextBox
End Class
