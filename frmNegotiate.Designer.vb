<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNegotiate
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nudPosLength = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.nudPosRemuneration = New System.Windows.Forms.NumericUpDown()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.nudPosGuarantees = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.nudPosCommandRights = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.nudPosSupply = New System.Windows.Forms.NumericUpDown()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.nudPosSalvage = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.nudPosTransport = New System.Windows.Forms.NumericUpDown()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cmdNegotiate = New System.Windows.Forms.Button()
        Me.labDesTransport = New System.Windows.Forms.Label()
        Me.labDesSalvage = New System.Windows.Forms.Label()
        Me.labDesSupply = New System.Windows.Forms.Label()
        Me.labDesCommand = New System.Windows.Forms.Label()
        Me.labDesGuarantees = New System.Windows.Forms.Label()
        Me.labDesRemuneration = New System.Windows.Forms.Label()
        Me.labDesLength = New System.Windows.Forms.Label()
        Me.labContract = New System.Windows.Forms.Label()
        Me.grpNegotiate = New System.Windows.Forms.GroupBox()
        Me.cmdWalk = New System.Windows.Forms.Button()
        Me.labPointsLeft = New System.Windows.Forms.Label()
        Me.cmdAccept = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.nudNegPts = New System.Windows.Forms.NumericUpDown()
        Me.labArea = New System.Windows.Forms.Label()
        Me.cbArea = New System.Windows.Forms.ComboBox()
        Me.lstContracts = New System.Windows.Forms.ListBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cmdReturn = New System.Windows.Forms.Button()
        CType(Me.nudPosLength, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPosRemuneration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPosGuarantees, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPosCommandRights, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPosSupply, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPosSalvage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPosTransport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpNegotiate.SuspendLayout()
        CType(Me.nudNegPts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(92, 91)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 25)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Mission Length"
        '
        'nudPosLength
        '
        Me.nudPosLength.Increment = New Decimal(New Int32() {5, 0, 0, 0})
        Me.nudPosLength.Location = New System.Drawing.Point(12, 89)
        Me.nudPosLength.Name = "nudPosLength"
        Me.nudPosLength.Size = New System.Drawing.Size(74, 31)
        Me.nudPosLength.TabIndex = 0
        Me.nudPosLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudPosLength.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(92, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(122, 25)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Remuneration"
        '
        'nudPosRemuneration
        '
        Me.nudPosRemuneration.Increment = New Decimal(New Int32() {5, 0, 0, 0})
        Me.nudPosRemuneration.Location = New System.Drawing.Point(12, 126)
        Me.nudPosRemuneration.Name = "nudPosRemuneration"
        Me.nudPosRemuneration.Size = New System.Drawing.Size(74, 31)
        Me.nudPosRemuneration.TabIndex = 1
        Me.nudPosRemuneration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudPosRemuneration.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(92, 165)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 25)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Guarantees"
        '
        'nudPosGuarantees
        '
        Me.nudPosGuarantees.Increment = New Decimal(New Int32() {5, 0, 0, 0})
        Me.nudPosGuarantees.Location = New System.Drawing.Point(12, 163)
        Me.nudPosGuarantees.Name = "nudPosGuarantees"
        Me.nudPosGuarantees.Size = New System.Drawing.Size(74, 31)
        Me.nudPosGuarantees.TabIndex = 2
        Me.nudPosGuarantees.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudPosGuarantees.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(92, 202)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(151, 25)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Command Rights"
        '
        'nudPosCommandRights
        '
        Me.nudPosCommandRights.Increment = New Decimal(New Int32() {5, 0, 0, 0})
        Me.nudPosCommandRights.Location = New System.Drawing.Point(12, 200)
        Me.nudPosCommandRights.Name = "nudPosCommandRights"
        Me.nudPosCommandRights.Size = New System.Drawing.Size(74, 31)
        Me.nudPosCommandRights.TabIndex = 3
        Me.nudPosCommandRights.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudPosCommandRights.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(92, 276)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(107, 25)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Supply Fees"
        '
        'nudPosSupply
        '
        Me.nudPosSupply.Increment = New Decimal(New Int32() {5, 0, 0, 0})
        Me.nudPosSupply.Location = New System.Drawing.Point(12, 274)
        Me.nudPosSupply.Name = "nudPosSupply"
        Me.nudPosSupply.Size = New System.Drawing.Size(74, 31)
        Me.nudPosSupply.TabIndex = 5
        Me.nudPosSupply.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudPosSupply.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(92, 313)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 25)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Salvage Rights"
        '
        'nudPosSalvage
        '
        Me.nudPosSalvage.Increment = New Decimal(New Int32() {5, 0, 0, 0})
        Me.nudPosSalvage.Location = New System.Drawing.Point(12, 311)
        Me.nudPosSalvage.Name = "nudPosSalvage"
        Me.nudPosSalvage.Size = New System.Drawing.Size(74, 31)
        Me.nudPosSalvage.TabIndex = 6
        Me.nudPosSalvage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudPosSalvage.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(92, 239)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(126, 25)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Transport Fees"
        '
        'nudPosTransport
        '
        Me.nudPosTransport.Increment = New Decimal(New Int32() {5, 0, 0, 0})
        Me.nudPosTransport.Location = New System.Drawing.Point(12, 237)
        Me.nudPosTransport.Name = "nudPosTransport"
        Me.nudPosTransport.Size = New System.Drawing.Size(74, 31)
        Me.nudPosTransport.TabIndex = 4
        Me.nudPosTransport.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudPosTransport.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(12, 61)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(235, 25)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "Desired Contractual Position"
        '
        'cmdNegotiate
        '
        Me.cmdNegotiate.Location = New System.Drawing.Point(568, 204)
        Me.cmdNegotiate.Name = "cmdNegotiate"
        Me.cmdNegotiate.Size = New System.Drawing.Size(144, 34)
        Me.cmdNegotiate.TabIndex = 7
        Me.cmdNegotiate.Text = "Negotiate"
        Me.cmdNegotiate.UseVisualStyleBackColor = True
        Me.cmdNegotiate.Visible = False
        '
        'labDesTransport
        '
        Me.labDesTransport.AutoSize = True
        Me.labDesTransport.Location = New System.Drawing.Point(261, 241)
        Me.labDesTransport.Name = "labDesTransport"
        Me.labDesTransport.Size = New System.Drawing.Size(24, 25)
        Me.labDesTransport.TabIndex = 45
        Me.labDesTransport.Text = "..."
        '
        'labDesSalvage
        '
        Me.labDesSalvage.AutoSize = True
        Me.labDesSalvage.Location = New System.Drawing.Point(261, 315)
        Me.labDesSalvage.Name = "labDesSalvage"
        Me.labDesSalvage.Size = New System.Drawing.Size(24, 25)
        Me.labDesSalvage.TabIndex = 44
        Me.labDesSalvage.Text = "..."
        '
        'labDesSupply
        '
        Me.labDesSupply.AutoSize = True
        Me.labDesSupply.Location = New System.Drawing.Point(261, 278)
        Me.labDesSupply.Name = "labDesSupply"
        Me.labDesSupply.Size = New System.Drawing.Size(24, 25)
        Me.labDesSupply.TabIndex = 43
        Me.labDesSupply.Text = "..."
        '
        'labDesCommand
        '
        Me.labDesCommand.AutoSize = True
        Me.labDesCommand.Location = New System.Drawing.Point(261, 204)
        Me.labDesCommand.Name = "labDesCommand"
        Me.labDesCommand.Size = New System.Drawing.Size(24, 25)
        Me.labDesCommand.TabIndex = 42
        Me.labDesCommand.Text = "..."
        '
        'labDesGuarantees
        '
        Me.labDesGuarantees.AutoSize = True
        Me.labDesGuarantees.Location = New System.Drawing.Point(261, 167)
        Me.labDesGuarantees.Name = "labDesGuarantees"
        Me.labDesGuarantees.Size = New System.Drawing.Size(24, 25)
        Me.labDesGuarantees.TabIndex = 41
        Me.labDesGuarantees.Text = "..."
        '
        'labDesRemuneration
        '
        Me.labDesRemuneration.AutoSize = True
        Me.labDesRemuneration.Location = New System.Drawing.Point(261, 130)
        Me.labDesRemuneration.Name = "labDesRemuneration"
        Me.labDesRemuneration.Size = New System.Drawing.Size(24, 25)
        Me.labDesRemuneration.TabIndex = 40
        Me.labDesRemuneration.Text = "..."
        '
        'labDesLength
        '
        Me.labDesLength.AutoSize = True
        Me.labDesLength.Location = New System.Drawing.Point(261, 93)
        Me.labDesLength.Name = "labDesLength"
        Me.labDesLength.Size = New System.Drawing.Size(24, 25)
        Me.labDesLength.TabIndex = 39
        Me.labDesLength.Text = "..."
        '
        'labContract
        '
        Me.labContract.AutoSize = True
        Me.labContract.Location = New System.Drawing.Point(12, 9)
        Me.labContract.Name = "labContract"
        Me.labContract.Size = New System.Drawing.Size(186, 25)
        Me.labContract.TabIndex = 46
        Me.labContract.Text = "Contract Negotiations"
        '
        'grpNegotiate
        '
        Me.grpNegotiate.Controls.Add(Me.cmdWalk)
        Me.grpNegotiate.Controls.Add(Me.labPointsLeft)
        Me.grpNegotiate.Controls.Add(Me.cmdAccept)
        Me.grpNegotiate.Controls.Add(Me.Label10)
        Me.grpNegotiate.Controls.Add(Me.Label1)
        Me.grpNegotiate.Controls.Add(Me.nudNegPts)
        Me.grpNegotiate.Controls.Add(Me.labArea)
        Me.grpNegotiate.Controls.Add(Me.cbArea)
        Me.grpNegotiate.Location = New System.Drawing.Point(12, 396)
        Me.grpNegotiate.Name = "grpNegotiate"
        Me.grpNegotiate.Size = New System.Drawing.Size(440, 263)
        Me.grpNegotiate.TabIndex = 50
        Me.grpNegotiate.TabStop = False
        Me.grpNegotiate.Text = "Contract Negotiations"
        Me.grpNegotiate.Visible = False
        '
        'cmdWalk
        '
        Me.cmdWalk.Location = New System.Drawing.Point(196, 223)
        Me.cmdWalk.Name = "cmdWalk"
        Me.cmdWalk.Size = New System.Drawing.Size(112, 34)
        Me.cmdWalk.TabIndex = 53
        Me.cmdWalk.Text = "Walk Away"
        Me.cmdWalk.UseVisualStyleBackColor = True
        '
        'labPointsLeft
        '
        Me.labPointsLeft.AutoSize = True
        Me.labPointsLeft.Location = New System.Drawing.Point(94, 123)
        Me.labPointsLeft.Name = "labPointsLeft"
        Me.labPointsLeft.Size = New System.Drawing.Size(24, 25)
        Me.labPointsLeft.TabIndex = 56
        Me.labPointsLeft.Text = "..."
        Me.labPointsLeft.Visible = False
        '
        'cmdAccept
        '
        Me.cmdAccept.Location = New System.Drawing.Point(314, 223)
        Me.cmdAccept.Name = "cmdAccept"
        Me.cmdAccept.Size = New System.Drawing.Size(119, 34)
        Me.cmdAccept.TabIndex = 10
        Me.cmdAccept.Text = "Accept"
        Me.cmdAccept.UseVisualStyleBackColor = True
        Me.cmdAccept.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(22, 94)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(355, 25)
        Me.Label10.TabIndex = 54
        Me.Label10.Text = "Select number of Negotiating Points to Use"
        Me.Label10.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(22, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(411, 25)
        Me.Label1.TabIndex = 53
        Me.Label1.Text = "Select Contract Part to Negotiate (Priority Matters!)"
        Me.Label1.Visible = False
        '
        'nudNegPts
        '
        Me.nudNegPts.Location = New System.Drawing.Point(22, 121)
        Me.nudNegPts.Name = "nudNegPts"
        Me.nudNegPts.Size = New System.Drawing.Size(66, 31)
        Me.nudNegPts.TabIndex = 52
        Me.nudNegPts.Value = New Decimal(New Int32() {9, 0, 0, 0})
        Me.nudNegPts.Visible = False
        '
        'labArea
        '
        Me.labArea.AutoSize = True
        Me.labArea.Location = New System.Drawing.Point(22, 155)
        Me.labArea.Name = "labArea"
        Me.labArea.Size = New System.Drawing.Size(24, 25)
        Me.labArea.TabIndex = 51
        Me.labArea.Text = "..."
        Me.labArea.Visible = False
        '
        'cbArea
        '
        Me.cbArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbArea.FormattingEnabled = True
        Me.cbArea.Items.AddRange(New Object() {"Mission Length", "Remuneration", "Guarantees", "Command Rights", "Transport Fees", "Supply Fees", "Salvage Rights"})
        Me.cbArea.Location = New System.Drawing.Point(22, 55)
        Me.cbArea.Name = "cbArea"
        Me.cbArea.Size = New System.Drawing.Size(294, 33)
        Me.cbArea.TabIndex = 50
        Me.cbArea.Visible = False
        '
        'lstContracts
        '
        Me.lstContracts.FormattingEnabled = True
        Me.lstContracts.ItemHeight = 25
        Me.lstContracts.Location = New System.Drawing.Point(568, 89)
        Me.lstContracts.Name = "lstContracts"
        Me.lstContracts.Size = New System.Drawing.Size(422, 104)
        Me.lstContracts.TabIndex = 51
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(568, 61)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(255, 25)
        Me.Label11.TabIndex = 52
        Me.Label11.Text = "Contracts available this month:"
        '
        'cmdReturn
        '
        Me.cmdReturn.Location = New System.Drawing.Point(718, 204)
        Me.cmdReturn.Name = "cmdReturn"
        Me.cmdReturn.Size = New System.Drawing.Size(193, 34)
        Me.cmdReturn.TabIndex = 57
        Me.cmdReturn.Text = "Return to Operations"
        Me.cmdReturn.UseVisualStyleBackColor = True
        '
        'frmNegotiate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1030, 742)
        Me.Controls.Add(Me.cmdReturn)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.lstContracts)
        Me.Controls.Add(Me.grpNegotiate)
        Me.Controls.Add(Me.labContract)
        Me.Controls.Add(Me.labDesTransport)
        Me.Controls.Add(Me.labDesSalvage)
        Me.Controls.Add(Me.labDesSupply)
        Me.Controls.Add(Me.labDesCommand)
        Me.Controls.Add(Me.labDesGuarantees)
        Me.Controls.Add(Me.labDesRemuneration)
        Me.Controls.Add(Me.labDesLength)
        Me.Controls.Add(Me.cmdNegotiate)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.nudPosTransport)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.nudPosSalvage)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.nudPosSupply)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.nudPosCommandRights)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.nudPosGuarantees)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.nudPosRemuneration)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.nudPosLength)
        Me.Name = "frmNegotiate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Contract Negotiations"
        CType(Me.nudPosLength, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPosRemuneration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPosGuarantees, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPosCommandRights, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPosSupply, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPosSalvage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPosTransport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpNegotiate.ResumeLayout(False)
        Me.grpNegotiate.PerformLayout()
        CType(Me.nudNegPts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As Label
    Friend WithEvents nudPosLength As NumericUpDown
    Friend WithEvents Label3 As Label
    Friend WithEvents nudPosRemuneration As NumericUpDown
    Friend WithEvents Label4 As Label
    Friend WithEvents nudPosGuarantees As NumericUpDown
    Friend WithEvents Label5 As Label
    Friend WithEvents nudPosCommandRights As NumericUpDown
    Friend WithEvents Label6 As Label
    Friend WithEvents nudPosSupply As NumericUpDown
    Friend WithEvents Label7 As Label
    Friend WithEvents nudPosSalvage As NumericUpDown
    Friend WithEvents Label8 As Label
    Friend WithEvents nudPosTransport As NumericUpDown
    Friend WithEvents Label9 As Label
    Friend WithEvents cmdNegotiate As Button
    Friend WithEvents labDesTransport As Label
    Friend WithEvents labDesSalvage As Label
    Friend WithEvents labDesSupply As Label
    Friend WithEvents labDesCommand As Label
    Friend WithEvents labDesGuarantees As Label
    Friend WithEvents labDesRemuneration As Label
    Friend WithEvents labDesLength As Label
    Friend WithEvents labContract As Label
    Friend WithEvents grpNegotiate As GroupBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents nudNegPts As NumericUpDown
    Friend WithEvents labArea As Label
    Friend WithEvents cbArea As ComboBox
    Friend WithEvents cmdAccept As Button
    Friend WithEvents labPointsLeft As Label
    Friend WithEvents lstContracts As ListBox
    Friend WithEvents Label11 As Label
    Friend WithEvents cmdWalk As Button
    Friend WithEvents cmdReturn As Button
End Class
