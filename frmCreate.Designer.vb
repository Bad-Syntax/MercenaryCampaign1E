<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCreate
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
        Me.lstBase = New System.Windows.Forms.ListBox()
        Me.lstExperience = New System.Windows.Forms.ListBox()
        Me.lstQuality = New System.Windows.Forms.ListBox()
        Me.lstRegiments = New System.Windows.Forms.ListBox()
        Me.cmdFinished = New System.Windows.Forms.Button()
        Me.labData = New System.Windows.Forms.Label()
        Me.labSquad = New System.Windows.Forms.Label()
        Me.chkHireling = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.nudSupplies = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lstBattalions = New System.Windows.Forms.ListBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lstCompanies = New System.Windows.Forms.ListBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lstLances = New System.Windows.Forms.ListBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lstSquads = New System.Windows.Forms.ListBox()
        Me.cmdRegimentsAdd = New System.Windows.Forms.Button()
        Me.cmdRegimentsRemove = New System.Windows.Forms.Button()
        Me.cmdBattalionsRemove = New System.Windows.Forms.Button()
        Me.cmdBattalionsAdd = New System.Windows.Forms.Button()
        Me.cmdCompaniesRemove = New System.Windows.Forms.Button()
        Me.cmdCompaniesAdd = New System.Windows.Forms.Button()
        Me.cmdLancesRemove = New System.Windows.Forms.Button()
        Me.cmdLancesAdd = New System.Windows.Forms.Button()
        Me.cmdSquadsRemove = New System.Windows.Forms.Button()
        Me.cmdSquadsAdd = New System.Windows.Forms.Button()
        Me.labRegiment = New System.Windows.Forms.Label()
        Me.labBattalion = New System.Windows.Forms.Label()
        Me.labCompany = New System.Windows.Forms.Label()
        Me.labLance = New System.Windows.Forms.Label()
        Me.labSquadTemplate = New System.Windows.Forms.Label()
        Me.cbPreExisting = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        CType(Me.nudSupplies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstBase
        '
        Me.lstBase.FormattingEnabled = True
        Me.lstBase.ItemHeight = 25
        Me.lstBase.Items.AddRange(New Object() {"Light Mech", "Medium Mech", "Heavy Mech", "Assault Mech", "Light Fighter", "Medium Fighter", "Heavy Fighter", "Light LAM", "Medium LAM", "Regular Infantry", "Motorized Infantry", "Jump Infantry", "Scout", "Light Armor", "Heavy Armor", "Artillery", "Aircraft", "Support", "Airmobile Regular Infantry", "Airmobile Motorized Infantry", "Airmobile Jump Infantry", "Airmobile Scout", "Leopard Dropship", "Union Dropship", "Overlord Dropship", "Fury Dropship", "Gazelle Dropship", "Seeker Dropship", "Triumph Dropship", "Condor Dropship", "Excaliber Dropship", "Scout Jumpship", "Invader Jumpship", "Monolith Jumpship", "Star Lord Jumpship", "Merchant Jumpship", "Avenger Dropship", "Leopard CV Dropship", "Intruder Dropship", "Buccaneer Dropship", "Achilles Dropship", "Monarch Dropship", "Fortress Dropship", "Vengeance Dropship", "Mule Dropship", "Mammoth Dropship", "Behemoth Dropship"})
        Me.lstBase.Location = New System.Drawing.Point(879, 13)
        Me.lstBase.Margin = New System.Windows.Forms.Padding(4)
        Me.lstBase.Name = "lstBase"
        Me.lstBase.Size = New System.Drawing.Size(258, 454)
        Me.lstBase.TabIndex = 0
        '
        'lstExperience
        '
        Me.lstExperience.FormattingEnabled = True
        Me.lstExperience.ItemHeight = 25
        Me.lstExperience.Items.AddRange(New Object() {"Green", "Regular", "Veteran", "Elite"})
        Me.lstExperience.Location = New System.Drawing.Point(879, 475)
        Me.lstExperience.Margin = New System.Windows.Forms.Padding(4)
        Me.lstExperience.Name = "lstExperience"
        Me.lstExperience.Size = New System.Drawing.Size(258, 104)
        Me.lstExperience.TabIndex = 3
        '
        'lstQuality
        '
        Me.lstQuality.FormattingEnabled = True
        Me.lstQuality.ItemHeight = 25
        Me.lstQuality.Items.AddRange(New Object() {"New", "Salvaged", "Destroyed"})
        Me.lstQuality.Location = New System.Drawing.Point(879, 587)
        Me.lstQuality.Margin = New System.Windows.Forms.Padding(4)
        Me.lstQuality.Name = "lstQuality"
        Me.lstQuality.Size = New System.Drawing.Size(258, 79)
        Me.lstQuality.TabIndex = 4
        '
        'lstRegiments
        '
        Me.lstRegiments.FormattingEnabled = True
        Me.lstRegiments.ItemHeight = 25
        Me.lstRegiments.Location = New System.Drawing.Point(12, 46)
        Me.lstRegiments.Margin = New System.Windows.Forms.Padding(4)
        Me.lstRegiments.Name = "lstRegiments"
        Me.lstRegiments.Size = New System.Drawing.Size(258, 129)
        Me.lstRegiments.TabIndex = 7
        '
        'cmdFinished
        '
        Me.cmdFinished.Location = New System.Drawing.Point(492, 1056)
        Me.cmdFinished.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdFinished.Name = "cmdFinished"
        Me.cmdFinished.Size = New System.Drawing.Size(122, 36)
        Me.cmdFinished.TabIndex = 8
        Me.cmdFinished.Text = "Finished!"
        Me.cmdFinished.UseVisualStyleBackColor = True
        '
        'labData
        '
        Me.labData.AutoSize = True
        Me.labData.Location = New System.Drawing.Point(1145, 13)
        Me.labData.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.labData.Name = "labData"
        Me.labData.Size = New System.Drawing.Size(109, 25)
        Me.labData.TabIndex = 9
        Me.labData.Text = "Empty Label"
        '
        'labSquad
        '
        Me.labSquad.AutoSize = True
        Me.labSquad.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.labSquad.Location = New System.Drawing.Point(314, 763)
        Me.labSquad.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.labSquad.Name = "labSquad"
        Me.labSquad.Size = New System.Drawing.Size(220, 21)
        Me.labSquad.TabIndex = 10
        Me.labSquad.Text = "Click here to add selected unit."
        '
        'chkHireling
        '
        Me.chkHireling.AutoSize = True
        Me.chkHireling.Location = New System.Drawing.Point(879, 673)
        Me.chkHireling.Name = "chkHireling"
        Me.chkHireling.Size = New System.Drawing.Size(107, 29)
        Me.chkHireling.TabIndex = 11
        Me.chkHireling.Text = "Hireling?"
        Me.chkHireling.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(88, 1021)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(202, 25)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Initial Supplies (Months)"
        '
        'nudSupplies
        '
        Me.nudSupplies.Location = New System.Drawing.Point(15, 1019)
        Me.nudSupplies.Maximum = New Decimal(New Int32() {24, 0, 0, 0})
        Me.nudSupplies.Minimum = New Decimal(New Int32() {1, 0, 0, 0})
        Me.nudSupplies.Name = "nudSupplies"
        Me.nudSupplies.Size = New System.Drawing.Size(67, 31)
        Me.nudSupplies.TabIndex = 13
        Me.nudSupplies.Value = New Decimal(New Int32() {1, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 25)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Regiments:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 192)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(93, 25)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Battalions:"
        '
        'lstBattalions
        '
        Me.lstBattalions.FormattingEnabled = True
        Me.lstBattalions.ItemHeight = 25
        Me.lstBattalions.Location = New System.Drawing.Point(12, 225)
        Me.lstBattalions.Margin = New System.Windows.Forms.Padding(4)
        Me.lstBattalions.Name = "lstBattalions"
        Me.lstBattalions.Size = New System.Drawing.Size(258, 129)
        Me.lstBattalions.TabIndex = 15
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 373)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 25)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Companies:"
        '
        'lstCompanies
        '
        Me.lstCompanies.FormattingEnabled = True
        Me.lstCompanies.ItemHeight = 25
        Me.lstCompanies.Location = New System.Drawing.Point(12, 406)
        Me.lstCompanies.Margin = New System.Windows.Forms.Padding(4)
        Me.lstCompanies.Name = "lstCompanies"
        Me.lstCompanies.Size = New System.Drawing.Size(258, 129)
        Me.lstCompanies.TabIndex = 17
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 549)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 25)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Lances:"
        '
        'lstLances
        '
        Me.lstLances.FormattingEnabled = True
        Me.lstLances.ItemHeight = 25
        Me.lstLances.Location = New System.Drawing.Point(13, 582)
        Me.lstLances.Margin = New System.Windows.Forms.Padding(4)
        Me.lstLances.Name = "lstLances"
        Me.lstLances.Size = New System.Drawing.Size(258, 129)
        Me.lstLances.TabIndex = 19
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 725)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 25)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Squads:"
        '
        'lstSquads
        '
        Me.lstSquads.FormattingEnabled = True
        Me.lstSquads.ItemHeight = 25
        Me.lstSquads.Location = New System.Drawing.Point(14, 758)
        Me.lstSquads.Margin = New System.Windows.Forms.Padding(4)
        Me.lstSquads.Name = "lstSquads"
        Me.lstSquads.Size = New System.Drawing.Size(258, 254)
        Me.lstSquads.TabIndex = 21
        '
        'cmdRegimentsAdd
        '
        Me.cmdRegimentsAdd.Location = New System.Drawing.Point(277, 46)
        Me.cmdRegimentsAdd.Name = "cmdRegimentsAdd"
        Me.cmdRegimentsAdd.Size = New System.Drawing.Size(31, 34)
        Me.cmdRegimentsAdd.TabIndex = 23
        Me.cmdRegimentsAdd.Text = "+"
        Me.cmdRegimentsAdd.UseVisualStyleBackColor = True
        '
        'cmdRegimentsRemove
        '
        Me.cmdRegimentsRemove.Location = New System.Drawing.Point(277, 141)
        Me.cmdRegimentsRemove.Name = "cmdRegimentsRemove"
        Me.cmdRegimentsRemove.Size = New System.Drawing.Size(31, 34)
        Me.cmdRegimentsRemove.TabIndex = 24
        Me.cmdRegimentsRemove.Text = "-"
        Me.cmdRegimentsRemove.UseVisualStyleBackColor = True
        '
        'cmdBattalionsRemove
        '
        Me.cmdBattalionsRemove.Location = New System.Drawing.Point(277, 320)
        Me.cmdBattalionsRemove.Name = "cmdBattalionsRemove"
        Me.cmdBattalionsRemove.Size = New System.Drawing.Size(31, 34)
        Me.cmdBattalionsRemove.TabIndex = 26
        Me.cmdBattalionsRemove.Text = "-"
        Me.cmdBattalionsRemove.UseVisualStyleBackColor = True
        '
        'cmdBattalionsAdd
        '
        Me.cmdBattalionsAdd.Location = New System.Drawing.Point(277, 225)
        Me.cmdBattalionsAdd.Name = "cmdBattalionsAdd"
        Me.cmdBattalionsAdd.Size = New System.Drawing.Size(31, 34)
        Me.cmdBattalionsAdd.TabIndex = 25
        Me.cmdBattalionsAdd.Text = "+"
        Me.cmdBattalionsAdd.UseVisualStyleBackColor = True
        '
        'cmdCompaniesRemove
        '
        Me.cmdCompaniesRemove.Location = New System.Drawing.Point(277, 501)
        Me.cmdCompaniesRemove.Name = "cmdCompaniesRemove"
        Me.cmdCompaniesRemove.Size = New System.Drawing.Size(31, 34)
        Me.cmdCompaniesRemove.TabIndex = 28
        Me.cmdCompaniesRemove.Text = "-"
        Me.cmdCompaniesRemove.UseVisualStyleBackColor = True
        '
        'cmdCompaniesAdd
        '
        Me.cmdCompaniesAdd.Location = New System.Drawing.Point(277, 406)
        Me.cmdCompaniesAdd.Name = "cmdCompaniesAdd"
        Me.cmdCompaniesAdd.Size = New System.Drawing.Size(31, 34)
        Me.cmdCompaniesAdd.TabIndex = 27
        Me.cmdCompaniesAdd.Text = "+"
        Me.cmdCompaniesAdd.UseVisualStyleBackColor = True
        '
        'cmdLancesRemove
        '
        Me.cmdLancesRemove.Location = New System.Drawing.Point(278, 677)
        Me.cmdLancesRemove.Name = "cmdLancesRemove"
        Me.cmdLancesRemove.Size = New System.Drawing.Size(31, 34)
        Me.cmdLancesRemove.TabIndex = 30
        Me.cmdLancesRemove.Text = "-"
        Me.cmdLancesRemove.UseVisualStyleBackColor = True
        '
        'cmdLancesAdd
        '
        Me.cmdLancesAdd.Location = New System.Drawing.Point(278, 582)
        Me.cmdLancesAdd.Name = "cmdLancesAdd"
        Me.cmdLancesAdd.Size = New System.Drawing.Size(31, 34)
        Me.cmdLancesAdd.TabIndex = 29
        Me.cmdLancesAdd.Text = "+"
        Me.cmdLancesAdd.UseVisualStyleBackColor = True
        '
        'cmdSquadsRemove
        '
        Me.cmdSquadsRemove.Location = New System.Drawing.Point(277, 978)
        Me.cmdSquadsRemove.Name = "cmdSquadsRemove"
        Me.cmdSquadsRemove.Size = New System.Drawing.Size(31, 34)
        Me.cmdSquadsRemove.TabIndex = 32
        Me.cmdSquadsRemove.Text = "-"
        Me.cmdSquadsRemove.UseVisualStyleBackColor = True
        '
        'cmdSquadsAdd
        '
        Me.cmdSquadsAdd.Location = New System.Drawing.Point(279, 758)
        Me.cmdSquadsAdd.Name = "cmdSquadsAdd"
        Me.cmdSquadsAdd.Size = New System.Drawing.Size(31, 34)
        Me.cmdSquadsAdd.TabIndex = 31
        Me.cmdSquadsAdd.Text = "+"
        Me.cmdSquadsAdd.UseVisualStyleBackColor = True
        '
        'labRegiment
        '
        Me.labRegiment.AutoSize = True
        Me.labRegiment.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.labRegiment.Location = New System.Drawing.Point(314, 51)
        Me.labRegiment.Name = "labRegiment"
        Me.labRegiment.Size = New System.Drawing.Size(19, 21)
        Me.labRegiment.TabIndex = 33
        Me.labRegiment.Text = "..."
        '
        'labBattalion
        '
        Me.labBattalion.AutoSize = True
        Me.labBattalion.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.labBattalion.Location = New System.Drawing.Point(314, 230)
        Me.labBattalion.Name = "labBattalion"
        Me.labBattalion.Size = New System.Drawing.Size(19, 21)
        Me.labBattalion.TabIndex = 34
        Me.labBattalion.Text = "..."
        '
        'labCompany
        '
        Me.labCompany.AutoSize = True
        Me.labCompany.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.labCompany.Location = New System.Drawing.Point(314, 411)
        Me.labCompany.Name = "labCompany"
        Me.labCompany.Size = New System.Drawing.Size(19, 21)
        Me.labCompany.TabIndex = 35
        Me.labCompany.Text = "..."
        '
        'labLance
        '
        Me.labLance.AutoSize = True
        Me.labLance.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.labLance.Location = New System.Drawing.Point(314, 587)
        Me.labLance.Name = "labLance"
        Me.labLance.Size = New System.Drawing.Size(19, 21)
        Me.labLance.TabIndex = 36
        Me.labLance.Text = "..."
        '
        'labSquadTemplate
        '
        Me.labSquadTemplate.AutoSize = True
        Me.labSquadTemplate.Location = New System.Drawing.Point(879, 705)
        Me.labSquadTemplate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.labSquadTemplate.Name = "labSquadTemplate"
        Me.labSquadTemplate.Size = New System.Drawing.Size(24, 25)
        Me.labSquadTemplate.TabIndex = 37
        Me.labSquadTemplate.Text = "..."
        '
        'cbPreExisting
        '
        Me.cbPreExisting.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbPreExisting.FormattingEnabled = True
        Me.cbPreExisting.Location = New System.Drawing.Point(314, 16)
        Me.cbPreExisting.Name = "cbPreExisting"
        Me.cbPreExisting.Size = New System.Drawing.Size(283, 33)
        Me.cbPreExisting.TabIndex = 38
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(603, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(198, 25)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Load a pre-existing unit"
        '
        'frmCreate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1578, 1138)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cbPreExisting)
        Me.Controls.Add(Me.labSquadTemplate)
        Me.Controls.Add(Me.labLance)
        Me.Controls.Add(Me.labCompany)
        Me.Controls.Add(Me.labBattalion)
        Me.Controls.Add(Me.labRegiment)
        Me.Controls.Add(Me.cmdSquadsRemove)
        Me.Controls.Add(Me.cmdSquadsAdd)
        Me.Controls.Add(Me.cmdLancesRemove)
        Me.Controls.Add(Me.cmdLancesAdd)
        Me.Controls.Add(Me.cmdCompaniesRemove)
        Me.Controls.Add(Me.cmdCompaniesAdd)
        Me.Controls.Add(Me.cmdBattalionsRemove)
        Me.Controls.Add(Me.cmdBattalionsAdd)
        Me.Controls.Add(Me.cmdRegimentsRemove)
        Me.Controls.Add(Me.cmdRegimentsAdd)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lstSquads)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lstLances)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lstCompanies)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lstBattalions)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.nudSupplies)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkHireling)
        Me.Controls.Add(Me.labSquad)
        Me.Controls.Add(Me.labData)
        Me.Controls.Add(Me.cmdFinished)
        Me.Controls.Add(Me.lstRegiments)
        Me.Controls.Add(Me.lstQuality)
        Me.Controls.Add(Me.lstExperience)
        Me.Controls.Add(Me.lstBase)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmCreate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create New Mercenary Unit"
        CType(Me.nudSupplies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstBase As ListBox
    Friend WithEvents lstExperience As ListBox
    Friend WithEvents lstQuality As ListBox
    Friend WithEvents lstRegiments As ListBox
    Friend WithEvents cmdFinished As Button
    Friend WithEvents labData As Label
    Friend WithEvents labSelected As Label
    Friend WithEvents chkHireling As CheckBox
    Friend WithEvents Label1 As Label
    Friend WithEvents nudSupplies As NumericUpDown
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents lstBattalions As ListBox
    Friend WithEvents Label4 As Label
    Friend WithEvents lstCompanies As ListBox
    Friend WithEvents Label5 As Label
    Friend WithEvents lstLances As ListBox
    Friend WithEvents Label6 As Label
    Friend WithEvents lstSquads As ListBox
    Friend WithEvents cmdRegimentsAdd As Button
    Friend WithEvents cmdRegimentsRemove As Button
    Friend WithEvents cmdBattalionsRemove As Button
    Friend WithEvents cmdBattalionsAdd As Button
    Friend WithEvents cmdCompaniesRemove As Button
    Friend WithEvents cmdCompaniesAdd As Button
    Friend WithEvents cmdLancesRemove As Button
    Friend WithEvents cmdLancesAdd As Button
    Friend WithEvents cmdSquadsRemove As Button
    Friend WithEvents cmdSquadsAdd As Button
    Friend WithEvents labRegiment As Label
    Friend WithEvents labBattalion As Label
    Friend WithEvents labCompany As Label
    Friend WithEvents labLance As Label
    Friend WithEvents labSquad As Label
    Friend WithEvents labSquadTemplate As Label
    Friend WithEvents cbPreExisting As ComboBox
    Friend WithEvents Label7 As Label
End Class
