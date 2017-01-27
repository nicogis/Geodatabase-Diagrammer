<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormGDBDiagrammer
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
        Me.btnRunGDBDiagrammer = New System.Windows.Forms.Button()
        Me.chkAbstract = New System.Windows.Forms.CheckBox()
        Me.chkFieldMetadataD = New System.Windows.Forms.CheckBox()
        Me.chkFieldAlias = New System.Windows.Forms.CheckBox()
        Me.chkOmitAnno = New System.Windows.Forms.CheckBox()
        Me.txtOutputFile = New System.Windows.Forms.TextBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.btnSaveFileDialog = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optPS = New System.Windows.Forms.RadioButton()
        Me.optTT = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstDomain = New System.Windows.Forms.ListBox()
        Me.lstTemp = New System.Windows.Forms.ListBox()
        Me.lstDataset = New System.Windows.Forms.ListBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.chkSummary = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SaveFileDialog2 = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnRunGDBDiagrammer
        '
        Me.btnRunGDBDiagrammer.Location = New System.Drawing.Point(176, 222)
        Me.btnRunGDBDiagrammer.Name = "btnRunGDBDiagrammer"
        Me.btnRunGDBDiagrammer.Size = New System.Drawing.Size(123, 23)
        Me.btnRunGDBDiagrammer.TabIndex = 0
        Me.btnRunGDBDiagrammer.Text = "Generate Diagram"
        Me.btnRunGDBDiagrammer.UseVisualStyleBackColor = True
        '
        'chkAbstract
        '
        Me.chkAbstract.AutoSize = True
        Me.chkAbstract.Location = New System.Drawing.Point(12, 54)
        Me.chkAbstract.Name = "chkAbstract"
        Me.chkAbstract.Size = New System.Drawing.Size(238, 17)
        Me.chkAbstract.TabIndex = 1
        Me.chkAbstract.Text = "Use Metadata abstract for Class Descriptions"
        Me.chkAbstract.UseVisualStyleBackColor = True
        '
        'chkFieldMetadataD
        '
        Me.chkFieldMetadataD.AutoSize = True
        Me.chkFieldMetadataD.Location = New System.Drawing.Point(12, 77)
        Me.chkFieldMetadataD.Name = "chkFieldMetadataD"
        Me.chkFieldMetadataD.Size = New System.Drawing.Size(264, 17)
        Me.chkFieldMetadataD.TabIndex = 2
        Me.chkFieldMetadataD.Text = "Use the field metadata definition for the description"
        Me.chkFieldMetadataD.UseVisualStyleBackColor = True
        '
        'chkFieldAlias
        '
        Me.chkFieldAlias.AutoSize = True
        Me.chkFieldAlias.Location = New System.Drawing.Point(12, 100)
        Me.chkFieldAlias.Name = "chkFieldAlias"
        Me.chkFieldAlias.Size = New System.Drawing.Size(196, 17)
        Me.chkFieldAlias.TabIndex = 3
        Me.chkFieldAlias.Text = "Use the field alias for the description"
        Me.chkFieldAlias.UseVisualStyleBackColor = True
        '
        'chkOmitAnno
        '
        Me.chkOmitAnno.AutoSize = True
        Me.chkOmitAnno.Location = New System.Drawing.Point(12, 123)
        Me.chkOmitAnno.Name = "chkOmitAnno"
        Me.chkOmitAnno.Size = New System.Drawing.Size(183, 17)
        Me.chkOmitAnno.TabIndex = 4
        Me.chkOmitAnno.Text = "Omit fields for annotation classes."
        Me.chkOmitAnno.UseVisualStyleBackColor = True
        '
        'txtOutputFile
        '
        Me.txtOutputFile.Location = New System.Drawing.Point(12, 27)
        Me.txtOutputFile.Name = "txtOutputFile"
        Me.txtOutputFile.Size = New System.Drawing.Size(213, 20)
        Me.txtOutputFile.TabIndex = 5
        '
        'btnSaveFileDialog
        '
        Me.btnSaveFileDialog.Location = New System.Drawing.Point(231, 26)
        Me.btnSaveFileDialog.Name = "btnSaveFileDialog"
        Me.btnSaveFileDialog.Size = New System.Drawing.Size(68, 23)
        Me.btnSaveFileDialog.TabIndex = 6
        Me.btnSaveFileDialog.Text = "Browse"
        Me.btnSaveFileDialog.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optPS)
        Me.GroupBox1.Controls.Add(Me.optTT)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 146)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(167, 70)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Font Type for Visio Output"
        '
        'optPS
        '
        Me.optPS.AutoSize = True
        Me.optPS.Location = New System.Drawing.Point(21, 42)
        Me.optPS.Name = "optPS"
        Me.optPS.Size = New System.Drawing.Size(71, 17)
        Me.optPS.TabIndex = 1
        Me.optPS.TabStop = True
        Me.optPS.Text = "Postscript"
        Me.optPS.UseVisualStyleBackColor = True
        '
        'optTT
        '
        Me.optTT.AutoSize = True
        Me.optTT.Location = New System.Drawing.Point(21, 19)
        Me.optTT.Name = "optTT"
        Me.optTT.Size = New System.Drawing.Size(74, 17)
        Me.optTT.TabIndex = 0
        Me.optTT.TabStop = True
        Me.optTT.Text = "True Type"
        Me.optTT.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(127, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Create Visio Diagram File:"
        '
        'lstDomain
        '
        Me.lstDomain.FormattingEnabled = True
        Me.lstDomain.Location = New System.Drawing.Point(10, 278)
        Me.lstDomain.Name = "lstDomain"
        Me.lstDomain.Size = New System.Drawing.Size(70, 30)
        Me.lstDomain.TabIndex = 9
        Me.lstDomain.Visible = False
        '
        'lstTemp
        '
        Me.lstTemp.FormattingEnabled = True
        Me.lstTemp.Location = New System.Drawing.Point(198, 278)
        Me.lstTemp.Name = "lstTemp"
        Me.lstTemp.Size = New System.Drawing.Size(89, 30)
        Me.lstTemp.TabIndex = 10
        Me.lstTemp.Visible = False
        '
        'lstDataset
        '
        Me.lstDataset.FormattingEnabled = True
        Me.lstDataset.Location = New System.Drawing.Point(101, 278)
        Me.lstDataset.Name = "lstDataset"
        Me.lstDataset.Size = New System.Drawing.Size(81, 30)
        Me.lstDataset.TabIndex = 11
        Me.lstDataset.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(15, 222)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'chkSummary
        '
        Me.chkSummary.AutoSize = True
        Me.chkSummary.Location = New System.Drawing.Point(12, 325)
        Me.chkSummary.Name = "chkSummary"
        Me.chkSummary.Size = New System.Drawing.Size(81, 17)
        Me.chkSummary.TabIndex = 13
        Me.chkSummary.Text = "CheckBox1"
        Me.chkSummary.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(262, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "v10.5"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 248)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(311, 22)
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'FormGDBDiagrammer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(311, 270)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.chkSummary)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lstDataset)
        Me.Controls.Add(Me.lstTemp)
        Me.Controls.Add(Me.lstDomain)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnSaveFileDialog)
        Me.Controls.Add(Me.txtOutputFile)
        Me.Controls.Add(Me.chkOmitAnno)
        Me.Controls.Add(Me.chkFieldAlias)
        Me.Controls.Add(Me.chkFieldMetadataD)
        Me.Controls.Add(Me.chkAbstract)
        Me.Controls.Add(Me.btnRunGDBDiagrammer)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormGDBDiagrammer"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Geodatabase Diagrammer"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRunGDBDiagrammer As System.Windows.Forms.Button
    Friend WithEvents chkAbstract As System.Windows.Forms.CheckBox
    Friend WithEvents chkFieldMetadataD As System.Windows.Forms.CheckBox
    Friend WithEvents chkFieldAlias As System.Windows.Forms.CheckBox
    Friend WithEvents chkOmitAnno As System.Windows.Forms.CheckBox
    Friend WithEvents txtOutputFile As System.Windows.Forms.TextBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnSaveFileDialog As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optPS As System.Windows.Forms.RadioButton
    Friend WithEvents optTT As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstDomain As System.Windows.Forms.ListBox
    Friend WithEvents lstTemp As System.Windows.Forms.ListBox
    Friend WithEvents lstDataset As System.Windows.Forms.ListBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents chkSummary As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SaveFileDialog2 As System.Windows.Forms.SaveFileDialog
End Class
