<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain2))
        Me.pbIT = New System.Windows.Forms.PictureBox()
        Me.mcSales = New System.Windows.Forms.MonthCalendar()
        Me.btnGen = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtBranch = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtArea = New System.Windows.Forms.TextBox()
        Me.txtCompany = New System.Windows.Forms.TextBox()
        Me.lblSite = New System.Windows.Forms.Label()
        Me.wbAds = New System.Windows.Forms.WebBrowser()
        CType(Me.pbIT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbIT
        '
        Me.pbIT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbIT.Image = Global.PTUGenerator.My.Resources.Resources.itd_728x90
        Me.pbIT.Location = New System.Drawing.Point(2, 176)
        Me.pbIT.Name = "pbIT"
        Me.pbIT.Size = New System.Drawing.Size(734, 95)
        Me.pbIT.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.pbIT.TabIndex = 0
        Me.pbIT.TabStop = False
        '
        'mcSales
        '
        Me.mcSales.Location = New System.Drawing.Point(6, 7)
        Me.mcSales.Name = "mcSales"
        Me.mcSales.TabIndex = 0
        '
        'btnGen
        '
        Me.btnGen.Location = New System.Drawing.Point(237, 16)
        Me.btnGen.Name = "btnGen"
        Me.btnGen.Size = New System.Drawing.Size(75, 66)
        Me.btnGen.TabIndex = 1
        Me.btnGen.Text = "Generate"
        Me.btnGen.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(237, 93)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 66)
        Me.btnExit.TabIndex = 2
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(343, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(155, 25)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "BranchCode :"
        '
        'txtBranch
        '
        Me.txtBranch.BackColor = System.Drawing.Color.White
        Me.txtBranch.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBranch.Location = New System.Drawing.Point(504, 32)
        Me.txtBranch.Name = "txtBranch"
        Me.txtBranch.ReadOnly = True
        Me.txtBranch.Size = New System.Drawing.Size(201, 29)
        Me.txtBranch.TabIndex = 3
        Me.txtBranch.Text = "ROX"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(343, 77)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 25)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Area :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(343, 113)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(124, 25)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Company :"
        '
        'txtArea
        '
        Me.txtArea.BackColor = System.Drawing.Color.White
        Me.txtArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArea.Location = New System.Drawing.Point(504, 73)
        Me.txtArea.Name = "txtArea"
        Me.txtArea.ReadOnly = True
        Me.txtArea.Size = New System.Drawing.Size(201, 29)
        Me.txtArea.TabIndex = 4
        Me.txtArea.Text = "GSC"
        '
        'txtCompany
        '
        Me.txtCompany.BackColor = System.Drawing.Color.White
        Me.txtCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompany.Location = New System.Drawing.Point(504, 109)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.ReadOnly = True
        Me.txtCompany.Size = New System.Drawing.Size(201, 29)
        Me.txtCompany.TabIndex = 5
        Me.txtCompany.Text = "Photo"
        '
        'lblSite
        '
        Me.lblSite.AutoSize = True
        Me.lblSite.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSite.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblSite.Location = New System.Drawing.Point(599, 157)
        Me.lblSite.Name = "lblSite"
        Me.lblSite.Size = New System.Drawing.Size(124, 13)
        Me.lblSite.TabIndex = 12
        Me.lblSite.Text = "http://pgc-itdept.org"
        '
        'wbAds
        '
        Me.wbAds.Location = New System.Drawing.Point(2, 176)
        Me.wbAds.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbAds.Name = "wbAds"
        Me.wbAds.Size = New System.Drawing.Size(734, 95)
        Me.wbAds.TabIndex = 13
        '
        'frmMain2
        '
        Me.AcceptButton = Me.btnGen
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(735, 271)
        Me.Controls.Add(Me.lblSite)
        Me.Controls.Add(Me.txtCompany)
        Me.Controls.Add(Me.txtArea)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtBranch)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnGen)
        Me.Controls.Add(Me.mcSales)
        Me.Controls.Add(Me.pbIT)
        Me.Controls.Add(Me.wbAds)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SAP Sales"
        CType(Me.pbIT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pbIT As System.Windows.Forms.PictureBox
    Friend WithEvents mcSales As System.Windows.Forms.MonthCalendar
    Friend WithEvents btnGen As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtBranch As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtArea As System.Windows.Forms.TextBox
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents lblSite As System.Windows.Forms.Label
    Friend WithEvents wbAds As System.Windows.Forms.WebBrowser
End Class
