<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.dtpExtract = New System.Windows.Forms.DateTimePicker()
        Me.btnGen = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtCustomer = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtArea = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtBranch = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pbIT = New System.Windows.Forms.PictureBox()
        Me.wbAds = New System.Windows.Forms.WebBrowser()
        Me.tmrAds = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        CType(Me.pbIT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtpExtract
        '
        Me.dtpExtract.Location = New System.Drawing.Point(12, 12)
        Me.dtpExtract.Name = "dtpExtract"
        Me.dtpExtract.Size = New System.Drawing.Size(200, 20)
        Me.dtpExtract.TabIndex = 0
        '
        'btnGen
        '
        Me.btnGen.Location = New System.Drawing.Point(12, 38)
        Me.btnGen.Name = "btnGen"
        Me.btnGen.Size = New System.Drawing.Size(200, 36)
        Me.btnGen.TabIndex = 1
        Me.btnGen.Text = "&Generate"
        Me.btnGen.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtCustomer)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtArea)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtBranch)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(218, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(529, 62)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Details"
        '
        'txtCustomer
        '
        Me.txtCustomer.Location = New System.Drawing.Point(423, 26)
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.Size = New System.Drawing.Size(100, 20)
        Me.txtCustomer.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(360, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Customer :"
        '
        'txtArea
        '
        Me.txtArea.Location = New System.Drawing.Point(228, 26)
        Me.txtArea.Name = "txtArea"
        Me.txtArea.Size = New System.Drawing.Size(100, 20)
        Me.txtArea.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(187, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Area :"
        '
        'txtBranch
        '
        Me.txtBranch.Location = New System.Drawing.Point(58, 26)
        Me.txtBranch.Name = "txtBranch"
        Me.txtBranch.Size = New System.Drawing.Size(100, 20)
        Me.txtBranch.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Branch :"
        '
        'pbIT
        '
        Me.pbIT.Image = Global.PTUGenerator.My.Resources.Resources.itd_728x90
        Me.pbIT.Location = New System.Drawing.Point(12, 80)
        Me.pbIT.Name = "pbIT"
        Me.pbIT.Size = New System.Drawing.Size(735, 98)
        Me.pbIT.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.pbIT.TabIndex = 3
        Me.pbIT.TabStop = False
        '
        'wbAds
        '
        Me.wbAds.Location = New System.Drawing.Point(12, 80)
        Me.wbAds.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbAds.Name = "wbAds"
        Me.wbAds.Size = New System.Drawing.Size(735, 98)
        Me.wbAds.TabIndex = 4
        '
        'tmrAds
        '
        Me.tmrAds.Interval = 1000
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(757, 187)
        Me.Controls.Add(Me.pbIT)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnGen)
        Me.Controls.Add(Me.dtpExtract)
        Me.Controls.Add(Me.wbAds)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Extractor"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.pbIT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dtpExtract As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnGen As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtArea As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtBranch As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents pbIT As System.Windows.Forms.PictureBox
    Friend WithEvents wbAds As System.Windows.Forms.WebBrowser
    Friend WithEvents tmrAds As System.Windows.Forms.Timer

End Class
