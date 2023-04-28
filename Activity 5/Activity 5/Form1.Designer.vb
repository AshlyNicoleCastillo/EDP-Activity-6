<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.DataGridView = New System.Windows.Forms.TextBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.BtnUpload = New System.Windows.Forms.Button()
        Me.BtnNewFile = New System.Windows.Forms.Button()
        Me.BtnDelete = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnBackup = New System.Windows.Forms.Button()
        Me.BtnView = New System.Windows.Forms.Button()
        Me.BtnPrint = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView
        '
        Me.DataGridView.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.DataGridView.Font = New System.Drawing.Font("Segoe UI", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.DataGridView.ForeColor = System.Drawing.SystemColors.Desktop
        Me.DataGridView.Location = New System.Drawing.Point(25, 25)
        Me.DataGridView.Name = "DataGridView"
        Me.DataGridView.Size = New System.Drawing.Size(529, 43)
        Me.DataGridView.TabIndex = 1
        Me.DataGridView.Text = "Data Grid View"
        Me.DataGridView.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'DataGridView1
        '
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Window
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(25, 74)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 25
        Me.DataGridView1.Size = New System.Drawing.Size(529, 346)
        Me.DataGridView1.TabIndex = 2
        '
        'BtnUpload
        '
        Me.BtnUpload.BackColor = System.Drawing.Color.DodgerBlue
        Me.BtnUpload.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.BtnUpload.ForeColor = System.Drawing.SystemColors.Desktop
        Me.BtnUpload.Location = New System.Drawing.Point(607, 88)
        Me.BtnUpload.Name = "BtnUpload"
        Me.BtnUpload.Size = New System.Drawing.Size(140, 38)
        Me.BtnUpload.TabIndex = 3
        Me.BtnUpload.Text = "UPLOAD"
        Me.BtnUpload.UseVisualStyleBackColor = False
        '
        'BtnNewFile
        '
        Me.BtnNewFile.BackColor = System.Drawing.Color.PaleGreen
        Me.BtnNewFile.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.BtnNewFile.ForeColor = System.Drawing.SystemColors.Desktop
        Me.BtnNewFile.Location = New System.Drawing.Point(609, 263)
        Me.BtnNewFile.Name = "BtnNewFile"
        Me.BtnNewFile.Size = New System.Drawing.Size(140, 38)
        Me.BtnNewFile.TabIndex = 5
        Me.BtnNewFile.Text = "NEW FILE"
        Me.BtnNewFile.UseVisualStyleBackColor = False
        '
        'BtnDelete
        '
        Me.BtnDelete.BackColor = System.Drawing.Color.LightCoral
        Me.BtnDelete.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.BtnDelete.ForeColor = System.Drawing.SystemColors.Desktop
        Me.BtnDelete.Location = New System.Drawing.Point(609, 322)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(140, 38)
        Me.BtnDelete.TabIndex = 6
        Me.BtnDelete.Text = "DELETE"
        Me.BtnDelete.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label1.ForeColor = System.Drawing.SystemColors.Window
        Me.Label1.Location = New System.Drawing.Point(607, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(142, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "By Ashly Nicole Castillo"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Font = New System.Drawing.Font("Lucida Sans Unicode", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label2.ForeColor = System.Drawing.SystemColors.Window
        Me.Label2.Location = New System.Drawing.Point(576, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(201, 34)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Data Loading"
        '
        'BtnBackup
        '
        Me.BtnBackup.BackColor = System.Drawing.Color.SandyBrown
        Me.BtnBackup.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.BtnBackup.ForeColor = System.Drawing.SystemColors.Desktop
        Me.BtnBackup.Location = New System.Drawing.Point(607, 147)
        Me.BtnBackup.Name = "BtnBackup"
        Me.BtnBackup.Size = New System.Drawing.Size(140, 38)
        Me.BtnBackup.TabIndex = 11
        Me.BtnBackup.Text = "BACKUP"
        Me.BtnBackup.UseVisualStyleBackColor = False
        '
        'BtnView
        '
        Me.BtnView.BackColor = System.Drawing.Color.Thistle
        Me.BtnView.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.BtnView.ForeColor = System.Drawing.SystemColors.Desktop
        Me.BtnView.Location = New System.Drawing.Point(609, 205)
        Me.BtnView.Name = "BtnView"
        Me.BtnView.Size = New System.Drawing.Size(140, 38)
        Me.BtnView.TabIndex = 12
        Me.BtnView.Text = "VIEW"
        Me.BtnView.UseVisualStyleBackColor = False
        '
        'BtnPrint
        '
        Me.BtnPrint.BackColor = System.Drawing.Color.MediumTurquoise
        Me.BtnPrint.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.BtnPrint.ForeColor = System.Drawing.SystemColors.Desktop
        Me.BtnPrint.Location = New System.Drawing.Point(607, 382)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(140, 38)
        Me.BtnPrint.TabIndex = 13
        Me.BtnPrint.Text = "PRINT"
        Me.BtnPrint.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.BtnPrint)
        Me.Controls.Add(Me.BtnView)
        Me.Controls.Add(Me.BtnBackup)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnDelete)
        Me.Controls.Add(Me.BtnNewFile)
        Me.Controls.Add(Me.BtnUpload)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.DataGridView)
        Me.Name = "Form1"
        Me.Text = "Data Loading , Conversion and Database Backup"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView As TextBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents BtnUpload As Button
    Friend WithEvents BtnNewFile As Button
    Friend WithEvents BtnDelete As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents BtnBackup As Button
    Friend WithEvents BtnView As Button
    Friend WithEvents BtnPrint As Button
End Class
