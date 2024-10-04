<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Cal
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
        TableLayoutPanel1 = New TableLayoutPanel()
        Label_Title = New Label()
        Label_Status = New Label()
        TableLayoutPanel2 = New TableLayoutPanel()
        Button_YES = New Button()
        Button_Cancel = New Button()
        TableLayoutPanel1.SuspendLayout()
        TableLayoutPanel2.SuspendLayout()
        SuspendLayout()
        ' 
        ' TableLayoutPanel1
        ' 
        TableLayoutPanel1.ColumnCount = 1
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TableLayoutPanel1.Controls.Add(Label_Title, 0, 0)
        TableLayoutPanel1.Controls.Add(Label_Status, 0, 1)
        TableLayoutPanel1.Controls.Add(TableLayoutPanel2, 0, 2)
        TableLayoutPanel1.Dock = DockStyle.Fill
        TableLayoutPanel1.Location = New Point(0, 0)
        TableLayoutPanel1.Margin = New Padding(0)
        TableLayoutPanel1.Name = "TableLayoutPanel1"
        TableLayoutPanel1.RowCount = 4
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 50F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 80F))
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Absolute, 20F))
        TableLayoutPanel1.Size = New Size(800, 450)
        TableLayoutPanel1.TabIndex = 0
        ' 
        ' Label_Title
        ' 
        Label_Title.AutoSize = True
        Label_Title.Dock = DockStyle.Fill
        Label_Title.Font = New Font("Malgun Gothic", 16F, FontStyle.Bold, GraphicsUnit.Point, CByte(129))
        Label_Title.Location = New Point(3, 0)
        Label_Title.Name = "Label_Title"
        Label_Title.Size = New Size(794, 50)
        Label_Title.TabIndex = 0
        Label_Title.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' Label_Status
        ' 
        Label_Status.AutoSize = True
        Label_Status.BackColor = Color.WhiteSmoke
        Label_Status.Dock = DockStyle.Fill
        Label_Status.Font = New Font("Malgun Gothic", 28F, FontStyle.Bold, GraphicsUnit.Point, CByte(129))
        Label_Status.Location = New Point(10, 60)
        Label_Status.Margin = New Padding(10)
        Label_Status.Name = "Label_Status"
        Label_Status.Size = New Size(780, 280)
        Label_Status.TabIndex = 1
        Label_Status.Text = "Ready " & vbCrLf & "for " & vbCrLf & "Calibraition?"
        Label_Status.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' TableLayoutPanel2
        ' 
        TableLayoutPanel2.ColumnCount = 4
        TableLayoutPanel2.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 50F))
        TableLayoutPanel2.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50F))
        TableLayoutPanel2.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50F))
        TableLayoutPanel2.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 50F))
        TableLayoutPanel2.Controls.Add(Button_YES, 1, 0)
        TableLayoutPanel2.Controls.Add(Button_Cancel, 2, 0)
        TableLayoutPanel2.Dock = DockStyle.Fill
        TableLayoutPanel2.Location = New Point(3, 353)
        TableLayoutPanel2.Name = "TableLayoutPanel2"
        TableLayoutPanel2.RowCount = 1
        TableLayoutPanel2.RowStyles.Add(New RowStyle(SizeType.Percent, 100F))
        TableLayoutPanel2.Size = New Size(794, 74)
        TableLayoutPanel2.TabIndex = 2
        ' 
        ' Button_YES
        ' 
        Button_YES.Dock = DockStyle.Fill
        Button_YES.Location = New Point(60, 10)
        Button_YES.Margin = New Padding(10)
        Button_YES.Name = "Button_YES"
        Button_YES.Size = New Size(327, 54)
        Button_YES.TabIndex = 0
        Button_YES.Text = "YES"
        Button_YES.UseVisualStyleBackColor = True
        ' 
        ' Button_Cancel
        ' 
        Button_Cancel.Dock = DockStyle.Fill
        Button_Cancel.Location = New Point(407, 10)
        Button_Cancel.Margin = New Padding(10)
        Button_Cancel.Name = "Button_Cancel"
        Button_Cancel.Size = New Size(327, 54)
        Button_Cancel.TabIndex = 1
        Button_Cancel.Text = "CANCEL"
        Button_Cancel.UseVisualStyleBackColor = True
        ' 
        ' Cal
        ' 
        AutoScaleDimensions = New SizeF(144F, 144F)
        AutoScaleMode = AutoScaleMode.Dpi
        BackColor = Color.White
        ClientSize = New Size(800, 450)
        Controls.Add(TableLayoutPanel1)
        Font = New Font("Malgun Gothic", 14F, FontStyle.Regular, GraphicsUnit.Point, CByte(129))
        Name = "Cal"
        Text = "Cal"
        TableLayoutPanel1.ResumeLayout(False)
        TableLayoutPanel1.PerformLayout()
        TableLayoutPanel2.ResumeLayout(False)
        ResumeLayout(False)
    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label_Title As Label
    Friend WithEvents Label_Status As Label
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Button_YES As Button
    Friend WithEvents Button_Cancel As Button
End Class
