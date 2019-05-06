<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Add_newReplaceChar
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Add_newReplaceChar))
        Me.Find_TB = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Replace_TB = New System.Windows.Forms.TextBox()
        Me.OK_Bt = New System.Windows.Forms.Button()
        Me.Cancel_Bt = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Find_TB
        '
        Me.Find_TB.Location = New System.Drawing.Point(86, 12)
        Me.Find_TB.Name = "Find_TB"
        Me.Find_TB.Size = New System.Drawing.Size(100, 20)
        Me.Find_TB.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Найти"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Заменить"
        '
        'Replace_TB
        '
        Me.Replace_TB.Location = New System.Drawing.Point(86, 38)
        Me.Replace_TB.Name = "Replace_TB"
        Me.Replace_TB.Size = New System.Drawing.Size(100, 20)
        Me.Replace_TB.TabIndex = 2
        '
        'OK_Bt
        '
        Me.OK_Bt.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK_Bt.Location = New System.Drawing.Point(6, 67)
        Me.OK_Bt.Name = "OK_Bt"
        Me.OK_Bt.Size = New System.Drawing.Size(75, 23)
        Me.OK_Bt.TabIndex = 4
        Me.OK_Bt.Text = "Button1"
        Me.OK_Bt.UseVisualStyleBackColor = True
        '
        'Cancel_Bt
        '
        Me.Cancel_Bt.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Bt.Location = New System.Drawing.Point(111, 67)
        Me.Cancel_Bt.Name = "Cancel_Bt"
        Me.Cancel_Bt.Size = New System.Drawing.Size(75, 23)
        Me.Cancel_Bt.TabIndex = 5
        Me.Cancel_Bt.Text = "Отмена"
        Me.Cancel_Bt.UseVisualStyleBackColor = True
        '
        'Add_newReplaceChar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(191, 93)
        Me.Controls.Add(Me.Cancel_Bt)
        Me.Controls.Add(Me.OK_Bt)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Replace_TB)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Find_TB)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(207, 132)
        Me.MinimumSize = New System.Drawing.Size(207, 132)
        Me.Name = "Add_newReplaceChar"
        Me.Text = "Add_newReplaceChar"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Find_TB As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Replace_TB As TextBox
    Friend WithEvents OK_Bt As Button
    Friend WithEvents Cancel_Bt As Button
End Class
