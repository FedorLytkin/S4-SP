<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CharOption
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CharOption))
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ДобавитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИзменитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(1, 1)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(260, 186)
        Me.ListBox1.TabIndex = 0
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ДобавитьToolStripMenuItem, Me.ИзменитьToolStripMenuItem, Me.ToolStripSeparator1, Me.УдалитьToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(170, 76)
        '
        'ДобавитьToolStripMenuItem
        '
        Me.ДобавитьToolStripMenuItem.Name = "ДобавитьToolStripMenuItem"
        Me.ДобавитьToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.ДобавитьToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.ДобавитьToolStripMenuItem.Text = "Добавить"
        '
        'УдалитьToolStripMenuItem
        '
        Me.УдалитьToolStripMenuItem.Name = "УдалитьToolStripMenuItem"
        Me.УдалитьToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.Delete
        Me.УдалитьToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.УдалитьToolStripMenuItem.Text = "Удалить"
        '
        'ИзменитьToolStripMenuItem
        '
        Me.ИзменитьToolStripMenuItem.Name = "ИзменитьToolStripMenuItem"
        Me.ИзменитьToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.ИзменитьToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.ИзменитьToolStripMenuItem.Text = "Изменить"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(88, 193)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Сохранить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(186, 193)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Отмена"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(166, 6)
        '
        'CharOption
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(262, 218)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ListBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(278, 257)
        Me.MinimumSize = New System.Drawing.Size(278, 257)
        Me.Name = "CharOption"
        Me.Text = "Найти Заменить"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ДобавитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УдалитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИзменитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
End Class
