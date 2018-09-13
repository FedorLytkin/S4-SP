<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Checking_SOSTAV_Form
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Checking_SOSTAV_Form))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSplitButton2 = New System.Windows.Forms.ToolStripSplitButton()
        Me.ОткрытьКарточкуДокументаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.РедактированиеКарточкиДокументаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripSplitButton1 = New System.Windows.Forms.ToolStripSplitButton()
        Me.СправкаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ПоказатьВS4ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ОткрытьКарточкуToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПросмотрКарточкиДокументаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.РедактированиеКарточкиДокументаToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.УдалитьЗаписчьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОчиститьПолеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripSeparator1, Me.ToolStripButton2, Me.ToolStripSplitButton2, Me.ToolStripSeparator4, Me.ToolStripButton3})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(389, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        Me.ToolStripButton1.ToolTipText = "Выбор сборочной единицы из S4"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "ToolStripButton2"
        Me.ToolStripButton2.TextDirection = System.Windows.Forms.ToolStripTextDirection.Vertical270
        Me.ToolStripButton2.ToolTipText = "Открыть карточку Объекта"
        '
        'ToolStripSplitButton2
        '
        Me.ToolStripSplitButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripSplitButton2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОткрытьКарточкуДокументаToolStripMenuItem, Me.РедактированиеКарточкиДокументаToolStripMenuItem})
        Me.ToolStripSplitButton2.Image = Global.S4_SP.My.Resources.Resources.newKTD
        Me.ToolStripSplitButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripSplitButton2.Name = "ToolStripSplitButton2"
        Me.ToolStripSplitButton2.Size = New System.Drawing.Size(32, 22)
        Me.ToolStripSplitButton2.Text = "ToolStripSplitButton2"
        Me.ToolStripSplitButton2.ToolTipText = "Открыть карточку документ/объекта"
        '
        'ОткрытьКарточкуДокументаToolStripMenuItem
        '
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.newKTD
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Name = "ОткрытьКарточкуДокументаToolStripMenuItem"
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F3), System.Windows.Forms.Keys)
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Size = New System.Drawing.Size(329, 22)
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Text = "Просмотр карточки документа"
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.ToolTipText = "открыть карточку документ/объекта" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Откроется карта, БЕЗ взятия на редактирование " &
    "документа" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Параметры нельзя будет изменить"
        '
        'РедактированиеКарточкиДокументаToolStripMenuItem
        '
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.newZak
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Name = "РедактированиеКарточкиДокументаToolStripMenuItem"
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Size = New System.Drawing.Size(329, 22)
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Text = "Редактирование карточки документа"
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.ToolTipText = "Открыть карточку документ/объекта" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Откроется карточке, с правим изменение данных"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton3.Image = Global.S4_SP.My.Resources.Resources.S4
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton3.Text = "ToolStripButton3"
        Me.ToolStripButton3.ToolTipText = "Найти объект в S4"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSplitButton1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 297)
        Me.StatusStrip1.MinimumSize = New System.Drawing.Size(0, 30)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(389, 30)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripSplitButton1
        '
        Me.ToolStripSplitButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripSplitButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СправкаToolStripMenuItem})
        Me.ToolStripSplitButton1.Image = Global.S4_SP.My.Resources.Resources.Help1
        Me.ToolStripSplitButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripSplitButton1.Name = "ToolStripSplitButton1"
        Me.ToolStripSplitButton1.Size = New System.Drawing.Size(32, 28)
        Me.ToolStripSplitButton1.Text = "ToolStripSplitButton1"
        '
        'СправкаToolStripMenuItem
        '
        Me.СправкаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Help1
        Me.СправкаToolStripMenuItem.Name = "СправкаToolStripMenuItem"
        Me.СправкаToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.СправкаToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.СправкаToolStripMenuItem.Text = "Справка"
        Me.СправкаToolStripMenuItem.ToolTipText = "Для вызова справки нажмите F1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(292, 301)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Выход"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(0, 25)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(389, 272)
        Me.ListBox1.TabIndex = 3
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПоказатьВS4ToolStripMenuItem, Me.ToolStripSeparator5, Me.ОткрытьКарточкуToolStripMenuItem, Me.ToolStripMenuItem2, Me.ToolStripSeparator2, Me.УдалитьЗаписчьToolStripMenuItem, Me.ОчиститьПолеToolStripMenuItem, Me.ToolStripSeparator3, Me.ToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(217, 154)
        '
        'ПоказатьВS4ToolStripMenuItem
        '
        Me.ПоказатьВS4ToolStripMenuItem.Enabled = False
        Me.ПоказатьВS4ToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.S4
        Me.ПоказатьВS4ToolStripMenuItem.Name = "ПоказатьВS4ToolStripMenuItem"
        Me.ПоказатьВS4ToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.ПоказатьВS4ToolStripMenuItem.Size = New System.Drawing.Size(216, 22)
        Me.ПоказатьВS4ToolStripMenuItem.Text = "&Показать в S4"
        Me.ПоказатьВS4ToolStripMenuItem.ToolTipText = "Найти объект в S4"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(213, 6)
        '
        'ОткрытьКарточкуToolStripMenuItem
        '
        Me.ОткрытьКарточкуToolStripMenuItem.Enabled = False
        Me.ОткрытьКарточкуToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Album
        Me.ОткрытьКарточкуToolStripMenuItem.Name = "ОткрытьКарточкуToolStripMenuItem"
        Me.ОткрытьКарточкуToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F4
        Me.ОткрытьКарточкуToolStripMenuItem.Size = New System.Drawing.Size(216, 22)
        Me.ОткрытьКарточкуToolStripMenuItem.Text = "&Открыть карточку"
        Me.ОткрытьКарточкуToolStripMenuItem.ToolTipText = "Открыть карточку Объекта"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПросмотрКарточкиДокументаToolStripMenuItem, Me.РедактированиеКарточкиДокументаToolStripMenuItem1})
        Me.ToolStripMenuItem2.Image = Global.S4_SP.My.Resources.Resources.newKTD
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(216, 22)
        Me.ToolStripMenuItem2.Text = "Открыть карточку..."
        '
        'ПросмотрКарточкиДокументаToolStripMenuItem
        '
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.newKTD
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Name = "ПросмотрКарточкиДокументаToolStripMenuItem"
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Size = New System.Drawing.Size(278, 22)
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Text = "Просмотр карточки документа"
        '
        'РедактированиеКарточкиДокументаToolStripMenuItem1
        '
        Me.РедактированиеКарточкиДокументаToolStripMenuItem1.Image = Global.S4_SP.My.Resources.Resources.newZak
        Me.РедактированиеКарточкиДокументаToolStripMenuItem1.Name = "РедактированиеКарточкиДокументаToolStripMenuItem1"
        Me.РедактированиеКарточкиДокументаToolStripMenuItem1.Size = New System.Drawing.Size(278, 22)
        Me.РедактированиеКарточкиДокументаToolStripMenuItem1.Text = "Редактирование карточки документа"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(213, 6)
        '
        'УдалитьЗаписчьToolStripMenuItem
        '
        Me.УдалитьЗаписчьToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.clear_icon_9193
        Me.УдалитьЗаписчьToolStripMenuItem.Name = "УдалитьЗаписчьToolStripMenuItem"
        Me.УдалитьЗаписчьToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.Delete
        Me.УдалитьЗаписчьToolStripMenuItem.Size = New System.Drawing.Size(216, 22)
        Me.УдалитьЗаписчьToolStripMenuItem.Text = "&Удалить запись"
        Me.УдалитьЗаписчьToolStripMenuItem.ToolTipText = "Удалить выделенныую строку"
        '
        'ОчиститьПолеToolStripMenuItem
        '
        Me.ОчиститьПолеToolStripMenuItem.Name = "ОчиститьПолеToolStripMenuItem"
        Me.ОчиститьПолеToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Delete), System.Windows.Forms.Keys)
        Me.ОчиститьПолеToolStripMenuItem.Size = New System.Drawing.Size(216, 22)
        Me.ОчиститьПолеToolStripMenuItem.Text = "О&чистить поле"
        Me.ОчиститьПолеToolStripMenuItem.ToolTipText = "Удалить все поля в списке"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(213, 6)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Image = Global.S4_SP.My.Resources.Resources._ASSEMBLY
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(216, 22)
        Me.ToolStripMenuItem1.Text = "П&роверить состав"
        Me.ToolStripMenuItem1.ToolTipText = "Выбор сборочной единицы из S4"
        '
        'Checking_SOSTAV_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(389, 327)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Checking_SOSTAV_Form"
        Me.Text = "Проверка Дерева состава"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripSplitButton1 As ToolStripSplitButton
    Friend WithEvents СправкаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Button1 As Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ПоказатьВS4ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОткрытьКарточкуToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents УдалитьЗаписчьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОчиститьПолеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripButton3 As ToolStripButton
    Friend WithEvents ToolStripSplitButton2 As ToolStripSplitButton
    Friend WithEvents ОткрытьКарточкуДокументаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents РедактированиеКарточкиДокументаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents ToolStripMenuItem2 As ToolStripMenuItem
    Friend WithEvents ПросмотрКарточкиДокументаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents РедактированиеКарточкиДокументаToolStripMenuItem1 As ToolStripMenuItem
End Class
