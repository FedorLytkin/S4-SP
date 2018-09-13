<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class VSNRM2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(VSNRM2))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.БезРазделаДокументацияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.БезТехнолическихСвязейToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.Справка = New System.Windows.Forms.ToolStripButton()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ОткрытьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОчиститьДеревоToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьВетвьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripSeparator1, Me.ToolStripButton2, Me.ToolStripSeparator2, Me.ToolStripDropDownButton1, Me.ToolStripSeparator3, Me.Справка})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(427, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = Global.S4_SP.My.Resources.Resources.ASSEMBLY1
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        Me.ToolStripButton1.ToolTipText = "Выбрать объект"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = Global.S4_SP.My.Resources.Resources.excel_png_office_xlsx_icon_3
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "ToolStripButton2"
        Me.ToolStripButton2.ToolTipText = "Выгрузать объект"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.БезРазделаДокументацияToolStripMenuItem, Me.БезТехнолическихСвязейToolStripMenuItem})
        Me.ToolStripDropDownButton1.Image = Global.S4_SP.My.Resources.Resources.wrench_512
        Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(29, 22)
        Me.ToolStripDropDownButton1.Text = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.ToolTipText = "Настройки"
        '
        'БезРазделаДокументацияToolStripMenuItem
        '
        Me.БезРазделаДокументацияToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.DOCUMENT
        Me.БезРазделаДокументацияToolStripMenuItem.Name = "БезРазделаДокументацияToolStripMenuItem"
        Me.БезРазделаДокументацияToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.БезРазделаДокументацияToolStripMenuItem.Text = "С разделом Документация"
        Me.БезРазделаДокументацияToolStripMenuItem.ToolTipText = "В состав изделия будут добавляться объекты из раздела Документация"
        '
        'БезТехнолическихСвязейToolStripMenuItem
        '
        Me.БезТехнолическихСвязейToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.ручная_связь
        Me.БезТехнолическихСвязейToolStripMenuItem.Name = "БезТехнолическихСвязейToolStripMenuItem"
        Me.БезТехнолическихСвязейToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.БезТехнолическихСвязейToolStripMenuItem.Text = "С Ручными связями"
        Me.БезТехнолическихСвязейToolStripMenuItem.ToolTipText = "В состав изделия будут добавляться объекты прикрепленные 'ручной' связью"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'Справка
        '
        Me.Справка.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.Справка.Image = Global.S4_SP.My.Resources.Resources.evernote
        Me.Справка.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Справка.Name = "Справка"
        Me.Справка.Size = New System.Drawing.Size(23, 22)
        Me.Справка.Text = "ToolStripButton5"
        Me.Справка.ToolTipText = "Справка"
        '
        'TreeView1
        '
        Me.TreeView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.Location = New System.Drawing.Point(0, 25)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(427, 292)
        Me.TreeView1.TabIndex = 1
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОткрытьToolStripMenuItem, Me.ОчиститьДеревоToolStripMenuItem, Me.УдалитьВетвьToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(219, 70)
        '
        'ОткрытьToolStripMenuItem
        '
        Me.ОткрытьToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Album
        Me.ОткрытьToolStripMenuItem.Name = "ОткрытьToolStripMenuItem"
        Me.ОткрытьToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F4
        Me.ОткрытьToolStripMenuItem.Size = New System.Drawing.Size(218, 22)
        Me.ОткрытьToolStripMenuItem.Text = "Карточка Объекта"
        '
        'ОчиститьДеревоToolStripMenuItem
        '
        Me.ОчиститьДеревоToolStripMenuItem.Name = "ОчиститьДеревоToolStripMenuItem"
        Me.ОчиститьДеревоToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Delete), System.Windows.Forms.Keys)
        Me.ОчиститьДеревоToolStripMenuItem.Size = New System.Drawing.Size(218, 22)
        Me.ОчиститьДеревоToolStripMenuItem.Text = "Очистить дерево"
        '
        'УдалитьВетвьToolStripMenuItem
        '
        Me.УдалитьВетвьToolStripMenuItem.Name = "УдалитьВетвьToolStripMenuItem"
        Me.УдалитьВетвьToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.Delete
        Me.УдалитьВетвьToolStripMenuItem.Size = New System.Drawing.Size(218, 22)
        Me.УдалитьВетвьToolStripMenuItem.Text = "Удалить ветвь"
        '
        'VSNRM2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(427, 317)
        Me.Controls.Add(Me.TreeView1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "VSNRM2"
        Me.Text = "ВСНРМ 2.0"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents TreeView1 As TreeView
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ToolStripDropDownButton1 As ToolStripDropDownButton
    Friend WithEvents БезРазделаДокументацияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents БезТехнолическихСвязейToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents Справка As ToolStripButton
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ОткрытьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОчиститьДеревоToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УдалитьВетвьToolStripMenuItem As ToolStripMenuItem
End Class
