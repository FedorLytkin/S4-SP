﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
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
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.СПеречнемДеталейToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СПеречнемМатериаловToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.ЗаменятьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripComboBox1 = New System.Windows.Forms.ToolStripComboBox()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.Справка = New System.Windows.Forms.ToolStripButton()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ОткрытьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ОчиститьДеревоToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьВетвьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripSeparator1, Me.ToolStripButton2, Me.ToolStripSeparator2, Me.ToolStripDropDownButton1, Me.ToolStripSeparator3, Me.Справка, Me.ToolStripButton4})
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
        Me.ToolStripButton2.ToolTipText = "Выгрузить объект"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.БезРазделаДокументацияToolStripMenuItem, Me.БезТехнолическихСвязейToolStripMenuItem, Me.ToolStripSeparator5, Me.СПеречнемДеталейToolStripMenuItem, Me.СПеречнемМатериаловToolStripMenuItem, Me.ToolStripSeparator6, Me.ЗаменятьToolStripMenuItem, Me.ToolStripComboBox1})
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
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(218, 6)
        '
        'СПеречнемДеталейToolStripMenuItem
        '
        Me.СПеречнемДеталейToolStripMenuItem.Name = "СПеречнемДеталейToolStripMenuItem"
        Me.СПеречнемДеталейToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.СПеречнемДеталейToolStripMenuItem.Text = "С Перечнем деталей"
        '
        'СПеречнемМатериаловToolStripMenuItem
        '
        Me.СПеречнемМатериаловToolStripMenuItem.Name = "СПеречнемМатериаловToolStripMenuItem"
        Me.СПеречнемМатериаловToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.СПеречнемМатериаловToolStripMenuItem.Text = "С Перечнем материалов"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(218, 6)
        '
        'ЗаменятьToolStripMenuItem
        '
        Me.ЗаменятьToolStripMenuItem.Checked = True
        Me.ЗаменятьToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ЗаменятьToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources._space_bar_90666
        Me.ЗаменятьToolStripMenuItem.Name = "ЗаменятьToolStripMenuItem"
        Me.ЗаменятьToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ЗаменятьToolStripMenuItem.Text = "Заменять '~' на ' '"
        Me.ЗаменятьToolStripMenuItem.ToolTipText = "Символ ""~""(тильда) в материале, в разделах Материал, Стандарные и " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Прочие издели" &
    "я будет заменен на "" ""(пробел)"
        '
        'ToolStripComboBox1
        '
        Me.ToolStripComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ToolStripComboBox1.Items.AddRange(New Object() {"Технологический", "Конструкторский", "Комбинированный"})
        Me.ToolStripComboBox1.Name = "ToolStripComboBox1"
        Me.ToolStripComboBox1.Size = New System.Drawing.Size(121, 23)
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'Справка
        '
        Me.Справка.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.Справка.Enabled = False
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
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОткрытьToolStripMenuItem, Me.ToolStripSeparator4, Me.ОчиститьДеревоToolStripMenuItem, Me.УдалитьВетвьToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(219, 76)
        '
        'ОткрытьToolStripMenuItem
        '
        Me.ОткрытьToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Album
        Me.ОткрытьToolStripMenuItem.Name = "ОткрытьToolStripMenuItem"
        Me.ОткрытьToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F4
        Me.ОткрытьToolStripMenuItem.Size = New System.Drawing.Size(218, 22)
        Me.ОткрытьToolStripMenuItem.Text = "Карточка Объекта"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(215, 6)
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
        'ToolStripButton4
        '
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton4.Text = "ToolStripButton4"
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
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ЗаменятьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripComboBox1 As ToolStripComboBox
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents СПеречнемДеталейToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СПеречнемМатериаловToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents ToolStripButton4 As ToolStripButton
End Class