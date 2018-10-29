<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ВыбратьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыбратьОбъектToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СоздатьДеревоToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ПроверитьДеревоСоставаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ОткрытьВСНРМToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОтчетыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЖурналПроектовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЖурналОТДToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПравитьОбъектToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВСНРМ2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СервисToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НастройкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПомощьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СправкаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОПрограммеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОткрытьКарточкуДокументаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПросмотрКарточкиДокументаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.РедактированиеКарточкиДокументаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ОткрытьВСНРМToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ОчиститьИсториюToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MenuStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВыбратьToolStripMenuItem, Me.ОтчетыToolStripMenuItem, Me.СервисToolStripMenuItem, Me.ПомощьToolStripMenuItem, Me.ToolStripTextBox1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(612, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ВыбратьToolStripMenuItem
        '
        Me.ВыбратьToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВыбратьОбъектToolStripMenuItem, Me.СоздатьДеревоToolStripMenuItem, Me.ToolStripSeparator1, Me.ПроверитьДеревоСоставаToolStripMenuItem, Me.ToolStripSeparator2, Me.ОткрытьВСНРМToolStripMenuItem})
        Me.ВыбратьToolStripMenuItem.Name = "ВыбратьToolStripMenuItem"
        Me.ВыбратьToolStripMenuItem.Size = New System.Drawing.Size(66, 20)
        Me.ВыбратьToolStripMenuItem.Text = "Выбрать"
        '
        'ВыбратьОбъектToolStripMenuItem
        '
        Me.ВыбратьОбъектToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Открыть_древов_ВСНРМ3
        Me.ВыбратьОбъектToolStripMenuItem.Name = "ВыбратьОбъектToolStripMenuItem"
        Me.ВыбратьОбъектToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.ВыбратьОбъектToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.ВыбратьОбъектToolStripMenuItem.Text = "Открыть древовидное ВСНРМ"
        Me.ВыбратьОбъектToolStripMenuItem.ToolTipText = "Открыть ВСНРМ с древовидной структурой"
        '
        'СоздатьДеревоToolStripMenuItem
        '
        Me.СоздатьДеревоToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Открыть_древов_ВСНРМ2
        Me.СоздатьДеревоToolStripMenuItem.Name = "СоздатьДеревоToolStripMenuItem"
        Me.СоздатьДеревоToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.СоздатьДеревоToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.СоздатьДеревоToolStripMenuItem.Text = "Создать дерево в S4"
        Me.СоздатьДеревоToolStripMenuItem.ToolTipText = "По ВСНРМ, создать в архиве изделие с деревом состава" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(277, 6)
        '
        'ПроверитьДеревоСоставаToolStripMenuItem
        '
        Me.ПроверитьДеревоСоставаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.SP
        Me.ПроверитьДеревоСоставаToolStripMenuItem.Name = "ПроверитьДеревоСоставаToolStripMenuItem"
        Me.ПроверитьДеревоСоставаToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.D2), System.Windows.Forms.Keys)
        Me.ПроверитьДеревоСоставаToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.ПроверитьДеревоСоставаToolStripMenuItem.Text = "Проверить дерево состава"
        Me.ПроверитьДеревоСоставаToolStripMenuItem.ToolTipText = "Проверка на заполненность полей Формат, Масса, Материал у объектов в дереве соста" &
    "ва" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Проверка на наличиедокументов у объектов в дереве состава"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(277, 6)
        '
        'ОткрытьВСНРМToolStripMenuItem
        '
        Me.ОткрытьВСНРМToolStripMenuItem.Enabled = False
        Me.ОткрытьВСНРМToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.excel
        Me.ОткрытьВСНРМToolStripMenuItem.Name = "ОткрытьВСНРМToolStripMenuItem"
        Me.ОткрытьВСНРМToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.ОткрытьВСНРМToolStripMenuItem.Text = "Открыть ВСНРМ"
        '
        'ОтчетыToolStripMenuItem
        '
        Me.ОтчетыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ЖурналПроектовToolStripMenuItem, Me.ЖурналОТДToolStripMenuItem, Me.ПравитьОбъектToolStripMenuItem, Me.ВСНРМ2ToolStripMenuItem})
        Me.ОтчетыToolStripMenuItem.Name = "ОтчетыToolStripMenuItem"
        Me.ОтчетыToolStripMenuItem.Size = New System.Drawing.Size(70, 20)
        Me.ОтчетыToolStripMenuItem.Text = "Действие"
        '
        'ЖурналПроектовToolStripMenuItem
        '
        Me.ЖурналПроектовToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.download
        Me.ЖурналПроектовToolStripMenuItem.Name = "ЖурналПроектовToolStripMenuItem"
        Me.ЖурналПроектовToolStripMenuItem.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
            Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.ЖурналПроектовToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ЖурналПроектовToolStripMenuItem.Text = "Выгрузить КД"
        Me.ЖурналПроектовToolStripMenuItem.ToolTipText = "Выгрузить КД из Дерева состава" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Позволяет выгрузить все докумнты входящие в дан" &
    "ный проект"
        '
        'ЖурналОТДToolStripMenuItem
        '
        Me.ЖурналОТДToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.OTD
        Me.ЖурналОТДToolStripMenuItem.Name = "ЖурналОТДToolStripMenuItem"
        Me.ЖурналОТДToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.K), System.Windows.Forms.Keys)
        Me.ЖурналОТДToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ЖурналОТДToolStripMenuItem.Text = "Журнал ОТД"
        Me.ЖурналОТДToolStripMenuItem.ToolTipText = "Присводить выбраным документам инвентарный номер"
        '
        'ПравитьОбъектToolStripMenuItem
        '
        Me.ПравитьОбъектToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Order
        Me.ПравитьОбъектToolStripMenuItem.Name = "ПравитьОбъектToolStripMenuItem"
        Me.ПравитьОбъектToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
        Me.ПравитьОбъектToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ПравитьОбъектToolStripMenuItem.Text = "Править отмеч.."
        Me.ПравитьОбъектToolStripMenuItem.ToolTipText = "Править отмеченный объект"
        '
        'ВСНРМ2ToolStripMenuItem
        '
        Me.ВСНРМ2ToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Рисунок1
        Me.ВСНРМ2ToolStripMenuItem.Name = "ВСНРМ2ToolStripMenuItem"
        Me.ВСНРМ2ToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ВСНРМ2ToolStripMenuItem.Text = "ВСНРМ 2.0"
        Me.ВСНРМ2ToolStripMenuItem.ToolTipText = "Выгружает сводную ведомость состава из архива" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(В данной версии не доступна)"
        '
        'СервисToolStripMenuItem
        '
        Me.СервисToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.НастройкиToolStripMenuItem, Me.СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem})
        Me.СервисToolStripMenuItem.Name = "СервисToolStripMenuItem"
        Me.СервисToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.СервисToolStripMenuItem.Text = "Сервис"
        '
        'НастройкиToolStripMenuItem
        '
        Me.НастройкиToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.options_icon_30
        Me.НастройкиToolStripMenuItem.Name = "НастройкиToolStripMenuItem"
        Me.НастройкиToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F2), System.Windows.Forms.Keys)
        Me.НастройкиToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.НастройкиToolStripMenuItem.Text = "Настройки"
        Me.НастройкиToolStripMenuItem.ToolTipText = "Настройки приложения"
        '
        'СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem
        '
        Me.СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.IMBASE1
        Me.СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem.Name = "СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem"
        Me.СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem.Text = "Сверить ВСНРМ ..."
        Me.СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem.ToolTipText = "Сверить ВСНРМ с таблицей соответсвия"
        '
        'ПомощьToolStripMenuItem
        '
        Me.ПомощьToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СправкаToolStripMenuItem, Me.ОПрограммеToolStripMenuItem, Me.TestToolStripMenuItem})
        Me.ПомощьToolStripMenuItem.Name = "ПомощьToolStripMenuItem"
        Me.ПомощьToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.ПомощьToolStripMenuItem.Text = "Справка"
        '
        'СправкаToolStripMenuItem
        '
        Me.СправкаToolStripMenuItem.Enabled = False
        Me.СправкаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.Help1
        Me.СправкаToolStripMenuItem.Name = "СправкаToolStripMenuItem"
        Me.СправкаToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.СправкаToolStripMenuItem.Size = New System.Drawing.Size(168, 22)
        Me.СправкаToolStripMenuItem.Text = "Справка"
        Me.СправкаToolStripMenuItem.ToolTipText = "Вызвать справку"
        '
        'ОПрограммеToolStripMenuItem
        '
        Me.ОПрограммеToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.About1
        Me.ОПрограммеToolStripMenuItem.Name = "ОПрограммеToolStripMenuItem"
        Me.ОПрограммеToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F2
        Me.ОПрограммеToolStripMenuItem.Size = New System.Drawing.Size(168, 22)
        Me.ОПрограммеToolStripMenuItem.Text = "О программе"
        Me.ОПрограммеToolStripMenuItem.ToolTipText = "Информация о программе"
        '
        'TestToolStripMenuItem
        '
        Me.TestToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources._9
        Me.TestToolStripMenuItem.Name = "TestToolStripMenuItem"
        Me.TestToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.TestToolStripMenuItem.Size = New System.Drawing.Size(168, 22)
        Me.TestToolStripMenuItem.Text = "What's new"
        Me.TestToolStripMenuItem.ToolTipText = "Новые функции программы"
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.ReadOnly = True
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(170, 20)
        Me.ToolStripTextBox1.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolStripTextBox1.ToolTipText = "ФИО активного пользователя"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 24)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TreeView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.ListBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(612, 512)
        Me.SplitContainer1.SplitterDistance = 387
        Me.SplitContainer1.TabIndex = 2
        '
        'TreeView1
        '
        Me.TreeView1.AllowDrop = True
        Me.TreeView1.ContextMenuStrip = Me.ContextMenuStrip2
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.Location = New System.Drawing.Point(0, 0)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(612, 387)
        Me.TreeView1.TabIndex = 0
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ОткрытьКарточкуДокументаToolStripMenuItem, Me.ToolStripSeparator3, Me.ОткрытьВСНРМToolStripMenuItem1})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip2.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(264, 76)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Image = Global.S4_SP.My.Resources.Resources.DATABASE
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.ShortcutKeys = System.Windows.Forms.Keys.F4
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(263, 22)
        Me.ToolStripMenuItem1.Text = "&Посмотреть информацию в S4"
        '
        'ОткрытьКарточкуДокументаToolStripMenuItem
        '
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПросмотрКарточкиДокументаToolStripMenuItem, Me.РедактированиеКарточкиДокументаToolStripMenuItem})
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.newKTD
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Name = "ОткрытьКарточкуДокументаToolStripMenuItem"
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Size = New System.Drawing.Size(263, 22)
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.Text = "&Открыть карточку документа..."
        Me.ОткрытьКарточкуДокументаToolStripMenuItem.ToolTipText = "Открыть карточку документ/объекта"
        '
        'ПросмотрКарточкиДокументаToolStripMenuItem
        '
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.newKTD
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Name = "ПросмотрКарточкиДокументаToolStripMenuItem"
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F3), System.Windows.Forms.Keys)
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Size = New System.Drawing.Size(329, 22)
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.Text = "Пр&осмотр карточки документа"
        Me.ПросмотрКарточкиДокументаToolStripMenuItem.ToolTipText = "Открыть карточку документ/объекта" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Откроется карта, БЕЗ взятия на редактирование " &
    "документа" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Параметры нельзя будет изменить"
        '
        'РедактированиеКарточкиДокументаToolStripMenuItem
        '
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.newZak
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Name = "РедактированиеКарточкиДокументаToolStripMenuItem"
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Size = New System.Drawing.Size(329, 22)
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.Text = "&Редактирование карточки документа"
        Me.РедактированиеКарточкиДокументаToolStripMenuItem.ToolTipText = "Открыть карточку документ/объекта" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Откроется карточке, с правим изменение данных"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(260, 6)
        '
        'ОткрытьВСНРМToolStripMenuItem1
        '
        Me.ОткрытьВСНРМToolStripMenuItem1.Enabled = False
        Me.ОткрытьВСНРМToolStripMenuItem1.Image = Global.S4_SP.My.Resources.Resources.excel
        Me.ОткрытьВСНРМToolStripMenuItem1.Name = "ОткрытьВСНРМToolStripMenuItem1"
        Me.ОткрытьВСНРМToolStripMenuItem1.Size = New System.Drawing.Size(263, 22)
        Me.ОткрытьВСНРМToolStripMenuItem1.Text = "Открыть ВСНРМ"
        '
        'ListBox1
        '
        Me.ListBox1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(0, 0)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(612, 121)
        Me.ListBox1.TabIndex = 0
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОчиститьИсториюToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(230, 26)
        '
        'ОчиститьИсториюToolStripMenuItem
        '
        Me.ОчиститьИсториюToolStripMenuItem.Image = Global.S4_SP.My.Resources.Resources.clear_icon_9193
        Me.ОчиститьИсториюToolStripMenuItem.Name = "ОчиститьИсториюToolStripMenuItem"
        Me.ОчиститьИсториюToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Delete), System.Windows.Forms.Keys)
        Me.ОчиститьИсториюToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.ОчиститьИсториюToolStripMenuItem.Text = "Очистить историю"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(612, 536)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(628, 575)
        Me.Name = "Form1"
        Me.Text = "S4-SP"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ВыбратьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВыбратьОбъектToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПомощьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СправкаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОПрограммеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СервисToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents НастройкиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СоздатьДеревоToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TreeView1 As TreeView
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ОчиститьИсториюToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ПроверитьДеревоСоставаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents TestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОткрытьКарточкуДокументаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПросмотрКарточкиДокументаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents РедактированиеКарточкиДокументаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОтчетыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЖурналПроектовToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ОткрытьВСНРМToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОткрытьВСНРМToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents СверитьВСНРМСТаблицейСоответсвияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЖурналОТДToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripTextBox1 As ToolStripTextBox
    Friend WithEvents ПравитьОбъектToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВСНРМ2ToolStripMenuItem As ToolStripMenuItem
End Class
