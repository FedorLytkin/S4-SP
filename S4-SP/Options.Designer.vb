<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Options
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ОткрытьФайлToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(264, 526)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Сохранить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(345, 526)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "Отмена"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(5, 33)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(88, 26)
        Me.Button3.TabIndex = 11
        Me.Button3.Text = "Обзор"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.TextBox1.Location = New System.Drawing.Point(98, 33)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(308, 44)
        Me.TextBox1.TabIndex = 10
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОткрытьФайлToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(154, 26)
        '
        'ОткрытьФайлToolStripMenuItem
        '
        Me.ОткрытьФайлToolStripMenuItem.Name = "ОткрытьФайлToolStripMenuItem"
        Me.ОткрытьФайлToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
        Me.ОткрытьФайлToolStripMenuItem.Text = "Открыть файл"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(292, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Изменить ссылку на таблицу соответствия материалов"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(9, 19)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(131, 17)
        Me.CheckBox1.TabIndex = 12
        Me.CheckBox1.Text = "Поиск подкаталогов"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TextBox4)
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(417, 130)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Ссылочные данные"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CheckBox1)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 139)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(169, 41)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Настройки поиска файлов"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button5)
        Me.GroupBox3.Controls.Add(Me.Button4)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.ListBox2)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.ListBox1)
        Me.GroupBox3.Location = New System.Drawing.Point(3, 186)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(416, 210)
        Me.GroupBox3.TabIndex = 15
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Настройки списка Типа файлов"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(175, 61)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(64, 23)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "-"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(175, 32)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(64, 23)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "+"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(243, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(143, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Пользовательский список"
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(245, 32)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(165, 173)
        Me.ListBox2.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(157, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Список доступных в S4 типов"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Items.AddRange(New Object() {"Спецификация", "Сборочный чертеж (AutoCAD)", "Техпроцесс", "Документ (TXT)", "Документ (DOC)", "Габаритный чертеж (AutoCAD)", "Чертеж общего вида (AutoCAD)", "Монтажный чертеж (AutoCAD)", "Электромонтажный чертеж (AutoCAD)", "Схема электрическая структурная (AutoCAD)", "Схема электрическая функциональная (AutoCAD)", "Схема электрическая принципиальная (AutoCAD)", "Схема электрическая соединений (AutoCAD)", "Схема электрическая подключения (AutoCAD)", "Схема электрическая общая (AutoCAD)", "Схема электрическая расположения (AutoCAD)", "Схема электрическая прочая (AutoCAD)", "Технические условия (DOC)", "Чертеж детали (AutoCAD)", "Перечень элементов принципиальной схемы (InterMech)", "Ведомость ссылочных документов (TXT)", "Типовой техпроцесс", "Документ (InterMech)", "Выкопировка (LCAD)", "Упаковочный чертеж (AutoCAD)", "Расцеховочный маршрут (InterMech)", "Планировка LCAD", "Модель сборки (SolidWorks)", "Модель детали (SolidWorks)", "Модель детали (SolidEdge)", "Модель сборки (SolidEdge)", "Модель детали листовой (SolidEdge)", "Модель сборки (Inventor)", "Модель детали (Inventor)", "Чертеж детали (Inventor)", "Чертеж детали (SolidWorks)", "Чертеж детали (SolidEdge)", "Проект ImProject", "Ведомость покупных изделий (InterMech)", "Ведомость спецификаций (InterMech)", "Таблица наборов зажимов", "Таблица соединений (развернутая форма)", "Таблица соединений (сжатая форма)", "Таблица соединений внешнего монтажа", "Таблица подключений", "Кабельный журнал (InterMech)", "Чертеж основного комплекта (AutoCAD)", "Общие данные основного комплекта рабочих чертежей (AutoCAD)", "Модель детали (Pro/ENGINEER)", "Модель сборки (Pro/ENGINEER)", "Чертеж детали (Pro/ENGINEER)", "Модель сборки (Unigraphics)", "Макет (Pro/ENGINEER)", "Групповой техпроцесс (InterMech)", "Комплект ведомостей", "Комплект технологических документов (Intermech)", "Модель детали (Unigraphics)", "Чертеж детали (Unigraphics)", "Служебная записка", "Дополнительное извещение", "Комплект извещений", "Извещение об изменении", "Предварительное извещение", "Предложение об изменении", "Производственный заказ (DOC)", "Спецификация электрическая", "Спецификация ЕСПД", "Перечень спецификаций (InterMech)", "Спецификация заказа", "Модель детали (Компас)", "Модель сборки (Компас)", "Сборочный чертеж (SolidWorks)", "Сборочный чертеж (Inventor)", "Сборочный чертеж (SoldEdge)", "Сборочный чертеж (Pro/ENGINEER)", "Сборочный чертеж (Unigraphics)", "Сборочный чертеж (Компас)", "Чертеж детали (Компас)", "Файл проекта (Inventor)", "Чертеж детали (TIF)", "Чертеж детали (TIFF)", "Чертеж детали (PDF)", "Паспорт (PDF)", "Сборочный чертеж (TIF)", "Сборочный чертеж (TIFF)", "Сборочный чертеж (PDF)", "Документ (TIF)", "ВСНРМ (ХLSX)", "Спецификация (PDF)", "Документ (PDF)", "Ведомость эксплуатационных документов (PDF)", "Руководство по ремонту (PDF)", "Руководство по эксплуатации (PDF)", "Технические условия на ремонт (PDF)", "Схема электрическая подключения (TIF)", "Схема электрическая подключения (TIFF)", "Схема электрическая подключения (PDF)", "Документ (XLSX)", "Документ (DOCX)", "Макет (Inventor)", "Документ (Inventor)", "Программа и методика испытаний (DOC)", "Программа и методика испытаний (DOCX)", "Модель (Sat)", "Перечень элементов принципиальной схемы (TIF)", "Перечень элементов принципиальной схемы (PDF)", "Перечень элементов принципиальной схемы (TIFF)", "Схема электрическая принципиальная (PDF)", "Техническое задание (PDF)", "Программа и методика испытаний (PDF)", "Модель (STP)", "Деталь (IAM)", "Лекало"})
        Me.ListBox1.Location = New System.Drawing.Point(4, 32)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(165, 173)
        Me.ListBox1.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.TextBox3)
        Me.GroupBox4.Controls.Add(Me.TextBox2)
        Me.GroupBox4.Location = New System.Drawing.Point(3, 402)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(416, 76)
        Me.GroupBox4.TabIndex = 17
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Список спец.слов"
        '
        'TextBox3
        '
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox3.Location = New System.Drawing.Point(5, 13)
        Me.TextBox3.Multiline = True
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(157, 57)
        Me.TextBox3.TabIndex = 18
        Me.TextBox3.Text = "Атрибуты для номера версии (Вводить через точку с запятой"";""):"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(174, 13)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(236, 57)
        Me.TextBox2.TabIndex = 11
        Me.TextBox2.Text = "изм.; зам."
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.CheckBox2)
        Me.GroupBox5.Location = New System.Drawing.Point(178, 139)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(241, 41)
        Me.GroupBox5.TabIndex = 18
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Настройки Видимости ВСНРМ"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(9, 19)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(120, 17)
        Me.CheckBox2.TabIndex = 12
        Me.CheckBox2.Text = "Таблица ВИДИМА"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(1, 79)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(156, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Рабочая папка по умолчанию"
        '
        'TextBox4
        '
        Me.TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox4.ContextMenuStrip = Me.ContextMenuStrip2
        Me.TextBox4.Location = New System.Drawing.Point(98, 97)
        Me.TextBox4.Multiline = True
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ReadOnly = True
        Me.TextBox4.Size = New System.Drawing.Size(308, 26)
        Me.TextBox4.TabIndex = 13
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(5, 97)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(88, 26)
        Me.Button6.TabIndex = 14
        Me.Button6.Text = "Обзор"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(157, 26)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(156, 22)
        Me.ToolStripMenuItem1.Text = "Открыть папку"
        '
        'Options
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(424, 561)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximumSize = New System.Drawing.Size(440, 600)
        Me.MinimumSize = New System.Drawing.Size(440, 600)
        Me.Name = "Options"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents Label2 As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ОткрытьФайлToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents Button6 As Button
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
End Class
