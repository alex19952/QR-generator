from PyQt5.QtWidgets import (QPushButton, QWidget, 
QDesktopWidget, QMessageBox, QVBoxLayout, QLabel, 
QFrame, QHBoxLayout, QGridLayout, QButtonGroup, 
QDialog, QSizePolicy, QApplication, QCheckBox)

from PyQt5.QtGui import QIcon, QPixmap, QMovie, QPaintEvent

from PyQt5.QtCore import Qt

from os import popen as os_popen

import Templates


class MyPushButton(QPushButton):

    def __init__(self, name, group, path_resourses, parent=None):

        self.myGroup = group
        
        self.myParent = super(MyPushButton, self).__init__(parent)        
        
        self.name = name
        
        self.path_resourses = path_resourses

        self.style_active = """
            QPushButton {
                border: none;
                margin: 0px;
                padding: 0px;
                border-image: url(""" + self.path_resourses + """active_""" + self.name + """.png); 
            }
            QPushButton:hover {
                border-image: url(""" + self.path_resourses + """enter_""" + self.name + """.png);     
        }
            QPushButton:pressed {
                border-image: url(""" + self.path_resourses + """down_""" + self.name + """.png);
            }
            """

        self.style_notActive = """
            QPushButton {
                border: none;
                margin: 0px;
                padding: 0px;
                border-image: url(""" + self.path_resourses + """not_active_""" + str(self.name) + """.png); 
            }
            QPushButton:hover {
                border-image: url(""" + self.path_resourses + """enter_""" + str(self.name) + """.png);     
        }
            QPushButton:pressed {
                border-image: url(""" + path_resourses + """down_""" + str(self.name) + """.png);
            }
            """

        self.setStyleSheet(self.style_notActive)


    def mouseReleaseEvent(self, event):

        super(MyPushButton, self).mouseReleaseEvent(event)

        self.setStyleSheet(self.style_notActive)

        for button in self.myGroup.buttons():
            button.setStyleSheet(button.style_notActive)

        self.setStyleSheet(self.style_active)


    def click(self) -> None:

        self.setStyleSheet(self.style_active)

        return super().click()


class Window(QWidget):

    def __init__(self, name_project, path_resourses):

        super().__init__()

        self.name_window = name_project

        self.path_resourses = path_resourses

        self.setFixedSize(600, 450)

        qr = self.frameGeometry()

        cp = QDesktopWidget().availableGeometry().center()

        qr.moveCenter(cp)

        self.move(qr.topLeft())

        self.setWindowTitle(self.name_window)

        self.setWindowIcon(QIcon(self.path_resourses + 'icon.png'))

        self.setStyleSheet("background-color:white;")

        self.initUI()

        self.show()


    def paintEvent(self, a0: QPaintEvent) -> None:

        if self.flag_save:

            self.flag_save = False

            msg = QMessageBox()

            msg.setWindowTitle("Открыть файл")

            msg.setWindowIcon(QIcon(self.path_resourses + 'icon.png'))

            msg.setText("Открыть созданный файл?")

            buttonAceptar  = msg.addButton("Да", QMessageBox.YesRole)

            open_file = lambda: os_popen("\"" + self.path_doc_file + "\"")

            buttonCancelar = msg.addButton("Нет", QMessageBox.RejectRole)

            buttonAceptar.clicked.connect(open_file)

            msg.setDefaultButton(buttonAceptar)

            msg.setWindowFlag(Qt.WindowStaysOnTopHint)

            msg.exec_()
     
        return super().paintEvent(a0)


    def initUI(self):

        self.flag_save = False

        self.create_base_for_DCT = False
        
        self.path_doc_file = ''

        self.num_start_cell = 1

        self.list_template = [  Templates.template_65, 
                                Templates.template_40, 
                                Templates.template_24, 
                                Templates.template_12
                                ]

        self.label_list =    [
            '''
            <div style="text-align: center;"><font color="#A9A9A9"><b><font size="6"><font face="Verdana, Geneva, sans-serif">65 этикеток</font></font></b></font></div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="5"><font face="Verdana, Geneva, sans-serif">38 x 21.2 mm</font></font></font></div>
            <div style="text-align: center;">&nbsp;</div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="4"><font face="Verdana, Geneva, sans-serif">&quot;Europe100&quot; ELA001</font></font></font></div>
            ''',

            '''
            <div style="text-align: center;"><font color="#A9A9A9"><b><font size="6"><font face="Verdana, Geneva, sans-serif">40 этикеток</font></font></b></font></div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="5"><font face="Verdana, Geneva, sans-serif">52.5 x 29.7 mm</font></font></font></div>
            <div style="text-align: center;">&nbsp;</div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="4"><font face="Verdana, Geneva, sans-serif">&quot;Europe100&quot; ELA049</font></font></font></div>
            ''',

            '''
            <div style="text-align: center;"><font color="#A9A9A9"><b><font size="6"><font face="Verdana, Geneva, sans-serif">24 этикетки</font></font></b></font></div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="5"><font face="Verdana, Geneva, sans-serif">70 x 37 mm</font></font></font></div>
            <div style="text-align: center;">&nbsp;</div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="4"><font face="Verdana, Geneva, sans-serif">&quot;Europe100&quot; ELA011</font></font></font></div>
            ''',

            '''
            <div style="text-align: center;"><font color="#A9A9A9"><b><font size="6"><font face="Verdana, Geneva, sans-serif">12 этикеток</font></font></b></font></div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="5"><font face="Verdana, Geneva, sans-serif">105 x 48 mm</font></font></font></div>
            <div style="text-align: center;">&nbsp;</div>
            <div style="text-align: center;"><font color="#A9A9A9"><font size="4"><font face="Verdana, Geneva, sans-serif">&quot;Europe100&quot; ELA021</font></font></font></div>
            '''
            ]

        self.logo = self.LogoLabel()

        self.frame_menu = self.MenuFrame()

        self.frame_load = self.LoadFrame()
        
        self.frame_load.hide()

        self.main_box = QVBoxLayout(self)

        self.main_box.setContentsMargins(0,0,0,0)

        self.main_box.addWidget(self.logo)

        self.main_box.addWidget(self.frame_menu)

        self.main_box.addWidget(self.frame_load)

        self.main_box.addStretch(5)


    def LogoLabel(self):

        logo = QLabel(self)

        pixmap = QPixmap(self.path_resourses + 'logo4.png')

        logo.setFixedHeight(82)

        logo.setPixmap(pixmap)

        return logo


    def MenuFrame(self):

        frame = QFrame()

        frame.setLayout(self.MenuGridLayout())

        frame.setWindowOpacity(-50.5)

        return frame


    def LoadFrame(self):

        frame = QFrame()

        load_layout = QHBoxLayout()

        labAnim = QLabel()

        self.gif = QMovie(self.path_resourses + 'map2.gif')

        labAnim.setMovie(self.gif)

        load_layout.addStretch(5)

        load_layout.addWidget(labAnim)

        load_layout.addStretch(5)

        frame.setLayout(load_layout)

        self.gif.setCacheMode(QMovie.CacheMode.CacheAll)

        return frame


    def MenuGridLayout(self):

        grid_layout = QGridLayout()

        grid_layout.setContentsMargins(0,0,0,0)

        grid_layout.setRowMinimumHeight(0, 30)

        grid_layout.setRowMinimumHeight(2, 30)

        grid_layout.setColumnMinimumWidth(0, 30)

        grid_layout.setColumnMinimumWidth(3, 15)

        grid_layout.setColumnMinimumWidth(5, 30)

        right_layout = self.InfoAndStartLayout()

        grid_layout.addLayout(right_layout, 1, 4)

        left_layout = self.TemplatesLayout()

        grid_layout.addLayout(left_layout, 1, 1, 1, 2, Qt.AlignLeft|Qt.AlignTop)

        return grid_layout


    def TemplatesLayout(self):
        
        grid = QGridLayout()

        self.button_group_templates = QButtonGroup()

        self.button_group_templates.setExclusive(True)

        name_list = ['65', '40', '24', '12']

        for index, name_button in enumerate(name_list):
            button = MyPushButton(name_button, self.button_group_templates, self.path_resourses)
            button.setCheckable(True)
            button.setFixedSize(100,142)
            self.button_group_templates.addButton(button, index)
            grid.addWidget(button, index // 3, index % 3)

        self.button_group_templates.buttonClicked.connect(self.TemplateButtonClick)
        
        self.button_group_templates.buttons()[0].click()
        
        return grid


    def InfoAndStartLayout(self):

        box = QVBoxLayout()

        box.setContentsMargins(0,0,0,0)

        self.label_template_descr = QLabel()

        self.label_for_start_cell_selection_button = self.LabelStartCell(self.num_start_cell)

        self.start_cell_selection_button = self.StartCellSelectionButton()

        self.start_cell_selection_button.clicked.connect(self.StartCellSelectionMenu)
        
        self.create_base_check_box = self.CreateBaseCheckBox()

        self.create_base_check_box.stateChanged.connect(self.changeFlagForDCTBase)

        self.create_base_check_box.setCheckState(False)

        self.main_bytton = self.MainButton()

        box.addWidget(self.label_template_descr)

        box.addStretch(5)

        box.addWidget(self.label_for_start_cell_selection_button)

        box.addWidget(self.start_cell_selection_button)

        box.addStretch(5)

        box.addWidget(self.create_base_check_box)

        box.addStretch(2)

        box.addWidget(self.main_bytton)

        return box


    def StartCellSelectionMenu(self):

        index = self.button_group_templates.checkedId()

        TemplClass = self.list_template[index]
        
        self.var = TemplClass()

        self.win_start_cell_selection_menu = QDialog()

        self.win_start_cell_selection_menu.setFixedSize(318, 450)

        self.win_start_cell_selection_menu.setWindowTitle("Выбор начальной ячейки")

        self.win_start_cell_selection_menu.setWindowIcon(QIcon(self.path_resourses + 'icon.png'))

        self.win_start_cell_selection_menu.setStyleSheet("background-color:white;")

        qr = self.win_start_cell_selection_menu.frameGeometry()

        cp = QDesktopWidget().availableGeometry().center()

        qr.moveCenter(cp)

        self.win_start_cell_selection_menu.move(qr.topLeft())

        box = QGridLayout(spacing = 0)

        self.button_group_start_cell = QButtonGroup()

        for i in range(self.var.quantity_rows + 1):
            for j in range(self.var.quantity_columns + 1):
                if i == 0 and j == 0:
                    pass
                elif i == 0:
                    label = QLabel()
                    label.setText(  """<div style="text-align: center;">""" + 
                                    str(j) +
                                    """</div>""")
                    box.addWidget(label, i, j)
                elif j == 0:
                    label = QLabel()
                    label.setText(  """<div style="text-align: center;">""" + 
                                    str(i) +
                                    """</div>""")
                    box.addWidget(label, i , j)
                else:
                    index = (i - 1) * self.var.quantity_columns + j
                    button = QPushButton(self)
                    button.setObjectName(str(index))
                    button.setSizePolicy(
                        QSizePolicy.Expanding,
                        QSizePolicy.Expanding)
                    if index < self.num_start_cell:
                        button.setStyleSheet("""
                        QPushButton 
                            {
                            border-image: url(""" + self.path_resourses + """3.jpg);
                            }
                        QPushButton:hover 
                            {
                            border-image: url(""" + self.path_resourses + """4.jpg);
                            }
                        QPushButton:default
                            { border-image: url(""" + self.path_resourses + """4.jpg); }
                            """)
                    else:
                        button.setStyleSheet("""
                        QPushButton 
                            {
                            border-image: url(""" + self.path_resourses + """2.jpg);
                            }
                        QPushButton:hover 
                            {
                            border-image: url(""" + self.path_resourses + """1.jpg);
                            }
                        QPushButton:default
                            { border-image: url(""" + self.path_resourses + """1.jpg); }
                            """)
                    if int(button.objectName()) == self.num_start_cell:
                        button.setDefault(True)
                    button.clicked.connect(self.setNumStartCell)
                    box.addWidget(button, i, j)
                    self.button_group_start_cell.addButton(button, index)

        self.win_start_cell_selection_menu.setLayout(box)

        self.win_start_cell_selection_menu.exec_()


    def setNumStartCell(self):

        button = QApplication.instance().sender()

        self.num_start_cell = int(button.objectName())

        text = self.getTextLabelStartCell(self.num_start_cell)

        self.label_for_start_cell_selection_button.setText(text)

        for i in range(self.var.quantity_cells):
            temp_button = self.button_group_start_cell.buttons()[i]
            if i < self.num_start_cell - 1:
                temp_button.setStyleSheet("""
                    QPushButton 
                        {
                        border-image: url(""" + self.path_resourses + """3.jpg);
                        }
                    QPushButton:hover 
                        {
                        border-image: url(""" + self.path_resourses + """4.jpg);
                        }
                    QPushButton:default
                        { border-image: url(""" + self.path_resourses + """4.jpg); }
                        """)
            else:
                temp_button.setStyleSheet("""
                    QPushButton 
                        {
                        border-image: url(""" + self.path_resourses + """2.jpg);
                        }
                    QPushButton:hover 
                        {
                        border-image: url(""" + self.path_resourses + """1.jpg);s
                        }
                    QPushButton:default
                        { border-image: url(""" + self.path_resourses + """1.jpg); }
                        """)


    def TemplateButtonClick(self):

        index = self.button_group_templates.checkedId()

        text_1 = self.label_list[index]

        self.label_template_descr.setText(text_1)

        self.template_index = index
        
        self.num_start_cell = 1

        text_2 = self.getTextLabelStartCell(self.num_start_cell)

        self.label_for_start_cell_selection_button.setText(text_2)


    def getTextLabelStartCell(self, num_start_cell):

        text = '''
            <div style="text-align: center;">
            <font color="#A9A9A9">
            <font size="3">
            <font face="Verdana, Geneva, sans-serif">
            Шаблон заполняется<br>
            с ''' + str(num_start_cell) + ''' ячейки<br>
            <br>
            Нажмите, чтобы изменить:
            </font></font></font></div>
        '''

        return text


    def StartCellSelectionButton(self):

        button = QPushButton()

        style = """
            QPushButton {
                border: none;
                margin: 0px;
                padding: 0px;
                border-image: url(""" + self.path_resourses + """button_change_default.png); 
            }
            QPushButton:pressed {
                border-image: url(""" + self.path_resourses + """button_change_pressed.png);
            }
            """
        button.setStyleSheet(style)

        button.setFixedSize(200,40)

        return button


    def MainButton(self):
        
        button = QPushButton()

        style = """
            QPushButton {
                border: none;
                margin: 0px;
                padding: 0px;
                border-image: url(""" + self.path_resourses + """def_main_button.png); 
            }
            QPushButton:pressed {
                border-image: url(""" + self.path_resourses + """press_main_button.png);
            }
            """
        button.setStyleSheet(style)

        button.setFixedSize(200,40)
        
        return button


    def CreateBaseCheckBox (self):

        cb = QCheckBox("Создать базу для ТСД", self)

        cb.toggle()
        
        return cb


    def changeFlagForDCTBase(self, state):

        if state == Qt.Checked:
            self.create_base_for_DCT = True
        else:
            self.create_base_for_DCT = False


    def LabelStartCell(self, num_start_cell_fefault):

        label = QLabel()

        label.setFixedWidth(200)

        text = self.getTextLabelStartCell(num_start_cell_fefault)

        label.setText(text)

        label.setWordWrap(True)

        return label
