from docx.shared import Cm, Pt

class template_65():
    def __init__(self) -> None:
        self.page_height = Cm(29.7)
        self.page_width = Cm(21)
        self.top_margin = Cm(1)
        self.left_margin = Cm(0.99)
        self.right_margin = Cm(0.99)
        self.bottom_margin = Cm(0.86)

        self.quantity_rows = 13
        self.quantity_columns = 5
        self.quantity_cells = 65
        self.height_row = Cm(1.72)
        self.width_cell_QR = Cm(1.2)
        self.width_cell_descr = Cm(2.6)
        self.font_size = Pt(5)
        self.height_pic = min(self.height_row, self.width_cell_QR) * 0.9

        self.name_project = "Инвентаризация"

        self.indents_cells_QR = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'0'
        }
        self.indents_cells_descr = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'113'
        }

        self.type_save_file = "Документ Word (*.docx)"
        self.default_dir_name = "./QR коды"
        self.list_num_col_for_QR = [1]

class template_40():
    def __init__(self) -> None:
        self.page_height = Cm(29.7)
        self.page_width = Cm(21)
        self.top_margin = Cm(0)
        self.left_margin = Cm(0)
        self.right_margin = Cm(0)
        self.bottom_margin = Cm(0)

        self.quantity_rows = 10
        self.quantity_columns = 4
        self.quantity_cells = 40
        self.height_row = Cm(2.57)
        self.width_cell_QR = Cm(1.68)
        self.width_cell_descr = Cm(3.57)
        self.font_size = Pt(6)
        self.height_pic = min(self.height_row, self.width_cell_QR) * 0.9

        self.name_project = "Инвентаризация"

        self.indents_cells_QR = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'0'
        }
        self.indents_cells_descr = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'113'
        }

        self.type_save_file = "Документ Word (*.docx)"
        self.default_dir_name = "./QR коды"
        self.list_num_col_for_QR = [1]

class template_24():
    def __init__(self) -> None:
        self.page_height = Cm(29.7)
        self.page_width = Cm(21)
        self.top_margin = Cm(0)
        self.left_margin = Cm(0)
        self.right_margin = Cm(0)
        self.bottom_margin = Cm(0)

        self.quantity_rows = 8
        self.quantity_columns = 3
        self.quantity_cells = 24
        self.height_row = Cm(3.31) # 3.71 без полей
        self.width_cell_QR = Cm(2.24)
        self.width_cell_descr = Cm(4.76)
        self.font_size = Pt(7)
        self.height_pic = min(self.height_row, self.width_cell_QR) * 0.9

        self.name_project = "Инвентаризация"

        self.indents_cells_QR = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'0'
        }
        self.indents_cells_descr = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'113'
        }

        self.type_save_file = "Документ Word (*.docx)"
        self.default_dir_name = "./QR коды"
        self.list_num_col_for_QR = [1]


class template_12():
    def __init__(self) -> None:
        self.page_height = Cm(29.7)
        self.page_width = Cm(21)
        self.top_margin = Cm(0.45)
        self.left_margin = Cm(0.22)
        self.right_margin = Cm(0.79)
        self.bottom_margin = Cm(0.24)

        self.quantity_rows = 6
        self.quantity_columns = 2
        self.quantity_cells = 12
        self.height_row = Cm(4.4) # 4.8 без полей
        self.width_cell_QR = Cm(3.36) #32% от ширины
        self.width_cell_descr = Cm(7.14) #68% от ширины
        self.font_size = Pt(10)
        self.height_pic = min(self.height_row, self.width_cell_QR) * 0.9

        self.name_project = "Инвентаризация"

        self.indents_cells_QR = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'0'
        }
        self.indents_cells_descr = {
            'top':'113',
            'left':'0',
            'bottom':'113',
            'right':'113'
        }

        self.type_save_file = "Документ Word (*.docx)"
        self.default_dir_name = "./QR коды"
        self.list_num_col_for_QR = [1]