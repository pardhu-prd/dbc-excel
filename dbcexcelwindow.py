from dbcexcellogic import DbcExcelLogic
import sys
from PyQt5.QtWidgets import (
    QMainWindow,
    QApplication,
    QPushButton,
    QFileDialog,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QComboBox,
    QScrollArea,
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt


class DbcWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to DBC Converter")
        self.setGeometry(500, 300, 800, 500)
        self.setWindowIcon(
            QIcon(r"C:\Users\Pardhasaradhi\Desktop\icons\fotor-ai-2023100610743.jpg")
        )
        self.dbcexcellogic = DbcExcelLogic()

        output_excel_file = None
        self.selected_file_path = ""
        self.column_mappings = {}  # To store user-defined mappings

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        left_side_layout = QVBoxLayout()
        right_side_layout = QVBoxLayout()
        selected_output_layout = QVBoxLayout()

        # Left-side layout for buttons and labels
        dbc_excel_buttons_layout = QHBoxLayout()

        self.dbc_to_excel_button = QPushButton("DBC to Excel")
        self.dbc_to_excel_button.clicked.connect(self.open_dbc_to_excel)
        dbc_excel_buttons_layout.addWidget(self.dbc_to_excel_button)

        self.excel_to_dbc_button = QPushButton("Excel to DBC")
        self.excel_to_dbc_button.clicked.connect(self.open_excel_to_dbc)
        dbc_excel_buttons_layout.addWidget(self.excel_to_dbc_button)

        left_side_layout.addLayout(dbc_excel_buttons_layout)

        self.selected_file_label = QLabel("Selected File: ")
        selected_output_layout.addWidget(self.selected_file_label)

        self.mapped_label = QLabel(
            "Map the EXCEL column index to the Alphabet to your right box"
        )
        selected_output_layout.addWidget(self.mapped_label)
        self.mapped_label.setStyleSheet(
            "font-weight: bold; font-size: 14px; color: red;"
        )
        self.mapped_label.hide()

        self.convert_button = QPushButton("Convert")
        self.convert_button.setEnabled(False)
        self.convert_button.clicked.connect(self.convert_files)
        selected_output_layout.addWidget(self.convert_button)

        self.output_label = QLabel("Output: ")
        selected_output_layout.addWidget(self.output_label)

        left_side_layout.addLayout(selected_output_layout)

        main_layout.addLayout(left_side_layout, 3)

        # Right-side layout with alphabet buttons and combo boxes in a scrollable area
        alphabet_scroll_area = QScrollArea()
        alphabet_scroll_area.setWidgetResizable(True)
        alphabet_widget = QWidget()
        alphabet_layout = QVBoxLayout(alphabet_widget)

        alphabet_buttons = [str(i) for i in range(16)]

        alphabet_index_label = QLabel(
            "Map the DBC parameters\n to Excel Column Indexes"
        )
        alphabet_index_label.setAlignment(Qt.AlignCenter)
        alphabet_layout.addWidget(alphabet_index_label)

        self.column_name_boxes = {}

        for alphabet in alphabet_buttons:
            alphabet_row_layout = QHBoxLayout()
            alphabet_button = QPushButton(alphabet)
            alphabet_button.setFixedWidth(40)
            alphabet_button.setFixedHeight(30)
            alphabet_button.setEnabled(False)

            alphabet_row_layout.addWidget(alphabet_button)

            combo_box = QComboBox()
            combo_box.setEnabled(False)
            combo_box.addItems(
                [
                    "None",
                    "CAN ID",
                    "Decimal",
                    "CANID Type",
                    "Message Name",
                    "DLC",
                    "Comments",
                    "Signal Name",
                    "Start Bit",
                    "Length",
                    "Unit",
                    "Data Type",
                    "Offset",
                    "Minimum",
                    "Maximum",
                    "Endianness",
                    "Scale",
                ]
            )
            combo_box.currentIndexChanged.connect(
                lambda index, alpha=alphabet: self.update_mapping(alpha, index)
            )

            alphabet_row_layout.addWidget(combo_box)
            self.column_name_boxes[alphabet] = combo_box

            alphabet_layout.addLayout(alphabet_row_layout)

        alphabet_widget.setLayout(alphabet_layout)
        alphabet_scroll_area.setWidget(alphabet_widget)
        right_side_layout.addWidget(alphabet_scroll_area)

        main_layout.addLayout(right_side_layout, 1)

    def open_excel_to_dbc(self):
        for combo_box in self.column_name_boxes.values():
            combo_box.setEnabled(True)

        self.selected_file_path = self.dbcexcellogic.get_excel_file()
        self.selected_file_label.setText(f"Selected File:\n{self.selected_file_path}")
        self.mapped_label.show()
        self.convert_button.setEnabled(True)

    def open_dbc_to_excel(self):
        self.selected_file_path = self.dbcexcellogic.get_dbc_file()
        self.selected_file_label.setText(f"Selected File:\n{self.selected_file_path}")

        self.convert_button.setEnabled(True)

    def update_mapping(self, alphabet, index):
        data = self.column_name_boxes[alphabet].currentText()
        self.column_mappings[alphabet] = data

    def convert_files(self):
        if self.selected_file_path.endswith(".dbc"):
            try:
                output_excel_file = self.dbcexcellogic.convert_dbc_to_excel()
                self.output_label.setText(
                    f"output: DBC to Excel conversion successful\n {output_excel_file}"
                )
            except Exception as e:
                self.output_label.setText(f"Cannot convert to excel.\n Error: {e}")

        elif self.selected_file_path.endswith(
            ".xls"
        ) or self.selected_file_path.endswith(".xlsx"):
            try:
                output_dbc_file = self.dbcexcellogic.process_excel_to_dbc(
                    self.column_mappings
                )
                self.output_label.setText(
                    f"output:Excel to DBC conversion successful\n {output_dbc_file}"
                )
            except Exception as e:
                self.output_label.setText(f"Cannot convert to excel.\n Error: {e}")

        else:
            self.output_label.setText("Please select the appropriate files to convert")


def run_main():
    app = QApplication(sys.argv)
    window = DbcWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    run_main()
