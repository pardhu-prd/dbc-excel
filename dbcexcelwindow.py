"""DBC-EXCEL WINDOW
This module defines the DBC-EXCEL Window class for handling buttons in a PyQt5 GUI.
"""
import sys
from PyQt5.QtWidgets import (  # pylint: disable=no-name-in-module
    QMainWindow,
    QApplication,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QComboBox,
    QScrollArea,
    QMessageBox,
)
from PyQt5.QtGui import QIcon   # pylint: disable=no-name-in-module
from PyQt5.QtCore import Qt # pylint: disable=no-name-in-module
from dbcexcellogic import DbcExcelLogic


class DbcWindow(QMainWindow):
    """Responsible for managing buttons and mapping Excel column numbers to DBC parameters."""

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel to DBC Converter")
        self.setGeometry(500, 300, 800, 500)
        self.setWindowIcon(
            QIcon(r"C:\Users\Pardhasaradhi\Desktop\icons\fotor-ai-2023100610743.jpg")
        )
        self.dbcexcellogic = DbcExcelLogic()
        self.dbcexcellogic.exception_occurred.connect(  # pylint: disable=no-member
            self.handle_exception
        )
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
            "Map the parameters to the corresponding column indexes,like in the Excel file.")
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

        # Right-side layout with alphabet/column indexes buttons and combo boxes in a scrolable area
        # Alphabet = column index

        alphabet_scroll_area = QScrollArea()
        alphabet_scroll_area.setWidgetResizable(True)
        alphabet_widget = QWidget()
        alphabet_layout = QVBoxLayout(alphabet_widget)

        alphabet_buttons = [str(i) for i in range(16)]

        alphabet_index_label = QLabel(
            "Map the DBC parameters\n to Excel Column Indexes\n"
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

    def handle_exception(self, exception):
        """slot (function) to handle exceptions"""
        error_message = (
            f"Error: {str(exception)}\nPlease ensure your Excel file contains all the required parameters in separate columns "
            f"and that you have correctly mapped these parameters to the corresponding column indexes in the Excel file."
            f"\n\nFor example, if the 'CAN ID' is in column index 2, make sure you have mapped it accordingly in the same way."
            f"\nColumn indexes are always start from 0 in your Excel"
        )

        QMessageBox.critical(self, "Error", error_message)

    def open_excel_to_dbc(self):
        """Enabling the combo boxes, opens the Filedialog and selects the file"""
        for combo_box in self.column_name_boxes.values():
            combo_box.setEnabled(True)

        self.selected_file_path = self.dbcexcellogic.get_excel_file()
        self.selected_file_label.setText(f"Selected File:\n{self.selected_file_path}")
        self.mapped_label.show()
        self.convert_button.setEnabled(True)

    def open_dbc_to_excel(self):
        """Open the Filedialog and selects the dbc file"""
        self.selected_file_path = self.dbcexcellogic.get_dbc_file()
        self.selected_file_label.setText(f"Selected File:\n{self.selected_file_path}")
        self.convert_button.setEnabled(True)

    def update_mapping(self, alphabet, index):
        """Updating the column indexes with combo box parameters"""
        data = self.column_name_boxes[alphabet].currentText()
        self.column_mappings[alphabet] = data

    def convert_files(self):
        """Converts the one file into another file, 
        if file is dbc then converts to the excel or if file is excel then converts to dbc"""
        if self.selected_file_path.endswith(".dbc"):
            try:
                output_excel_file = self.dbcexcellogic.convert_dbc_to_excel()
                self.output_label.setText(
                    f"DBC to Excel conversion successful\n {output_excel_file}"
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
                if output_dbc_file is not None:
                    self.output_label.setText(
                        f"Excel to DBC conversion successful\n {output_dbc_file}"
                    )
            except Exception as e:
                self.output_label.setText(f"Cannot convert to excel.\n Error: {e}")

        else:
            self.output_label.setText("Please select the appropriate files to convert")


def run_main():
    """Runs the main file"""
    app = QApplication(sys.argv)
    window = DbcWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    run_main()
