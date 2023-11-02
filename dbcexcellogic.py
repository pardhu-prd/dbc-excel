"""DBC-Excel Logic"""
import os
import pandas as pd
import cantools
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtCore import QObject, pyqtSignal


class DbcExcelLogic(QObject):
    """This class do the back-end process of converting one file to another"""

    exception_occurred = pyqtSignal(Exception)

    def __init__(self):
        super().__init__()
        self.selected_file_path = None

    def get_dbc_file(self):
        """Get DBC file to load. This is a dialog for choosing a DBC file for a database."""

        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(
            None,
            "Select DBC File",
            "",
            "DBC Files (*.dbc);;All Files (*)",
            options=options,
        )
        # If file_name is a file name return the file name.
        if file_name:
            self.selected_file_path = file_name
            return file_name

    def get_excel_file(self):
        """Get EXCEL file to load. This is a dialog for choosing a EXCEL file for a database."""
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(
            None,
            "Select Excel File",
            "",
            "Excel Files (*.xls *.xlsx);;All Files (*)",
            options=options,
        )
        if file_name:
            self.selected_file_path = file_name
            return file_name

    def convert_dbc_to_excel(self):
        """Converts the DBC to EXCEL"""

        dbc_file = self.selected_file_path
        try:
            database = cantools.database.load_file(dbc_file)
        except AttributeError:
            return

        data = []

        # Define a function to determine data type and bits based on the signal definition
        def determine_data_type_and_bits(signal):
            if signal.is_signed:
                data_type = "signed"
            else:
                data_type = "unsigned"

            bits = signal.length

            return f"{data_type} {bits}"

        comments = ""
        # Iterate through messages and signals to populate the data list
        for message in database.messages:
            can_id_str = message.name
            if "CAN" in can_id_str:
                can_id = can_id_str.split("_")[2]
            else:
                can_id = can_id_str.split("_")[1][2:]
            if len(can_id) >= 8:
                can_id_type = "Extended"
                extended_can_id = str(int(can_id[0]) + 8) + can_id[1:]
                decimal_value = int(extended_can_id, 16)
            else:
                can_id_type = "Standard"
                decimal_value = int(can_id, 16)

            msg_com = ""
            for key, val in message.comments.items():
                msg_com += val

            # Determine whether the CAN ID is extended or standard
            can_id_type = "Extended" if decimal_value > 2047 else "Standard"

            dlc = message.length  # Extract DLC from the message
            data.append(
                {
                    "CAN ID": can_id,
                    "Decimal": decimal_value,
                    "CAN ID Type": can_id_type,
                    "Message Name": message.name,
                    "DLC": dlc,
                    "Comments": msg_com,
                }
            )
            msg_com = ""
            for signal in message.signals:
                data_type_and_bits = determine_data_type_and_bits(signal)
                endianness = signal.byte_order
                if signal.comments is None:
                    comments += "None"
                else:
                    for key, val in signal.comments.items():
                        comments += val
                scale = signal.scale

                data.append(
                    {
                        "CAN ID": can_id,
                        "Decimal": decimal_value,
                        "CAN ID Type": can_id_type,
                        "Message Name": message.name,
                        "Signal Name": signal.name,
                        "DLC": dlc,
                        "Start Bit": signal.start,
                        "Length": signal.length,
                        "Unit": signal.unit,
                        "Data Type": data_type_and_bits,
                        "Comments": comments,
                        "Offset": signal.offset,
                        "Minimum": signal.minimum,
                        "Maximum": signal.maximum,
                        "Endianness": endianness,
                        "Scale": scale,
                    }
                )
                comments = ""

        df = pd.DataFrame(data)

        parent_path = os.path.dirname(self.selected_file_path)
        dbc_filename = os.path.splitext(os.path.basename(dbc_file))[0]

        excel_file = os.path.join(parent_path, f"{dbc_filename}.xlsx")
        df.to_excel(excel_file, index=False, engine="openpyxl")

        return excel_file

    def process_excel_to_dbc(self, column_mappings):
        """Converts excel to dbc format."""
        df = pd.read_excel(self.selected_file_path)

        user_column_mappings = {
            "CAN ID": None,
            "Decimal": None,
            "CANID Type": None,
            "Message Name": None,
            "DLC": None,
            "Comments": None,
            "Signal Name": None,
            "Start Bit": None,
            "Length": None,
            "Unit": None,
            "Data Type": None,
            "Offset": None,
            "Minimum": None,
            "Maximum": None,
            "Endianness": None,
            "Scale": None,
        }

        for alphabet, data in column_mappings.items():
            if data in user_column_mappings:
                user_column_mappings[data] = int(alphabet)

        try:
            version = 'VERSION "created by BOSON"\n'
            version +="\n"
            version +="\n"

            ns = '''NS_ :\n'''
            ns+="\n"

            bs = 'BS_:\n'
            bs+="\n"

            bu = 'BU_:\n'
            bu+="\n"

            bo_content = ""
            cm_content = ""
            processed_messages = set()  # To keep track of processed messages
            canid = []
            # Process the DataFrame row by row
            for index, row in df.iterrows():
                message_name_col = user_column_mappings["Message Name"]
                can_id_col = user_column_mappings["CAN ID"]
                canid.append(row[can_id_col])
                dlc_col = user_column_mappings["DLC"]
                decimal_col = user_column_mappings["Decimal"]
                if (
                    not pd.isnull(row[message_name_col])
                    and row[message_name_col] not in processed_messages
                ):
                    # Start a new Battery Object
                    can_id = row[can_id_col]
                    bo_content += "\n"
                    bo_content += f"BO_ {row[decimal_col]} {row[message_name_col]}: {row[dlc_col]} Vector__XXX\n"
                    # Add this message to the set of processed messages
                    processed_messages.add(row[message_name_col])

                signal_name_col = user_column_mappings["Signal Name"]
                length_col = user_column_mappings["Length"]
                start_bit_col = user_column_mappings["Start Bit"]
                minimum_col = user_column_mappings["Minimum"]
                maximum_col = user_column_mappings["Maximum"]
                unit_col = user_column_mappings["Unit"]
                data_type_col = user_column_mappings["Data Type"]
                comments_col = user_column_mappings["Comments"]
                endianness_col = user_column_mappings["Endianness"]
                scale_col = user_column_mappings["Scale"]
                offset_col = user_column_mappings["Offset"]

                if not pd.isnull(row[signal_name_col]):
                    length = int(row[length_col])
                    start_bit = int(row[start_bit_col])

                    # Check if 'Minimum' is NaN and assign a default value of 0
                    if pd.isna(row[minimum_col]):
                        min_val = 0
                    else:
                        min_val = int(row[minimum_col])

                    # Check if 'Maximum' is NaN and assign a default value of 0
                    if pd.isna(row[maximum_col]):
                        max_val = 0
                    else:
                        max_val = int(row[maximum_col])

                    unit = row[unit_col]
                    if not pd.isna(row[signal_name_col]):
                        signal_name = row[signal_name_col]
                    comments = row[comments_col]
                    # Check the 'endianness' column to determine the output label
                    if not pd.isna(row[endianness_col]):
                        endianness = row[endianness_col]
                        output_label = "1" if endianness == "little_endian" else "0"

                    # Check if the 'Data Type' column contains 'signed' to determine the factor
                    if (
                        not pd.isnull(row[data_type_col])
                        and "unsigned" in row[data_type_col].strip().lower()
                    ):
                        fact = "+"
                        factor = output_label + fact
                    elif (
                        not pd.isnull(row[data_type_col])
                        and "signed" in row[data_type_col].strip().lower()
                    ):
                        fact = "-"
                        factor = output_label + fact

                    # Replace "nan" with an empty string in the 'unit' field
                    if pd.isna(unit):
                        unit = "NA"
                    sg_content = f' SG_ {signal_name}: {start_bit}|{length}@{factor} ({row[scale_col]},{row[offset_col]}) [{min_val}|{max_val}] "{unit}" Vector__XXX \n'
                    bo_content += sg_content
                    cm_content += "\n"
                    if isinstance(comments, str) and '"neutral"' in comments:
                        comments = comments.replace('"neutral"', '\\"neutral\\"')
                    cm_content += (
                        f'CM_ SG_ {row[decimal_col]} {signal_name} "{comments}";'
                    )
                else:
                    cm_content += f'CM_ BO_ {row[decimal_col]} "{row[comments_col]}";'

            # Combine all sections into the DBC content
            dbc_content = version + ns + bs + bu + bo_content + "\n" + "\n" + cm_content

            excel_dir = os.path.dirname(self.selected_file_path)
            excel_base_name = os.path.splitext(
                os.path.basename(self.selected_file_path)
            )[0]

            # Create the DBC file path by combining the directory and base name with ".dbc" extension
            dbc_file_path = os.path.join(excel_dir, f"{excel_base_name}.dbc")

            # Save the DBC content to the file with the created path
            with open(dbc_file_path, "w") as dbc_file:
                dbc_file.write(dbc_content)

            return dbc_file_path
        except Exception as e:
            self.exception_occurred.emit(e)
