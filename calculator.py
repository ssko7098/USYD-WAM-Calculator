import sys
import pdfplumber
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


def calculate_weighting(unit_code, unit_name):
    if "thesis" in unit_name.lower():
        return 8
    first_digit = int(unit_code[4])
    if first_digit == 1:
        return 0
    elif first_digit == 2:
        return 2
    elif first_digit == 3:
        return 3
    elif first_digit >= 4:
        return 4
    else:
        return 0


def extract_data(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        extracted_data = []
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split("\n")
            start_parsing = False
            for line in lines:
                if "Year" in line:
                    start_parsing = True
                    continue
                if start_parsing and "Credit points gained" in line:
                    break
                if start_parsing:
                    extracted_data.append(line)

        relevant_data = []
        for line in extracted_data:
            parts = line.split()
            if len(parts) > 5 and parts[-1].isdigit() and int(parts[-1]) > 0:
                relevant_data.append({
                    "Year": parts[0],
                    "Session": parts[1],
                    "Unit Code": parts[2],
                    "Unit Name": " ".join(parts[3:-3]),
                    "Mark": parts[-3],
                    "Grade": parts[-2],
                    "Credit Points": parts[-1],
                })

        return pd.DataFrame(relevant_data)


class WAMCalculatorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Transcript WAM Calculator")
        self.setGeometry(100, 100, 900, 600)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2E3440;
            }
            QLabel {
                color: #D8DEE9;
                font-size: 16px;
            }
            QPushButton {
                background-color: #5E81AC;
                color: #ECEFF4;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #81A1C1;
            }
            QTableWidget {
                background-color: #3B4252;
                color: #ECEFF4;
                gridline-color: #4C566A;
                font-size: 14px;
                border: 1px solid #4C566A;
            }
            QHeaderView::section {
                background-color: #5E81AC;
                color: #ECEFF4;
                font-size: 14px;
                padding: 4px;
                border: none;
            }
        """)
        self.init_ui()

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout()

        # Buttons
        self.load_button = QPushButton("Load Transcript")
        self.load_button.clicked.connect(self.load_transcript)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Year", "Session", "Unit Code", "Unit Name", "Mark", "Grade", "Credit Points"
        ])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setDefaultSectionSize(150)
        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.setAlternatingRowColors(True)

        # Results display
        self.wam_label = QLabel("WAM: --")
        self.wam_label.setFont(QFont("Arial", 18, QFont.Bold))
        self.wam_label.setAlignment(Qt.AlignCenter)

        self.eihwam_label = QLabel("EIHWAM: --")
        self.eihwam_label.setFont(QFont("Arial", 18, QFont.Bold))
        self.eihwam_label.setAlignment(Qt.AlignCenter)

        results_layout = QHBoxLayout()
        results_layout.addWidget(self.wam_label)
        results_layout.addWidget(self.eihwam_label)

        # Add widgets to layout
        main_layout.addWidget(self.load_button)
        main_layout.addWidget(self.table)
        main_layout.addLayout(results_layout)

        # Set main layout
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def load_transcript(self):
        # Open file picker
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Transcript File", "", "PDF Files (*.pdf)"
        )
        if not file_path:
            return

        # Process the selected file
        try:
            df = extract_data(file_path)
            self.populate_table(df)
            self.calculate_wam(df)
        except Exception as e:
            self.wam_label.setText("Error")
            self.eihwam_label.setText("Error")
            print(f"Error: {e}")

    def populate_table(self, df):
        self.table.setRowCount(len(df))
        for i, row in df.iterrows():
            self.table.setItem(i, 0, QTableWidgetItem(row['Year']))
            self.table.setItem(i, 1, QTableWidgetItem(row['Session']))
            self.table.setItem(i, 2, QTableWidgetItem(row['Unit Code']))
            self.table.setItem(i, 3, QTableWidgetItem(row['Unit Name']))
            self.table.setItem(i, 4, QTableWidgetItem(row['Mark']))
            self.table.setItem(i, 5, QTableWidgetItem(row['Grade']))
            self.table.setItem(i, 6, QTableWidgetItem(row['Credit Points']))

    def calculate_wam(self, df):
        W = []
        CP = []
        M = []

        for _, row in df.iterrows():
            unit_code = row['Unit Code']
            unit_name = row['Unit Name']
            mark = row['Mark']
            credit_points = row['Credit Points']

            W.append(calculate_weighting(unit_code, unit_name))
            CP.append(int(credit_points))
            M.append(float(mark))

        W_vector = pd.Series(W)
        CP_vector = pd.Series(CP)
        M_vector = pd.Series(M)

        wam = round(CP_vector.dot(M_vector) / sum(CP), 1)
        eihwam = round(np.sum(W_vector * CP_vector * M_vector) / W_vector.dot(CP_vector), 1)

        self.wam_label.setText(f"WAM: {wam}")
        self.eihwam_label.setText(f"EIHWAM: {eihwam}")


def main():
    app = QApplication(sys.argv)
    window = WAMCalculatorApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
