import sys
import pdfplumber
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout
)


pdf_path = "Academic Transcript.pdf"

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
        
        # Loop through each page
        for page in pdf.pages:
            text = page.extract_text()
            
            # Split text into lines
            lines = text.split("\n")
            
            # Parse the 'Unit of Study Results' section
            start_parsing = False
            for line in lines:
                if "Year" in line:
                    start_parsing = True
                    continue
                if start_parsing and "Credit points gained" in line:
                    break  # Stop when the relevant section ends
                if start_parsing:
                    extracted_data.append(line)

        # Process the extracted lines
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

        # Convert to DataFrame
        return pd.DataFrame(relevant_data)
    

class WAMCalculatorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Transcript WAM Calculator")
        self.setGeometry(100, 100, 800, 600)
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

        # Results display
        self.wam_label = QLabel("WAM: ")
        self.eihwam_label = QLabel("EIHWAM: ")
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
            self.wam_label.setText("Error: Unable to process file")
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

