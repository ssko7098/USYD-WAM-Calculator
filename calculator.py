import pdfplumber
import pandas as pd
import numpy as np
from tkinter import Tk
from tkinter.filedialog import askopenfilename

pdf_path = "Academic Transcript.pdf"

def calculate_weighting(unit_code, unit_name):
    # Check if the unit is a thesis unit
    if "thesis" in unit_name.lower():
        return 8
    # Calculate weighting based on unit code
    first_digit = int(unit_code[4])  # Extract the first digit after 'XXXX'
    if first_digit == 1:
        return 0
    elif first_digit == 2:
        return 2
    elif first_digit == 3:
        return 3
    elif first_digit >= 4:
        return 4
    else:
        return 0  # Default case if needed

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
            if len(parts) > 5 and parts[-1].isdigit() and int(parts[-1]) > 0:  # Ensure the line contains enough elements to represent data
                relevant_data.append({
                    "Year": parts[0],
                    "Session": parts[1],
                    "Unit Code": parts[2],
                    "Unit Name": " ".join(parts[3:-3]),  # Combine all parts of the unit name
                    "Mark": parts[-3],
                    "Grade": parts[-2],
                    "Credit Points": parts[-1],
                })

        # Convert to DataFrame
        df = pd.DataFrame(relevant_data)

    return df

def main():
    Tk().withdraw()
    print("Please choose your transcript file.")
    
    
    # Open file picker
    file_path = askopenfilename(
        title="Select Transcript File",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not file_path:
        print("No file selected. Exiting program.")
        return
    
    try:
        df = extract_data(file_path)
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

        WAM = round(CP_vector.dot(M_vector)/sum(CP), 1)
        print(f"WAM: {WAM}")

        EIHWAM =  round(np.sum(W_vector * CP_vector * M_vector) / W_vector.dot(CP_vector), 1)
        print(f"EIHWAM: {EIHWAM}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

