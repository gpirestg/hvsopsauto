import pandas as pd
import math
import os
from pathlib import Path

### THIS SCRIPT ###################################
### Generates 4 coorners SOPs and sorts dataframe #
###################################################
def main():
    def corner_sops(EASTING_OS, NORTHING_OS, WIDTH):
        WIDTH=(WIDTH/1000)/2
        angle = math.radians(-127)  # Convert degrees to radians

        # Calculate the corners relative to the center
        corners = [
            (WIDTH * math.cos(angle) - WIDTH * math.sin(angle), WIDTH * math.sin(angle) + WIDTH * math.cos(angle)),  # Top-right
            (WIDTH * math.cos(angle) + WIDTH * math.sin(angle), WIDTH * math.sin(angle) - WIDTH * math.cos(angle)),  # Bottom-right
            (-WIDTH * math.cos(angle) + WIDTH * math.sin(angle), -WIDTH * math.sin(angle) - WIDTH * math.cos(angle)),  # Bottom-left
            (-WIDTH * math.cos(angle) - WIDTH * math.sin(angle), -WIDTH * math.sin(angle) + WIDTH * math.cos(angle))  # Top-left
        ]

        # Convert to absolute coordinates by adding the center point
        absolute_corners = [(round(x + EASTING_OS,3), round(y + NORTHING_OS,3)) for x, y in corners]

        # returns data
        return absolute_corners

    ####START
    # File paths
    script_dir = Path(__file__).resolve().parent
    folder_path = script_dir / "shared" / "04_Aug25-BOQ_SOPs_from_CAD.xlsx"
    output_path = script_dir / "shared" / "04_Aug25-BOQ_SOPs_from_CAD_corners.xlsx"

    # Get all Excel files in the directory
    #excel_path = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]


    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(folder_path)

        # Ensure the column names match what we expect
        required_columns = ["EASTING (mm)", "NORTHING (mm)", "WIDTH (mm)"]
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Required columns not found in the Excel file.")

        # New columns for corner coordinates
        corner_labels = ['1', '2', '3', '4']
        for i, label in enumerate(corner_labels, 1):
            df[f'{label} EASTING (mm)'] = None
            df[f'{label} NORTHING (mm)'] = None

        # Iterate through each row in the DataFrame
        for index, row in df.iterrows():
            EASTING_OS = row['EASTING (mm)']
            NORTHING_OS = row['NORTHING (mm)']
            WIDTH = row['WIDTH (mm)']

            # Process the data with the corner_sops function
            corners = corner_sops(EASTING_OS, NORTHING_OS, WIDTH)

            # Add corner data to the corresponding new columns
            for i, corner in enumerate(corners):
                df.at[index, f'{i+1} EASTING (mm)'] = corner[0]
                df.at[index, f'{i+1} NORTHING (mm)'] = corner[1]

        # move TOC col to the be the last column
        df = df[[col for col in df.columns if col != 'FOUNDATION (T.O.C)'] + ['FOUNDATION (T.O.C)']]

        # drop columns
        df = df.drop(columns=["EASTING (mm)", "NORTHING (mm)","WIDTH (mm)","DESCRPTION"])

        # move columns
        col = df.pop("NEW_FOUNDATION REF")
        df.insert(2, "NEW_FOUNDATION REF", col)

        # Save the modified DataFrame to a new Excel file
        df.to_excel(output_path, index=False)
        print(f"New file with corner data created: {output_path}")

    except Exception as e:
        print(f"An error occurred: {e}")
