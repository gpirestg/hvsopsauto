import csv
import ezdxf
from pathlib import Path
import math
import pandas as pd

def main():
    # Global list to collect all matched point data
    matched_points = []

    def load_csv_centers(csv_filename):
        centers = []
        with open(csv_filename, mode='r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                x = float(row['center_x'])
                y = float(row['center_y'])
                centers.append((x, y))
        return centers

    def is_within_radius(pt1, pt2, radius=2.0):
        return math.dist(pt1, pt2) <= radius

    def filter_text_entities(input_dxf, centers, output_dxf):
        doc = ezdxf.readfile(input_dxf)
        msp = doc.modelspace()
        Z = 81500  # Constant level

        # Create new DXF output
        new_doc = ezdxf.new(dxfversion='R2010')
        new_msp = new_doc.modelspace()

        count = 0
        for entity in msp:
            if entity.dxftype() == "MTEXT":
                insert = entity.dxf.insert
                text_point = (insert[0], insert[1])
                text_content = entity.text

                # Filter for text starting with letter F
                if text_content.startswith("F"):
                    print(f"Text Content: {text_content}")
                    # Check if this text is within radius of any center
                    # Find the first matching center within radius
                    matching_center = next((center for center in centers if is_within_radius(text_point, center)), None)

                    if matching_center:
                        print("Found TXT inside radius!")
                        print("Insert", insert)
                        print("Center", matching_center)
                        new_msp.add_text(entity.dxf.text, dxfattribs={
                            'insert': matching_center,
                            'layer': entity.dxf.layer,
                            'height': 0.35
                        })

                        #store for xlxs data
                        store_matched_point(entity.dxf.text, matching_center, Z)

                        count += 1



        new_doc.saveas(output_dxf)
        print(f"Exported {count} matching TEXT entities to: {output_dxf}")


    def store_matched_point(text, center, level):
        """Store matched point data in a global list."""
        x, y = center
        matched_points.append({
            "Point Name": text,
            "X": x,
            "Y": y,
            "Level (mm)": level
        })

    def export_to_excel(filename="output.xlsx"):
        """Export matched points to an Excel file."""
        df = pd.DataFrame(matched_points)
        df.to_excel(filename, index=False)
        print(f"Saved {len(matched_points)} points to {filename}")

    ##############################################################################################
    #RUN APP
    #if __name__ == "__main__":
    script_dir = Path(__file__).resolve().parent
    #INPUT
    #csv_file = script_dir / "input" / "CAD_foundations_sop.csv"
    #file CSV below may require to be filtered by unique elements in the x & y axis
    csv_file = script_dir / "shared" / "01_Aug25-2D_SOPs_From_CAD.csv"
    dxf_input = script_dir / "uploads" / "03_Drawing_All_Text_Export.dxf"
    #OUTPUT
    dxf_output = script_dir / "shared" / "03_Aug25-Associated_Foundation_Name.dxf"
    excel_output = script_dir / "shared" / "03_Aug25-Associated_Foundation_Name_A.xlsx"

    centers = load_csv_centers(csv_file)
    filter_text_entities(dxf_input, centers, dxf_output)
    export_to_excel(excel_output)

    ##############################################################################################
    # Rename Excel
    excel_output2 = script_dir / "shared" / "03_Aug25-Associated_Foundation_Name_B.xlsx"

    # Rename map
    rename_map = {
        "Point Name":"FOUNDATION REF",
        "X": "EASTING (mm)",
        "Y": "NORTHING (mm)",
        "Level (mm)": "FOUNDATION (T.O.C)"
    }

    # Load and rename all sheets
    xls = pd.ExcelFile(excel_output)

    with pd.ExcelWriter(excel_output2, engine='openpyxl') as writer:
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df = df.rename(columns=rename_map)
            df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"✅ Renamed sheets saved to {excel_output2}")

