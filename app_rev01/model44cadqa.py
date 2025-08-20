#####################################################################################################################################
# Create CAD QA DXF
import ezdxf
from pathlib import Path
import pandas as pd

def main():
    # Path to save your DXF file
    script_dir = Path(__file__).resolve().parent
    file_excel = script_dir / "shared" / "04_Aug25-BOQ_SOPs_from_CAD_corners.xlsx"
    output_dxf = script_dir / "shared" / "04_CAD_QA_Final.dxf"

    #read excel
    df = pd.read_excel(file_excel)

    # Create new DXF
    doc = ezdxf.new(dxfversion='R2010')
    msp = doc.modelspace()

    # Loop through merged_df rows
    for _, row in df.iterrows():
        x = row["1 EASTING (mm)"]
        y = row["1 NORTHING (mm)"]
        z = 0

        # Build multiline text string
        text_lines = [
            f"CIRCUIT REF: {row.get('CIRCUIT REF', '')}",
            f"FOUNDATION REF: {row.get('FOUNDATION REF', '')}",
            f"NEW FOUNDATION REF: {row.get('NEW_FOUNDATION REF', '')}",
            f"DUCT REQUIRED: {row.get('DUCT REQUIRED', '')}",
            f"RATING (Kv): {row.get('RATING (Kv)', '')}",
            f"PHASE: {row.get('PHASE', '')}",
            f"EQUIPMENT: {row.get('EQUIPMENT', '')}",
            f"FOUNDATION TYPE: {row.get('FOUNDATION TYPE', '')}",
        ]
        full_text = "\n".join(text_lines)

        # Add MTEXT with corrected attribute
        msp.add_mtext(full_text, dxfattribs={
            "insert": (x, y, z),
            "char_height": 0.1,  # Correct attribute for text height
            "layer": "ANNOTATIONS"
        })
        ####################################################################
        ### ADD circles to SOPs
        # Function to add circle if coordinates exist and are valid
        def add_optional_circle(df_row, e_col, n_col, layer="ANNOTATIONS"):
            easting = df_row.get(e_col)
            northing = df_row.get(n_col)
            if pd.notna(easting) and pd.notna(northing):
                msp.add_circle(
                    center=(easting, northing, 0),
                    radius=0.05,  # 50mm diameter
                    dxfattribs={"layer": layer}
                )

        # Add additional circles for 2/3/4 EASTING/NORTHING if available
        add_optional_circle(row, "1 EASTING (mm)", "1 NORTHING (mm)")
        add_optional_circle(row, "2 EASTING (mm)", "2 NORTHING (mm)")
        add_optional_circle(row, "3 EASTING (mm)", "3 NORTHING (mm)")
        add_optional_circle(row, "4 EASTING (mm)", "4 NORTHING (mm)")
        ####################################################################
        # Collect all valid (x, y) points in order: 1 → 2 → 3 → 4
        line_points = []

        # Main point (1)
        x1 = row.get("1 EASTING (mm)")
        y1 = row.get("1 NORTHING (mm)")
        if pd.notna(x1) and pd.notna(y1):
            line_points.append((x1, y1))

        # Optional extra points (2 to 4)
        for i in range(2, 5):
            xi = row.get(f"{i} EASTING (mm)")
            yi = row.get(f"{i} NORTHING (mm)")
            if pd.notna(xi) and pd.notna(yi):
                line_points.append((xi, yi))

        # Draw polyline if we have at least 2 points
        if len(line_points) >= 2:
            # Append first point to end to close loop
            line_points.append(line_points[0])
            # Convert to 3D tuples (x, y, z=0)
            line_points_3d = [(pt[0], pt[1], 0) for pt in line_points]

            msp.add_lwpolyline(line_points_3d, dxfattribs={"layer": "ANNOTATIONS"})


    # Save DXF
    doc.saveas(output_dxf)
    print(f"✅ DXF file saved to: {output_dxf}")