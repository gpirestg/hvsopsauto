import pandas as pd
import os
from pathlib import Path

### THIS SCRIPT ################################################
### MATCHES Naming from string lines automation with data from #
### client data extracted from CAD files (SOPs and Text)       # 
################################################################

def main():
    # Clear console
    clear = lambda: os.system('cls')
    clear()
    
    # File paths
    script_dir = Path(__file__).resolve().parent
    #INPUT
    file1 = script_dir / "shared" / "02_Aug25-Names_From_2D_SOPs_Comb_2.xlsx"
    file2 = script_dir / "shared" / "03_Aug25-Associated_Foundation_Name_B.xlsx"
    #OUTPUT
    output_file = script_dir / "shared" / "new_vs_old.xlsx"

    # Load Excel files
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Convert coordinates to float and round
    df1["EASTING (mm)"] = pd.to_numeric(df1["EASTING (mm)"], errors="coerce").round(2)
    df1["NORTHING (mm)"] = pd.to_numeric(df1["NORTHING (mm)"], errors="coerce").round(2)
    df2["EASTING (mm)"] = pd.to_numeric(df2["EASTING (mm)"], errors="coerce").round(2)
    df2["NORTHING (mm)"] = pd.to_numeric(df2["NORTHING (mm)"], errors="coerce").round(2)

    # Prepare output
    output_rows = []
    tolerance = 0
    match_count = 0
    no_match_count = 0

    df2_used_indices = set()  # to track which rows in df2 have already matched

    for _, row1 in df1.iterrows():
        e1 = row1["EASTING (mm)"]
        n1 = row1["NORTHING (mm)"]
        matched = False

        for index2, row2 in df2.iterrows():
            if index2 in df2_used_indices:
                continue  # Skip already matched rows

            e2 = row2["EASTING (mm)"]
            n2 = row2["NORTHING (mm)"]

            if abs(e1 - e2) <= tolerance and abs(n1 - n2) <= tolerance:
                # Combine all columns from both rows
                #combined_data = {f"file1_{k}": v for k, v in row1.items()}
                combined_data = {k: v for k, v in row1.items()}
                combined_data.update({f"F2_{k}": v for k, v in row2.items()})
                combined_data["Match Status"] = "MATCHED"
                output_rows.append(combined_data)

                df2_used_indices.add(index2)
                match_count += 1
                matched = True
                break

        if not matched:
            #combined_data = {f"file1_{k}": v for k, v in row1.items()}
            combined_data = {k: v for k, v in row1.items()}
            # Add empty file2 columns
            for col in df2.columns:
                combined_data[f"F2_{col}"] = None
            combined_data["Match Status"] = "NO MATCH"
            output_rows.append(combined_data)
            no_match_count += 1

    # Convert to DataFrame and export
    df_output = pd.DataFrame(output_rows)
    df_output.to_excel(output_file, index=False)

    # Summary
    print(f"\n✅ Output saved to {output_file}")
    print(f"\n🔍 Matching Summary:")
    print(f"  Matches found     : {match_count}")
    print(f"  No matches found  : {no_match_count}")
    print("\n🎯 Process Complete.")

