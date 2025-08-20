import pandas as pd
### THIS SCRIPT #########################################
### MATCHES step 1 data with client BOQ schedulled data #
### Adds design foundations types and sizes             #
#########################################################
def main():
    # File paths
    from pathlib import Path
    script_dir = Path(__file__).resolve().parent
    file1 = script_dir / "uploads" / "CAAR-BOQ.xlsx"
    file2 = script_dir / "shared" / "new_vs_old.xlsx"
    output_file = script_dir / "shared" / "04_Aug25-BOQ_SOPs_from_CAD.xlsx"

    # Load both Excel files
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Rename file2's column temporarily to match df1 for merging
    df2_temp = df2.rename(columns={"FOUNDATION REF": "NEW_FOUNDATION REF"})
    df2_temp = df2_temp.rename(columns={"F2_FOUNDATION REF": "FOUNDATION REF"})

    # Perform inner merge on both CIRCUIT REF and FOUNDATION REF
    merged_df = pd.merge(
        df1, df2_temp,
        on=["CIRCUIT REF", "FOUNDATION REF"],
        how="inner",
        suffixes=('_BOQ', '_CAD')  # So shared columns from each file are distinguishable
    )

    # OPTIONAL: rename FOUNDATION REF back to F2_... if needed
    #merged_df.rename(columns={"FOUNDATION REF": "F2_FOUNDATION REF"}, inplace=True)

    #drop cells and rename
    merged_df.columns = merged_df.columns.str.strip()
    merged_df=merged_df.drop(["EASTING (mm)","NORTHING (mm)","FOUNDATION (T.O.C)_BOQ","F2_FOUNDATION (T.O.C)","Match Status"],axis=1)
    merged_df = merged_df.rename(columns={
        "FOUNDATION (T.O.C)_CAD": "FOUNDATION (T.O.C)",
        "F2_EASTING (mm)": "EASTING (mm)",
        "F2_NORTHING (mm)": "NORTHING (mm)"
    })

    #####################################################################################################################################
    # Associate with foundation types template
    # Load the foundation type lookup table
    f_type_file = script_dir / "uploads" / "04_Design_Foundations_Type.xlsx"
    df_f_type = pd.read_excel(f_type_file)
    df_f_type.columns = df_f_type.columns.str.strip()  # Just in case

    # Optional: check it's using the expected headers
    # print(df_f_type.columns)

    # Merge with suffixes to handle duplicate column names
    merged_df = pd.merge(
        merged_df,
        df_f_type,
        on="EQUIPMENT",
        how="left",
        suffixes=('', '_new')  # This will make the new column FOUNDATION TYPE_new
    )

    # Replace the old FOUNDATION TYPE with the new one
    merged_df["FOUNDATION TYPE"] = merged_df["FOUNDATION TYPE_new"]

    # Drop the extra column
    merged_df.drop(columns=["FOUNDATION TYPE_new"], inplace=True)


    #####################################################################################################################################
    # Save to Excel
    merged_df.to_excel(output_file, index=False)

    print(f"✅ {len(merged_df)} matching rows written to:\n{output_file}")

    #####################################################################################################################################
    # Create CAD QA DXF
    import ezdxf
    from pathlib import Path

    # Path to save your DXF file
    output_dxf = script_dir / "shared" / "04_CAD_QA_Step2.dxf"

    # Create new DXF
    doc = ezdxf.new(dxfversion='R2010')
    msp = doc.modelspace()

    # Loop through merged_df rows
    for _, row in merged_df.iterrows():
        x = row["EASTING (mm)"]
        y = row["NORTHING (mm)"]
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

    # Save DXF
    doc.saveas(output_dxf)
    print(f"✅ DXF file saved to: {output_dxf}")
