import pandas as pd
def main():
    from pathlib import Path

    # Paths
    script_dir = Path(__file__).resolve().parent
    file_excel = script_dir / "shared" / "04_Aug25-BOQ_SOPs_from_CAD_corners.xlsx"
    output_folder = script_dir / "shared" / "drawingdata"

    # Load Excel file
    df = pd.read_excel(file_excel)
    df.columns = df.columns.str.strip()  # Clean up headers

    # Loop through each unique CIRCUIT REF
    for circuit in df["CIRCUIT REF"].dropna().unique():
        circuit_df = df[df["CIRCUIT REF"] == circuit]

        # Create a safe filename by replacing spaces or slashes
        safe_circuit_name = str(circuit).replace(" ", "_").replace("/", "-")
        output_file = output_folder / f"{safe_circuit_name}.xlsx"

        # Write to Excel
        circuit_df.to_excel(output_file, index=False)

    print("✅ Done! Excel files written to:", output_folder)
