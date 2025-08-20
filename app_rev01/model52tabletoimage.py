import pandas as pd
import matplotlib.pyplot as plt

def main():
    # Define file paths
    from pathlib import Path
    script_dir = Path(__file__).resolve().parent
    input_path = script_dir / "shared" / "drawingdata"
    
    # Get all Excel files in the directory
    excel_path = list(input_path.glob("*.xlsx"))

    for file in excel_path:
        print(f"Processing {file.name} please wait!")

        # Load Excel file
        df = pd.read_excel(file)

        # Create figure and axis
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.axis('tight')
        ax.axis('off')

        # Create table
        table = ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')

        # Format table
        for (i, j), cell in table.get_celld().items():
            cell.set_fontsize(10)
            cell.set_facecolor('white')
            if i == 0:
                cell.set_text_props(weight='bold', fontname='Arial')
                cell.set_facecolor('#D3D3D3')
            else:
                cell.set_text_props(fontname='Arial')

        table.auto_set_column_width(col=list(range(len(df.columns))))

        # Generate image filename
        image_filename = file.stem + ".png"
        output_path = script_dir / "shared" / "drawingdata" / image_filename

        # Save image
        plt.savefig(output_path, bbox_inches="tight", dpi=300)
        plt.close()

        print(f"✅ Table image saved at: {output_path}")


    import zipfile
    # Create the zip in the current script directory
    zip_path = script_dir / "drawingdata.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in input_path.rglob("*"):
            if file_path.is_file():
                zf.write(file_path, arcname=file_path.relative_to(input_path))

    # Create downloads folder if it doesn't exist
    downloads_folder = script_dir / "downloads"
    downloads_folder.mkdir(exist_ok=True)

    # Move the zip to the downloads folder
    final_zip_path = downloads_folder / zip_path.name
    zip_path.replace(final_zip_path)

    print(f"ZIP created and moved to {final_zip_path}")
