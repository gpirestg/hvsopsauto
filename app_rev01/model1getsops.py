import ezdxf
import math
import csv
from pathlib import Path
import pandas as pd

def main():
    def distance(p1, p2):
        return math.sqrt((p2[0] - p1[0]) ** 2 +
                         (p2[1] - p1[1]) ** 2 +
                         (p2[2] - p1[2]) ** 2)

    def average_center(points):
        x_vals = [pt[0] for pt in points]
        y_vals = [pt[1] for pt in points]
        center_x = sum(x_vals) / len(x_vals)
        center_y = sum(y_vals) / len(y_vals)
        return (round(center_x, 3), round(center_y, 3))

    def process_and_export(file_path, output_csv, output_dxf):
        try:
            doc = ezdxf.readfile(file_path)
            msp = doc.modelspace()

            # New DXF output
            new_doc = ezdxf.new(dxfversion='R2010')
            new_msp = new_doc.modelspace()

            results = []

            for index, entity in enumerate(msp):
                if entity.dxftype() == "3DFACE":
                    points = [entity.dxf.vtx0, entity.dxf.vtx1, entity.dxf.vtx2]
                    if hasattr(entity.dxf, 'vtx3') and entity.dxf.vtx3 != entity.dxf.vtx2:
                        points.append(entity.dxf.vtx3)

                    perimeter = sum(distance(points[i], points[(i + 1) % len(points)]) for i in range(len(points)))

                    # Inside the loop, after perimeter and center have been calculated
                    if 5.9 <= perimeter <= 6.1:
                        center = average_center(points)
                        element_name = f"3DFACE_{index}"

                        # Write each vertex row to CSV
                        for pt in points:
                            results.append([
                                    element_name,
                                    round(pt[0], 3),
                                    round(pt[1], 3),
                                    round(pt[2], 3),
                                    round(perimeter, 3),
                                    round(center[0], 3),  # center_x
                                    round(center[1], 3)   # center_y
                                ])


                        # Add polyline to DXF
                        poly_points = [(pt[0], pt[1]) for pt in points]
                        poly_points.append(poly_points[0])  # close it
                        new_msp.add_lwpolyline(poly_points, close=True)

                        # Add text label to DXF
                        new_msp.add_text(element_name,dxfattribs={
                            'height': 0.25,
                            'layer': "LABELS",
                            'insert':(center[0], center[1])})



            # Write to CSV
            with open(output_csv, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["element", "x", "y", "z", "perimeter", "center_x", "center_y"])
                writer.writerows(results)

            # Save DXF
            new_doc.saveas(output_dxf)

            print(f"Exported {len(results)} coordinates to CSV: {output_csv}")
            print(f"Exported matching faces to DXF: {output_dxf}")

        except Exception as e:
            print(f"Error: {e}")

    #if __name__ == "__main__":
    # Get the current script's directory
    script_dir = Path(__file__).resolve().parent

    #INPUT
    dxf_file = script_dir / "uploads" / "01_shapes_foundation_all_.dxf"
    #OUTPUT
    output_csv = script_dir / "shared" / "01_Aug25-2D_SOPs_From_CAD.csv"
    output_csv_filter = script_dir / "shared" / "01_Aug25-2D_SOPs_From_CAD_F_.csv"
    output_dxf = script_dir / "shared" / "01_Aug25-2D_SOPs_From_CAD.dxf"

    #RUN
    process_and_export(dxf_file, output_csv, output_dxf)

    ###################################################################
    # Convert CSV to Excel
    output_excel = output_csv.with_suffix('.xlsx')
    df = pd.read_csv(output_csv)
    df.to_excel(output_excel, index=False)

    #FILTER data and convert to xlsx
    # Define output Excel path (same name, .xlsx extension)
    output_excel = output_csv_filter.with_suffix('.xlsx')

    # Load CSV
    df = pd.read_csv(output_csv)

    # Set z column to default value
    default_z_value = 81500  # You can change this as needed
    df['z'] = default_z_value

    # Drop duplicates based on center_x and center_y
    filtered_df = df.drop_duplicates(subset=['center_x', 'center_y'])

    # Rename columns
    filtered_df = filtered_df.rename(columns={
        'center_x': 'Easting OS',
        'center_y': 'Northing OS',
        'z': 'Level (mm)'
    })

    # Save to Excel
    filtered_df=filtered_df.drop(columns=['x','y','perimeter'])
    filtered_df.to_excel(output_excel, index=False)

    print(f"Filtered and renamed Excel file saved to: {output_excel}")

def display():
    print("Automation Started")

