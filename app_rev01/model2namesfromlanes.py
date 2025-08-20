import ezdxf
import pandas as pd
from shapely.geometry import LineString, Point
from pathlib import Path
from collections import defaultdict

def main():
    def list_filtered_entities(file_path, excel_path, output_excel_path, output_dxf_path, radius=2):
        try:
            # Load the DXF file
            doc = ezdxf.readfile(file_path)
            msp = doc.modelspace()

            # Read Excel file and extract X, Y coordinates
            df = pd.read_excel(excel_path, usecols=["Easting OS", "Northing OS", "Level (mm)"])
            points = [Point(x, y) for x, y in zip(df["Easting OS"], df["Northing OS"])]

            # Dictionary to store polylines and lines grouped by layer
            layer_lines = {}

            # Iterate through all entities in modelspace
            for entity in msp:
                layer = entity.dxf.layer

                # Check if layer name starts with "Bus_P1_" or "Bus_P2_" (Priority to Bus_P1_)
                #if layer.startswith("Bus_P1_") or layer.startswith("Bus_P2_"):
                entity_type = entity.dxftype()

                # If entity is LWPOLYLINE or LINE, extract vertex points
                if entity_type == "LWPOLYLINE":
                    vertices = [tuple(vertex[:2]) for vertex in entity.get_points()]
                elif entity_type == "LINE":
                    vertices = [(entity.dxf.start.x, entity.dxf.start.y), (entity.dxf.end.x, entity.dxf.end.y)]
                else:
                    continue

                polyline = LineString(vertices)
                if layer not in layer_lines:
                    layer_lines[layer] = []
                layer_lines[layer].append(polyline)
                print(f"Layer: {layer}, Polyline/Line Length: {polyline.length:.2f}")

            # Prioritize Bus_P1_ layers first
            #sorted_layers = sorted(layer_lines.keys(), key=lambda x: (not x.startswith("Bus_P1_"), x))

            # Create a new DXF document for point labels
            text_doc = ezdxf.new()
            text_msp = text_doc.modelspace()

            # Assign points to the nearest line within the specified radius and order them along the line
            assigned_points = set()  # To track used points

            # Dictionary to store point data for Excel export
            line_point_data = {}

            #for layer in sorted_layers:
            for layer in layer_lines:
                line_count = 1  # Reset line numbering for each layer
                for line in layer_lines[layer]:
                    available_points = [p for p in points if (p.x, p.y) not in assigned_points]  # Exclude used points

                    for point in available_points:
                        nearest_distance = line.distance(point)

                        if nearest_distance <= radius:
                            projected_distance = line.project(point)
                            print(projected_distance)
                            short_layer_name = layer.replace("Bus_P1_", "").replace("Bus_P2_", "")
                            line_name = f"{short_layer_name}_L{line_count}"
                            if line_name not in line_point_data:
                                line_point_data[line_name] = []
                            level = df.loc[(df["Easting OS"] == point.x) & (df["Northing OS"] == point.y), "Level (mm)"].values
                            id_value = df.loc[(df["Easting OS"] == point.x) & (df["Northing OS"] == point.y)].values
                            level = level[0] if len(level) > 0 else None
                            id_value = id_value[0] if len(id_value) > 0 else None
                            line_point_data[line_name].append((layer, line_name, point.x, point.y, projected_distance, level, id_value))
                            assigned_points.add((point.x, point.y))  # Mark point as used

                    line_count += 1  # Increment line count for the next line in the same layer

            # Export each set of points to a separate sheet in an Excel file, sorted by projected distance
            with pd.ExcelWriter(output_excel_path) as writer:
                if not line_point_data:  # Ensure at least one sheet exists
                    pd.DataFrame({"Message": ["No points assigned to any lines."]}).to_excel(writer, sheet_name="No Data", index=False)
                else:
                    for line_name, points in line_point_data.items():
                        if points:  # Only write sheets with points
                            points.sort(key=lambda p: p[4])  # Sort by projected distance
                            point_data = [(f"{p[0].replace('Bus_P1_', '').replace('Bus_P2_', '')}_L{p[1][-1]}P{i+1}", p[2], p[3], p[5], p[6]) for i, p in enumerate(points)]
                            line_df = pd.DataFrame(point_data, columns=["Point Name", "X", "Y", "Level (mm)", "Extra"])
                            line_df = line_df.drop(columns=["Extra"])
                            line_df.to_excel(writer, sheet_name=line_name, index=False)

                            # Add text labels to DXF
                            for name, x, y, _, _ in point_data:
                                text_msp.add_text(name, dxfattribs={"height": 0.25, "insert": (x, y, 0)})

            # Save the DXF file with text labels
            text_doc.saveas(output_dxf_path)
            print(f"Points assigned and saved to: {output_excel_path}")
            print(f"DXF file with point labels saved to: {output_dxf_path}")

        except Exception as e:
            print(f"Error: {e}")

    #### Start Model
    # Get the current script's directory
    script_dir = Path(__file__).resolve().parent
    #INPUT DATA
    dxf_file = script_dir / "uploads" / "02_Aug25_busbarlanes.dxf"
    excel_path = script_dir / "shared" / "01_Aug25-2D_SOPs_From_CAD_F_.xlsx"
    #OUTPUT DATA
    output_excel_path = script_dir / "shared" / "02_Aug25-Names_From_2D_SOPs.xlsx"
    output_dxf_path = script_dir / "shared" / "02_Aug25-Names_From_2D_SOPs.dxf"
    #RUN SCRIPT
    list_filtered_entities(dxf_file, excel_path, output_excel_path, output_dxf_path, radius=2)

    ###########################################################################################
    ###########################################################################################
    ### Creat and sort optional spread sheets formats
    output_combined_path = script_dir / "shared" / "02_Aug25-Names_From_2D_SOPs_Comb_1.xlsx"

    # Load all sheet names
    xls = pd.ExcelFile(output_excel_path)
    sheet_names = xls.sheet_names

    # Group sheets by prefix
    prefix_groups = defaultdict(dict)

    for sheet in sheet_names:
        if sheet.endswith(('_L1', '_L2', '_L3')):
            prefix, suffix = sheet.rsplit('_', 1)
            prefix_groups[prefix][suffix] = sheet

    # Rename map
    rename_map = {
        "Point Name": "FOUNDATION REF",
        "X": "EASTING (mm)",
        "Y": "NORTHING (mm)",
        "Level (mm)": "FOUNDATION (T.O.C)"
    }

    # Sort and combine
    with pd.ExcelWriter(output_combined_path, engine='openpyxl') as writer:
        for prefix, parts in prefix_groups.items():
            combined_df = pd.DataFrame()
            for suffix in ['L1', 'L2', 'L3']:
                sheetname = parts.get(suffix)
                if sheetname:
                    df = pd.read_excel(xls, sheet_name=sheetname)
                    df = df.rename(columns=rename_map)
                    df["CIRCUIT REF"] = prefix  # ➕ add column here
                    cols = ["CIRCUIT REF"] + [col for col in df.columns if col != "CIRCUIT REF"]
                    df = df[cols]
                    combined_df = pd.concat([combined_df, df], ignore_index=True)

            combined_df.to_excel(writer, sheet_name=prefix, index=False)

    print(f"✅ Combined sheets saved to {output_combined_path}")

    # Combine in one sheet
    output_combined_path2 = script_dir / "shared" / "02_Aug25-Names_From_2D_SOPs_Comb_2.xlsx"
    all_sheets = pd.read_excel(output_combined_path, sheet_name=None)
    combined_df = pd.concat(all_sheets.values(), ignore_index=True)
    combined_df.to_excel(output_combined_path2, index=False)
    print(f"✅ Combined sheet saved to: {output_combined_path2}")
