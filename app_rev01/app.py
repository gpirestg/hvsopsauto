from pathlib import Path
import streamlit as st
import time
from datetime import datetime

st.set_page_config(page_title="HVDC Foundation Setting Out Automation", layout="wide")
st.title("HVDC Foundation Setting Out Automattion")
st.subheader("Please upload required files below!")

# Save location (absolute, next to this script)
script_dir = Path(__file__).resolve().parent
UPLOAD_DIR = script_dir / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Session state defaults
if "automation_status" not in st.session_state:
    st.session_state["automation_status"] = "idle"  # idle | running | completed
if "automation_completed_at" not in st.session_state:
    st.session_state["automation_completed_at"] = None

# Reusable uploader
def labelled_uploader(label, file_types, key):
    col_label, col_upload, col_empty = st.columns([3, 2, 3])
    with col_label:
        st.subheader(label)
    with col_upload:
        file = st.file_uploader(
            "",
            type=file_types,
            accept_multiple_files=False,
            key=key,
            label_visibility="collapsed"
        )
    if file:
        save_path = UPLOAD_DIR / file.name
        save_path.write_bytes(file.getbuffer())
        st.session_state[f"{key}_saved_path"] = str(save_path)
        st.success(f"Uploaded: {file.name}\nSaved to: {save_path}")

# Upload buttons
labelled_uploader("Foundations Geometry", ["dxf"], "ga_dxf")
labelled_uploader("Busbar Runs String Lines Template", ["dxf"], "busbar_dxf")
labelled_uploader("Drawing Foundations ID Text", ["dxf"], "found_id_dxf")
labelled_uploader("Design Foundations Type", ["xlsx"], "found_type_sheet")
labelled_uploader("BoQ Spreadsheet", ["xlsx"], "boq_sheet")

# Require all 5 uploads before enabling Run
required = {
    "ga_dxf": "Foundations Geometry",
    "busbar_dxf": "Busbar Runs String Lines Template",
    "found_id_dxf": "Drawing Foundations ID Text",
    "found_type_sheet": "Design Foundations Type",
    "boq_sheet": "BoQ Spreadsheet",
}
missing = [label for k, label in required.items() if f"{k}_saved_path" not in st.session_state]
ready = len(missing) == 0

if st.session_state["automation_status"] == "idle":
    if not ready:
        st.info("Waiting for: " + ", ".join(missing))
    else:
        st.success("All required files uploaded.")

    run_clicked = st.button(
        "Run Automation",
        type="primary",
        disabled=not ready,
    )

    if run_clicked and ready:
        st.session_state["automation_status"] = "running"
        st.rerun()

elif st.session_state["automation_status"] == "running":
    with st.spinner("Running automation..."):
        # Collect paths to pass to your pipeline
        input_paths = {k: st.session_state.get(f"{k}_saved_path") for k in required.keys()}
        # TODO: call your actual processing function here, e.g.:
        log_area = st.empty()
        if "logs" not in st.session_state:
            st.session_state.logs = []
        def log(msg):
            st.session_state.logs.append(msg)
            log_area.code("\n".join(st.session_state.logs), language="text")
        log("Starting automation...")
        #####################################################################
        # Step 1 Automation
        log("Step 1/5: Calculating foundations SOPs from geometry file...")
        import model1getsops
        model1getsops.main()
        time.sleep(1)
        #####################################################################
        # Step 2 Automation
        log("Step 2/5: Generate foundations names from CAD busbar lanes...")
        import model2namesfromlanes
        model2namesfromlanes.main()
        time.sleep(1)
        #####################################################################
        # Step 3 Automation
        log("Step 3/5: Associate foundations CAD id text with SOPs info...")
        import model3associatenames
        model3associatenames.main()
        time.sleep(1)
        #####################################################################
        # Step 4 Automation
        log("Step 4/5: Checking data consistency between CAD files...")
        import model41combineboqcad
        model41combineboqcad.main()
        log("Step 4/5: Matching CAD busbar lanes SOPs with excel BoQ info...")
        time.sleep(1)
        log("Step 4/5: Assigning foundation type to HV equipment as per design specs...")
        import model42combineboqcad
        model42combineboqcad.main()
        time.sleep(0.5)
        log("Step 4/5: Generating foundations 4 corners SOPs info...")
        import model43combineboqcad
        model43combineboqcad.main()
        time.sleep(0.5)
        log("Step 4/5: Generating CAD-QA control files...")
        import model44cadqa
        model44cadqa.main()
        time.sleep(0.5)
        log("Step 5/5: Generating busbar lanes tables...")
        import model51sortbybusbarlane
        model51sortbybusbarlane.main()
        time.sleep(0.5)
        log("Step 5/5: Generating drawing tables png files...")
        import model52tabletoimage
        model52tabletoimage.main()

        # copy relevant files to download folder
        script_dir = Path(__file__).resolve().parent
        downloads_path = script_dir / "downloads"
        file_1 = script_dir / "shared" / "04_Aug25-BOQ_SOPs_from_CAD_corners.xlsx"
        file_2 = script_dir / "shared" / "04_CAD_QA_Final.dxf"
        import shutil
        # Copy both files into the downloads folder
        shutil.copy(file_1, downloads_path / file_1.name)
        shutil.copy(file_2, downloads_path / file_2.name)
        
        log("Process completed!")
        time.sleep(3)

    #########################################################################################
    # End of automation
    st.session_state["automation_status"] = "completed"
    st.session_state["automation_completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.logs = []
    st.rerun()

elif st.session_state["automation_status"] == "completed":
    st.success(f"Automation completed at {st.session_state['automation_completed_at']}.")

    # show download folder files
    downloads_path = script_dir / "downloads"
    files_in_downloads = sorted([f for f in downloads_path.iterdir() if f.is_file()])

    if files_in_downloads:
        # Radio selector
        selected_file = st.radio(
            "Available output files:",
            [f.name for f in files_in_downloads]
        )

        # Download button
        selected_path = downloads_path / selected_file
        with open(selected_path, "rb") as f:
            data = f.read()

        st.download_button(
            label=f"Download {selected_file}",
            data=data,
            file_name=selected_file
        )
    else:
        st.info("No output files available in the downloads folder.")


    # Optional: show where files came from
    with st.expander("Show input file paths"):
        for k, label in required.items():
            st.write(f"• {label}: {st.session_state.get(f'{k}_saved_path')}")

    col_a, col_b = st.columns([1, 3])
    with col_a:
        restart = st.button("Restart Automation")
    if restart:
        st.session_state["automation_status"] = "idle"
        st.rerun()
