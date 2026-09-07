import streamlit as st
import backend

st.set_page_config(page_title="Style Master Extractor", layout="wide")

st.title("Order Style Master Extractor")
st.write("Upload your Order Details and Master File, configure input parameters, and extract formatted style blocks.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Order Details File")
    order_file = st.file_uploader("Upload Order Details (.xlsx)", type=["xlsx"], key="order_uploader")
    
    st.markdown("**Order File Configuration**")
    order_col_a, order_col_b = st.columns(2)
    with order_col_a:
        style_col_input = st.text_input("Style Column (Letter or Number)", value="B")
    with order_col_b:
        start_row_input = st.number_input("Data Start Row", min_value=1, value=6, step=1)

with col2:
    st.subheader("2. ANRO Master File")
    master_file = st.file_uploader("Upload Master File (.xlsx)", type=["xlsx"], key="master_uploader")

    selected_sheet = None
    if master_file is not None:
        try:
            available_sheets = backend.get_sheet_names_from_bytes(master_file.getvalue())
            default_index = 0
            # Pre-select known tabs if present
            for preferred in ["14K SI", "14K I1", "14K LGD"]:
                if preferred in available_sheets:
                    default_index = available_sheets.index(preferred)
                    break
            
            selected_sheet = st.selectbox("Select Master Tab to Search", available_sheets, index=default_index)
        except Exception as e:
            st.error(f"Error reading master file sheets: {e}")

st.markdown("---")

# Execution trigger
can_process = order_file is not None and master_file is not None and selected_sheet is not None

if st.button("Process and Extract Styles", type="primary", disabled=not can_process):
    with st.spinner("Extracting master styles and formatting images..."):
        try:
            output_bytes, found, missing = backend.process_workbook_extraction(
                order_bytes=order_file.getvalue(),
                master_bytes=master_file.getvalue(),
                master_sheet_name=selected_sheet,
                order_start_row=int(start_row_input),
                order_style_col=style_col_input.strip()
            )

            st.success("Extraction complete!")

            # Summary metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Ordered Styles", len(found) + len(missing))
            m2.metric("Styles Found & Extracted", len(found))
            m3.metric("Styles Not Found", len(missing))

            # Download Output File
            st.download_button(
                label=f"📥 Download Extracted File ({selected_sheet})",
                data=output_bytes,
                file_name=f"Extracted_{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.markdown("---")

            # Style status panels
            status_col1, status_col2 = st.columns(2)

            with status_col1:
                st.subheader(f"✅ Found Styles ({len(found)})")
                if found:
                    st.dataframe({"Style Number": found}, width='stretch', height=300)
                else:
                    st.info("No matching styles found.")

            with status_col2:
                st.subheader(f"❌ Missing / Not Found Styles ({len(missing)})")
                if missing:
                    st.dataframe({"Style Number": missing}, width='stretch', height=300)
                else:
                    st.success("All style numbers were found in the master file!")

        except Exception as err:
            st.error(f"Processing failed: {err}")