import io
import os
import tempfile

import pandas as pd
import streamlit as st

from address_sorter import AddressSorter


st.set_page_config(page_title="Address Sorter", layout="wide")
st.title("Address Sorter (Web UI)")
st.caption("Upload a CSV/XLSX, review categorized results, and download the sorted workbook.")

with st.sidebar:
	st.header("1) Upload Input File")
	uploaded = st.file_uploader("CSV or Excel (.csv, .xlsx)", type=["csv", "xlsx"]) 
	st.markdown("---")
	st.header("2) Actions")
	process_clicked = st.button("Process Addresses", type="primary", disabled=not uploaded)

if uploaded and process_clicked:
	# Persist uploaded file to a temp path for AddressSorter (expects a path)
	suffix = ".xlsx" if uploaded.name.lower().endswith(".xlsx") else ".csv"
	with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
		tmp.write(uploaded.getvalue())
		input_path = tmp.name

	# Run the sorter pipeline
	with st.spinner("Processing... This can take a moment for large files."):
		sorter = AddressSorter(input_path)
		sorter.load_data()
		roe_candidates = sorter.initial_sort()
		sorter.process_roe_deduplication(roe_candidates)
		sorter.create_flagged_tab()
		sorter.create_unit_count_tab()

	# Clean up temp file
	try:
		os.unlink(input_path)
	except Exception:
		pass

	# Summary metrics
	col1, col2, col3, col4 = st.columns(4)
	with col1:
		st.metric("Total", len(sorter.tabs.get("All", pd.DataFrame())))
	with col2:
		st.metric("ROE", len(sorter.tabs.get("ROE", pd.DataFrame())))
	with col3:
		st.metric("Remove", len(sorter.tabs.get("Remove", pd.DataFrame())))
	with col4:
		flagged_df = sorter.tabs.get("Flagged for Review", pd.DataFrame())
		st.metric("Flagged", len(flagged_df))

	st.markdown("---")

	# Tabbed preview of results
	sheet_names = list(sorter.tabs.keys())
	st.subheader("Preview Sheets")
	tab_objects = st.tabs(sheet_names)
	for i, name in enumerate(sheet_names):
		with tab_objects[i]:
			df = sorter.tabs.get(name)
			if df is None or (hasattr(df, "empty") and df.empty):
				st.info("No data in this sheet.")
			else:
				st.dataframe(df, use_container_width=True, hide_index=True)

	# Provide a download of the Excel workbook
	buffer = io.BytesIO()
	with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
		for tab_name, df in sorter.tabs.items():
			if df is not None and not df.empty:
				df.to_excel(writer, sheet_name=tab_name, index=False)
	buffer.seek(0)

	default_output_name = os.path.splitext(uploaded.name)[0] + "_sorted.xlsx"
	st.download_button(
		label="Download Sorted Excel",
		data=buffer,
		file_name=default_output_name,
		mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		type="primary",
	)
else:
	st.info("Upload a CSV or XLSX file in the sidebar, then click 'Process Addresses'.")
