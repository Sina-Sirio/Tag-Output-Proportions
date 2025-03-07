import streamlit as st
import pandas as pd
import io

# Custom CSS for an Apple-esque aesthetic
st.markdown(
    """
    <style>
    body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        background-color: #f5f5f7;
    }
    .stButton>button {
        background-color: #007aff;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5em 1em;
    }
    .sidebar .sidebar-content {
        background-image: linear-gradient(#fff, #f5f5f7);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Topic Validator")
st.write("Drag and drop your Excel file below. Then select the correct topics for this document.")

# File uploader: allow only Excel files
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    else:
        if "Topic" not in df.columns:
            st.error("The Excel file does not have a 'Topic' column.")
        else:
            # Keep only the columns we need
            df = df[["document", "Topic"]].copy()
            
            # Extract unique topics (drop NA and sort)
            unique_topics = sorted(df["Topic"].dropna().unique())
            
            st.subheader("Select the Correct Topics")
            # Use multiselect to mimic toggle buttons for correct topics.
            correct_topics = st.multiselect("Click to select topics that are actually correct", unique_topics, default=unique_topics)
            
            # Define a style function to highlight rows in the validated table.
            def highlight_row(row):
                # Green if correct, red if not.
                color = "background-color: #c6f6d5" if row["Topic"] in correct_topics else "background-color: #fed7d7"
                return [color] * len(row)
            
            styled_df = df.style.apply(highlight_row, axis=1)
            
            st.subheader("Validated Data")
            st.write("Rows highlighted **green** are correct; **red** are incorrect.")
            st.dataframe(styled_df, height=500)
            
            # Function to create an Excel file for the validated data with colours and borders.
            def to_excel_with_styles(dataframe, correct_topics):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet("Validated")
                    writer.sheets["Validated"] = worksheet
                    
                    # Define formats.
                    correct_format = workbook.add_format({'bg_color': '#c6f6d5', 'border': 1})
                    incorrect_format = workbook.add_format({'bg_color': '#fed7d7', 'border': 1})
                    header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D3D3D3'})
                    
                    # Write the header.
                    for col_num, value in enumerate(dataframe.columns):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # Write the data rows.
                    for row_num, row in enumerate(dataframe.itertuples(index=False), start=1):
                        for col_num, value in enumerate(row):
                            # Apply correct formatting based on the Topic.
                            fmt = correct_format if row.Topic in correct_topics else incorrect_format
                            worksheet.write(row_num, col_num, value, fmt)
                    
                    # Set column widths for a nicer layout.
                    worksheet.set_column(0, 0, 50)  # document column
                    worksheet.set_column(1, 1, 20)  # Topic column
                processed_data = output.getvalue()
                return processed_data
            
            excel_data = to_excel_with_styles(df, correct_topics)
            st.subheader("Download Validated Data")
            st.download_button("Download Excel", data=excel_data, file_name="validated_topics.xlsx")
            
            # Create an aggregated topics overview: counts and percentage proportions.
            st.subheader("Topics Overview")
            topic_counts = df["Topic"].value_counts().rename_axis("Topic").reset_index(name="Count")
            total = topic_counts["Count"].sum()
            topic_counts["Percentage"] = (topic_counts["Count"] / total * 100).round(1)
            # Order by count descending.
            topic_counts = topic_counts.sort_values(by="Count", ascending=False)
            
            # Define a style function for the overview.
            def style_topic_row(row):
                color = "background-color: #c6f6d5" if row["Topic"] in correct_topics else "background-color: #fed7d7"
                return [color] * len(row)
            
            styled_overview = topic_counts.style.apply(style_topic_row, axis=1)
            st.dataframe(styled_overview, height=300)
            
            # Function to create an Excel file for the topics overview with styles.
            def to_excel_overview(dataframe, correct_topics):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet("Overview")
                    writer.sheets["Overview"] = worksheet
                    
                    # Define formats.
                    correct_format = workbook.add_format({'bg_color': '#c6f6d5', 'border': 1})
                    incorrect_format = workbook.add_format({'bg_color': '#fed7d7', 'border': 1})
                    header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D3D3D3'})
                    
                    # Write header.
                    for col_num, value in enumerate(dataframe.columns):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # Write data rows.
                    for row_num, row in enumerate(dataframe.itertuples(index=False), start=1):
                        for col_num, value in enumerate(row):
                            fmt = correct_format if row.Topic in correct_topics else incorrect_format
                            worksheet.write(row_num, col_num, value, fmt)
                    
                    # Set column widths.
                    worksheet.set_column(0, 0, 30)  # Topic column
                    worksheet.set_column(1, 2, 15)  # Count and Percentage columns
                processed_data = output.getvalue()
                return processed_data
            
            overview_excel_data = to_excel_overview(topic_counts, correct_topics)
            st.subheader("Download Topics Overview")
            st.download_button("Download Overview Excel", data=overview_excel_data, file_name="topics_overview.xlsx")
