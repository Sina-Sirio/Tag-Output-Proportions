import streamlit as st
import pandas as pd
import io

# Detect if dark mode is enabled
is_dark_mode = st.get_option("theme.base") == "dark"

# Custom CSS to make the UI look good in both light and dark modes
st.markdown(
    f"""
    <style>
    body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        background-color: {"#1E1E1E" if is_dark_mode else "#f5f5f7"};
        color: {"#ffffff" if is_dark_mode else "#000000"};
    }}
    .stButton>button {{
        background-color: #007aff;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5em 1em;
    }}
    .stDataFrame {{
        background-color: {"#262730" if is_dark_mode else "white"};
        color: {"#ffffff" if is_dark_mode else "#000000"};
    }}
    .sidebar .sidebar-content {{
        background-image: {"linear-gradient(#2D2D2D, #1E1E1E)" if is_dark_mode else "linear-gradient(#fff, #f5f5f7)"};
    }}
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
            df = df[["document", "Topic"]].copy()
            unique_topics = sorted(df["Topic"].dropna().unique())

            st.subheader("Select the Correct Topics")
            correct_topics = st.multiselect("Click to select topics that are actually correct", unique_topics, default=unique_topics)

            # Improved lighter color scheme
            def highlight_row(row):
                color = "#A3E4A3" if row["Topic"] in correct_topics else "#F5A3A3"
                return [f"background-color: {color}; color: {'#000000' if is_dark_mode else '#000000'}"] * len(row)

            styled_df = df.style.apply(highlight_row, axis=1)

            st.subheader("Validated Data")
            st.write("Rows highlighted **green** are correct; **red** are incorrect.")
            st.dataframe(styled_df, height=500)

            # Export function for Excel with better colors
            def to_excel_with_styles(dataframe, correct_topics):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet("Validated")
                    writer.sheets["Validated"] = worksheet

                    correct_format = workbook.add_format({'bg_color': '#A3E4A3', 'border': 1, 'font_color': '#000000'})
                    incorrect_format = workbook.add_format({'bg_color': '#F5A3A3', 'border': 1, 'font_color': '#000000'})
                    header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#3A3B3C' if is_dark_mode else '#D3D3D3', 'font_color': '#FFFFFF' if is_dark_mode else '#000000'})

                    for col_num, value in enumerate(dataframe.columns):
                        worksheet.write(0, col_num, value, header_format)

                    for row_num, row in enumerate(dataframe.itertuples(index=False), start=1):
                        for col_num, value in enumerate(row):
                            fmt = correct_format if row.Topic in correct_topics else incorrect_format
                            worksheet.write(row_num, col_num, value, fmt)

                    worksheet.set_column(0, 0, 50)
                    worksheet.set_column(1, 1, 20)
                return output.getvalue()

            excel_data = to_excel_with_styles(df, correct_topics)
            st.subheader("Download Validated Data")
            st.download_button("Download Excel", data=excel_data, file_name="validated_topics.xlsx")

            st.subheader("Topics Overview")
            topic_counts = df["Topic"].value_counts().rename_axis("Topic").reset_index(name="Count")
            total = topic_counts["Count"].sum()
            topic_counts["Percentage"] = (topic_counts["Count"] / total * 100).round(1)
            topic_counts = topic_counts.sort_values(by="Count", ascending=False)

            def style_topic_row(row):
                color = "#A3E4A3" if row["Topic"] in correct_topics else "#F5A3A3"
                return [f"background-color: {color}; color: {'#000000' if is_dark_mode else '#000000'}"] * len(row)

            styled_overview = topic_counts.style.apply(style_topic_row, axis=1)
            st.dataframe(styled_overview, height=300)

            # Export function for topic overview Excel with lighter colors
            def to_excel_overview(dataframe, correct_topics):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet("Overview")
                    writer.sheets["Overview"] = worksheet

                    correct_format = workbook.add_format({'bg_color': '#A3E4A3', 'border': 1, 'font_color': '#000000'})
                    incorrect_format = workbook.add_format({'bg_color': '#F5A3A3', 'border': 1, 'font_color': '#000000'})
                    header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#3A3B3C' if is_dark_mode else '#D3D3D3', 'font_color': '#FFFFFF' if is_dark_mode else '#000000'})

                    for col_num, value in enumerate(dataframe.columns):
                        worksheet.write(0, col_num, value, header_format)

                    for row_num, row in enumerate(dataframe.itertuples(index=False), start=1):
                        for col_num, value in enumerate(row):
                            fmt = correct_format if row.Topic in correct_topics else incorrect_format
                            worksheet.write(row_num, col_num, value, fmt)

                    worksheet.set_column(0, 0, 30)
                    worksheet.set_column(1, 2, 15)
                return output.getvalue()

            overview_excel_data = to_excel_overview(topic_counts, correct_topics)
            st.subheader("Download Topics Overview")
            st.download_button("Download Overview Excel", data=overview_excel_data, file_name="topics_overview.xlsx")
