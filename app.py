import pandas as pd
import streamlit as st
 
# Function to rename duplicate column names
def rename_duplicate_columns(columns):
    seen = {}
    new_columns = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_columns.append(col)
    return new_columns
 
# Function to clean score values
def clean_scores(df):
    for col in df.columns[1:]:  # Assuming first column is Employee Names
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df[col] = df[col].apply(lambda x: x if 0 <= x <= 5 else 0)
    return df
 
# Caching the data loading to improve performance
@st.cache_data
def load_data(uploaded_file):
    # Read Excel file with MultiIndex headers
    raw_df = pd.read_excel(
        uploaded_file,
        skiprows=1,
        header=[0, 1],
        sheet_name="Employees sheet",
        engine="openpyxl"
    )
    # Flatten MultiIndex headers and rename duplicates
    cleaned_columns = rename_duplicate_columns(
        ['_'.join([str(c) for c in col if pd.notna(c)]).strip() for col in raw_df.columns]
    )
    raw_df.columns = cleaned_columns
 
    cleaned_df = clean_scores(raw_df.copy())
    return raw_df, cleaned_df
 
# Function to extract categories and subcategories from the Excel file
def extract_categories_and_subcategories(uploaded_file, sheet_name="Employees sheet"):
    df = pd.read_excel(
        uploaded_file,
        skiprows=1,
        sheet_name=sheet_name,
        header=[0, 1],
        engine="openpyxl"
    )
    category_mapping = {}
    for col in df.columns:
        category, subcategory = col
        if pd.notna(category) and pd.notna(subcategory):
            category = category.strip()
            subcategory = subcategory.strip()
            if "Unnamed" in category:
                continue
            if category not in category_mapping:
                category_mapping[category] = []
            category_mapping[category].append(subcategory)
    return category_mapping
 
# Dashboard Title
st.title("Employee Skill Matrix Dashboard")
 
# File uploader in sidebar
uploaded_file = st.sidebar.file_uploader("Upload Skill Matrix Excel File", type=["xlsx"])
 
if uploaded_file is not None:
    # Load data and extract categories
    raw_df, cleaned_df = load_data(uploaded_file)
    category_dict = extract_categories_and_subcategories(uploaded_file)
    # Sidebar filter selections
    st.sidebar.header("Filter Options")
    dynamic_categories = list(category_dict.keys())
    selected_categories = st.sidebar.multiselect("Select Categories", dynamic_categories)
 
    selected_subcategories = {}
    scores = {}
    # Build filters with unique keys for each category_subcategory combination
    for category in selected_categories:
        subcategories = category_dict[category]
        selected_subs = st.sidebar.multiselect(f"Select Subcategories for {category}", subcategories)
        if selected_subs:
            selected_subcategories[category] = selected_subs
            for subcat in selected_subs:
                # Create a unique key for the slider using both category and subcategory
                full_col_name = f"{category}_{subcat}"
                scores[full_col_name] = st.sidebar.slider(
                    f"Minimum Score for {category} - {subcat}",
                    1, 5, 3, key=full_col_name
                )
 
    # Generate report button
    if st.sidebar.button("Generate Report"):
        filtered_df = cleaned_df.copy()
        conditions = []
        for category, subcats in selected_subcategories.items():
            for subcat in subcats:
                full_col_name = f"{category}_{subcat}"
                if full_col_name in filtered_df.columns:
                    # Ensure the column is numeric
                    filtered_df[full_col_name] = pd.to_numeric(
                        filtered_df[full_col_name], errors='coerce'
                    ).fillna(0)
                    conditions.append(filtered_df[full_col_name] >= scores[full_col_name])
                else:
                    st.write(f"Column not found: {full_col_name}")
 
        # Combine all filter conditions with AND logic
        if conditions:
            combined_condition = pd.concat(conditions, axis=1).all(axis=1)
            filtered_df = filtered_df.loc[combined_condition]
        else:
            st.write("No filter conditions applied.")
 
        # If no records match, show a warning
        if filtered_df.empty:
            st.warning("No matching records found!")
        else:
            # Only display the employee name column and the filtered subcategory columns
            employee_col = cleaned_df.columns[0]  # Assuming first column is Employee Names
            filter_cols = []
            for category, subcats in selected_subcategories.items():
                for subcat in subcats:
                    full_col_name = f"{category}_{subcat}"
                    if full_col_name in filtered_df.columns:
                        filter_cols.append(full_col_name)
            display_cols = [employee_col] + filter_cols
 
            # Rename columns to show only subcategories
            display_df = filtered_df[display_cols].copy().reset_index(drop=True)
            display_df.columns = [col.split("_")[-1] if "_" in col else col for col in display_df.columns]
            display_df.index += 1

            st.write("### Filtered Employee Report")
            st.dataframe(display_df)
            # Option to download the filtered data as CSV
            csv = filtered_df[display_cols].to_csv(index=False).encode('utf-8')
            st.download_button(
                "Download Filtered Report as CSV",
                data=csv,
                file_name='filtered_report.csv',
                mime='text/csv'
            )
else:
    st.info("Please upload an Excel file to start.")
