import streamlit as st
import pandas as pd

# File path for the Excel sheet
file_name = "crop_harvest_data.xlsx"

def create_excel():
    # Sample data
    data = {
        "Crop Name": ["Rice", "Wheat", "Maize", "Sugarcane", "Cotton"],
        "Region": ["Punjab", "Haryana", "Maharashtra", "Uttar Pradesh", "Gujarat"],
        "Season": ["Kharif", "Rabi", "Kharif", "Annual", "Kharif"],
        "Harvest Quantity (MT)": [5000, 4200, 3000, 8000, 2500],
        "Harvest Date": ["2025-01-05", "2025-01-06", "2025-01-07", "2025-01-08", "2025-01-09"],
    }

    # Create a DataFrame
    df = pd.DataFrame(data)

    # Save to Excel
    df.to_excel(file_name, index=False)
    st.success(f"Excel file '{file_name}' created successfully!")

def update_excel():
    try:
        df = pd.read_excel(file_name)
        st.write("Current Data:", df)
    except FileNotFoundError:
        st.error(f"The file '{file_name}' does not exist. Please create it first.")
        return

    # Adding new data
    new_data = pd.DataFrame([{
        "Crop Name": "Soybean",
        "Region": "Madhya Pradesh",
        "Season": "Kharif",
        "Harvest Quantity (MT)": 2000,
        "Harvest Date": "2025-01-10",
    }])
    df = pd.concat([df, new_data], ignore_index=True)

    # Save updated file
    df.to_excel(file_name, index=False)
    st.success("New data added and file updated successfully!")

st.title("Crop Harvest Data Manager")

if st.button("Create Excel"):
    create_excel()

if st.button("Update Excel"):
    update_excel()
