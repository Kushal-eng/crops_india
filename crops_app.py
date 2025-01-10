import streamlit as st
import pandas as pd

# Ensure openpyxl is imported for handling Excel files
from openpyxl import Workbook

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

    # Save to Excel using openpyxl
    try:
        df.to_excel(file_name, index=False, engine="openpyxl")
        st.success(f"Excel file '{file_name}' created successfully!")
    except Exception as e:
        st.error(f"Error creating Excel file: {e}")

def update_excel():
    # Load the Excel sheet
    try:
        df = pd.read_excel(file_name, engine="openpyxl")
        st.write("Current Data:", df)
    except FileNotFoundError:
        st.error(f"The file '{file_name}' does not exist. Please create it first.")
        return
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return

    # Adding new data
    new_data = pd.DataFrame([{
        "Crop Name": "Soybean",
        "Region": "Madhya Pradesh",
        "Season": "Kharif",
        "Harvest Quantity (MT)": 2000,
        "Harvest Date": "2025-01-10",
    }])
    try:
        df = pd.concat([df, new_data], ignore_index=True)

        # Save updated file using openpyxl
        df.to_excel(file_name, index=False, engine="openpyxl")
        st.success("New data added and file updated successfully!")
    except Exception as e:
        st.error(f"Error updating Excel file: {e}")

# Streamlit UI
st.title("Crop Harvest Data Manager")

if st.button("Create Excel File"):
    create_excel()

if st.button("Update Excel File"):
    update_excel()
