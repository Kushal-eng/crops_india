import pandas as pd

# File path for the Excel sheet
file_name = ""C:\Users\kotra\OneDrive\Desktop\kushal projects\crop_harvest_data.xlsx""

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
    print(f"Excel file '{file_name}' created successfully!")

def update_excel():
    # Load the Excel sheet
    try:
        df = pd.read_excel(file_name)
    except FileNotFoundError:
        print(f"The file '{file_name}' does not exist. Please create it first.")
        return

    # Display existing data
    print("Current Data:")
    print(df)

    # Example: Adding new data
    new_data = {
        "Crop Name": "Soybean",
        "Region": "Madhya Pradesh",
        "Season": "Kharif",
        "Harvest Quantity (MT)": 2000,
        "Harvest Date": "2025-01-10",
    }
    df = df.append(new_data, ignore_index=True)

    # Save the updated DataFrame to the Excel file
    df.to_excel(file_name, index=False)
    print("New data added and file updated successfully!")

# Main function
if __name__ == "__main__":
    print("1. Create Excel")
    print("2. Update Excel")
    choice = int(input("Enter your choice: "))

    if choice == 1:
        create_excel()
    elif choice == 2:
        update_excel()
    else:
        print("Invalid choice!")
