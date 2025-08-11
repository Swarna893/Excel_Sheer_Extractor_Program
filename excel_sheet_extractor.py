# import pandas as pd
# import os

# def extract_events_sheet_from_file(file_path):
#     """
#     Extracts the 'events' sheet from a single Excel file and saves it.

#     Args:
#         file_path (str): The full path to the Excel file.
#     """
#     if not os.path.isfile(file_path):
#         print(f"Error: The file '{file_path}' does not exist.")
#         return

#     try:
#         # Read all sheets from the Excel file to check for 'event'
#         excel_file = pd.ExcelFile(file_path)

#         # Check if the 'events' sheet exists
#         if 'events' in excel_file.sheet_names:
#             print(f"Found 'events' sheet in '{os.path.basename(file_path)}'. Extracting...")

#             # Read the 'event' sheet into a DataFrame
#             event_df = pd.read_excel(file_path, sheet_name='events')

#             # Create the new filename
#             base_name = os.path.splitext(os.path.basename(file_path))[0]
#             directory = os.path.dirname(file_path)
#             new_filename = f"events_{base_name}.xlsx"
#             new_file_path = os.path.join(directory, new_filename)

#             # Save the DataFrame to a new Excel file
#             event_df.to_excel(new_file_path, index=False)
#             print(f"Successfully saved '{new_filename}' to the same directory.")
#         else:
#             print(f"Skipping '{os.path.basename(file_path)}' as it does not contain an 'events' sheet.")

#     except Exception as e:
#         print(f"An error occurred while processing the file: {e}")

# # Specify the full path to your single Excel file
# file_to_process = r'D:/Final_Year_Project/datasets_Copy/up01AUG.mdb.xls' 

# # Call the function
# extract_events_sheet_from_file(file_to_process)




import pandas as pd
import os

def extract_events_sheets_from_directory(directory_path):
    """
    Extracts the 'events' sheet from every Excel file in a directory and 
    saves it as a new Excel file.

    Args:
        directory_path (str): The path to the directory containing the Excel files.
    """
    # 📝 Check if the provided path is a valid directory
    if not os.path.isdir(directory_path):
        print(f"Error: The provided path '{directory_path}' is not a valid directory.")
        return

    print(f"Starting to process files in directory: {directory_path}\n")

    # 📝 Loop through all files in the directory
    for filename in os.listdir(directory_path):
        # 📝 Check if the file is an Excel file (either .xlsx or .xls)
        if filename.endswith(('.xlsx', '.xls')):
            file_path = os.path.join(directory_path, filename)
            
            try:
                # 📝 Read all sheets from the Excel file
                excel_file = pd.ExcelFile(file_path)
                
                # 📝 Check if the 'events' sheet exists in the file
                if 'events' in excel_file.sheet_names:
                    print(f"Found 'events' sheet in '{filename}'. Extracting...")
                    
                    # 📝 Read the 'events' sheet into a DataFrame
                    events_df = pd.read_excel(file_path, sheet_name='events')
                    
                    # 📝 Create the new filename with the "events_" prefix
                    base_name = os.path.splitext(filename)[0]
                    new_filename = f"events_{base_name}.xlsx"
                    new_file_path = os.path.join(directory_path, new_filename)
                    
                    # 📝 Save the DataFrame to a new Excel file
                    events_df.to_excel(new_file_path, index=False)
                    print(f"Successfully saved '{new_filename}'.\n")
                else:
                    print(f"Skipping '{filename}' as it does not contain an 'events' sheet.\n")
            
            except Exception as e:
                print(f"An error occurred while processing '{filename}': {e}\n")

# 📌 Specify the path to the directory you want to process
# Replace this with the path to your folder containing all the Excel files.
# Use a raw string (r'') to avoid 'unicodeescape' errors.
directory_to_process = r'D:\Final_Year_Project\datasets_Copy'

# 🚀 Call the function to start the process
extract_events_sheets_from_directory(directory_to_process)