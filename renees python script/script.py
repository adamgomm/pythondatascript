import pandas as pd

def custom_sort_key(value):
    # Split the value into alphabetic and numeric parts
    alpha_part = ''.join(filter(str.isalpha, value))
    num_part = ''.join(filter(str.isdigit, value))
    
    # Convert the numeric part to an integer if it's not empty
    num_part = int(num_part) if num_part else 0
    
    # Return a tuple with the alphabetic part and the numeric part converted to integers
    return (alpha_part, num_part)

def organize_excel(input_file, output_file):
    # Read the Excel file into a pandas DataFrame, excluding the header row
    df_header = pd.read_excel(input_file, nrows=1, header=None)
    df_data = pd.read_excel(input_file, skiprows=1, header=None)
    
    # Get the cell references as a list, excluding the first row
    cell_references = df_data[0].tolist()
    
    # Sort the cell references using the custom key
    sorted_cell_references = sorted(cell_references, key=custom_sort_key)
    
    # Create a new DataFrame with sorted cell references
    df_sorted = pd.DataFrame(sorted_cell_references, columns=[0])
    
    # Merge the sorted DataFrame with the original DataFrame to maintain data integrity
    df_result = pd.concat([df_header, df_sorted.merge(df_data, on=0, how='left')], ignore_index=True)
    
    # Write the sorted DataFrame to a new Excel file
    df_result.to_excel(output_file, index=False, header=False)
    
    print("Excel spreadsheet organized successfully!")

# Example usage:
input_file = r"C:\Users\Curis\Desktop\renees python script\Adam Excel sheet.xlsx"  # Replace with the path to your input Excel file
output_file = r"C:\Users\Curis\Desktop\renees python script\output.xlsx"  # Replace with the desired output Excel file path

organize_excel(input_file, output_file)
