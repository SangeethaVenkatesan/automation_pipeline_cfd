import pandas as pd
import logging
from dotenv import load_dotenv
import os 

import threading
import time 
load_dotenv()

CHORD_LENGTH =  float(os.environ['CHORD_LENGTH'])
import plotly.express as px
import xlsxwriter


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

logging.info(f'Current Chord Length value :{CHORD_LENGTH}')


def calculate_coefficient(value, speed, density, chord_length):
    if speed == 0 or density == 0:
        return float('nan')  # Return NaN if speed or density is zero to avoid division by zero
    return value / 0.5 * density * (speed ** 2) * chord_length



def try_calculate(value, speed, density, chord_length):
    try:
        return calculate_coefficient(value, speed, density, chord_length)
    except TypeError:
        print(f"Error with row: Value: {value}, Speed: {speed}, Density: {density}, Chord Length: {chord_length}")
        return float('nan')

def read_excel(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        sheets_dict = {}

        for sheet_name in xl.sheet_names:
            sheets_dict[sheet_name] = xl.parse(sheet_name)
        
        return sheets_dict

    except Exception as e:
        logging.error("Error reading Excel file: %s", e)
        return {}
    

def process_single_sheet(degree, df, results):
    if df.empty:
        logging.warning(f"DataFrame for degree {degree} is empty.")
        return
    # Trim spaces from column names
    df.columns = df.columns.str.strip()

    # Check if the required columns are present
    required_columns = ['Drag', 'Lift', 'Momentum', 'Density']
    if not all(col in df.columns for col in required_columns):
        logging.error(f"Missing required columns in DataFrame for degree {degree}.")
        return

    grouped = df.groupby('m/s')
    aggregated_data = grouped.agg({'Drag': 'mean', 'Lift': 'mean', 'Momentum': 'mean', 'Density': 'first'}).reset_index()

    # Renaming columns for clarity

    aggregated_data.rename(columns={
        'Drag': 'Average Drag', 
        'Lift': 'Average Lift', 
        'Momentum': 'Average Momentum'
    }, inplace=True)

    # Calculating coefficients
    aggregated_data['Coefficient of Lift'] = aggregated_data.apply(lambda row: try_calculate(row['Average Lift'], row['m/s'], row['Density'], CHORD_LENGTH), axis=1)
    aggregated_data['Coefficient of Drag'] = aggregated_data.apply(lambda row: try_calculate(row['Average Drag'], row['m/s'], row['Density'], CHORD_LENGTH), axis=1)
    aggregated_data['Degree'] = int(degree)


    cols = aggregated_data.columns.tolist()
    cols = [cols[-1]] + cols[:-1]  # Move the last column (Degree) to the front
    aggregated_data = aggregated_data[cols]


    results[degree] = aggregated_data


        

def process_sheets_with_threading(data_dict):
    threads = []
    results = {}

    for degree, df in data_dict.items():
        thread = threading.Thread(target=process_single_sheet, args=(degree, df, results))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

    return pd.concat(results.values()) if results else pd.DataFrame()


if __name__ == "__main__":  
    start_time = time.time()  # Start the timer
    file_path = 'data/Sangeetha_JP_Data_Edited.xlsx'
    data_dict = read_excel(file_path)
    logging.info("Data loaded successfully")
    
    processed_data = process_sheets_with_threading(data_dict)
    processed_data = processed_data.round(3)  # Round all numeric data to 3 decimal places
    
    # Sort the DataFrame in ascending order based on 'Degree'
    processed_data.sort_values(by='Degree', inplace=True)

    # Define output file path
    output_file_path = 'output/Processed_Data_updated.xlsx'
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        processed_data.to_excel(writer, sheet_name='Processed Data', index=False)

        # Access the XlsxWriter workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets['Processed Data']

        # Define a header format: bold text with borders
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'bg_color': '#ADD8E6'  # Light blue color
        })

        # Define a cell format for borders
        cell_format = workbook.add_format({'border': 1})

        # Write the column headers with the defined format
        for col_num, value in enumerate(processed_data.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Apply cell format for borders to all cells
        for row in range(1, len(processed_data) + 1):
            for col in range(len(processed_data.columns)):
                worksheet.write(row, col, processed_data.iloc[row - 1, col], cell_format)
        # Set the column width and format for better readability
        for column in processed_data:
            column_length = max(processed_data[column].astype(str).map(len).max(), len(column))
            col_idx = processed_data.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, column_length)

        logging.info(f"Data saved successfully to {output_file_path}")
        # Generate Plotly figures
        fig1 = px.line(processed_data, x='Degree', y='Average Drag', title='Degree vs Average Drag')
        fig2 = px.line(processed_data, x='Degree', y='Average Lift', title='Degree vs Average Lift')
        fig3 = px.line(processed_data, x='Average Drag', y='Average Lift', title='Average Drag vs Average Lift')
        fig4 = px.line(processed_data, x='Degree', y=['Average Drag', 'Average Lift'], title='Degree vs Average Drag and Average Lift')

        # Save figures as images
        fig1.write_image("fig1.png")
        fig2.write_image("fig2.png")
        fig3.write_image("fig3.png")
        fig4.write_image("fig4.png")

        # Calculate dynamic image height based on an estimated height
        estimated_image_height = 40  # Adjust this value as needed to prevent overlap
        start_row_for_images = len(processed_data) + 2  # Start placing images a few rows after the data
        
        # After writing the DataFrame to the Excel sheet
        image_insert_col = 10  # Column index for images ('I' in Excel)

        
        for i, img_file in enumerate(["fig1.png", "fig2.png", "fig3.png", "fig4.png"]):
            # Calculate the cell position (e.g., 'I1', 'I16', etc.)
            cell_position = f"{xlsxwriter.utility.xl_col_to_name(image_insert_col)}{start_row_for_images + i * estimated_image_height}"
            worksheet.insert_image(cell_position, img_file)

        logging.info(f"Data and plots saved successfully to {output_file_path}")
        end_time = time.time()  # End the timer
        execution_time = end_time - start_time  # Calculate total execution time
        logging.info(f"Data processed and saved successfully in {execution_time:.2f} seconds")
    