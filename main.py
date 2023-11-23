import pandas as pd
import logging
from dotenv import load_dotenv
import os 

load_dotenv()

CHORD_LENGTH =  float(os.environ['CHORD_LENGTH'])

logging.info(f'Current Chord Length value :{CHORD_LENGTH}')

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def calculate_coefficient(value, speed, density, chord_length):
    print(f"Value type: {type(value)}, Speed type: {type(speed)}, Density type: {type(density)}, Chord Length type: {type(chord_length)}")

    print(f'Value: {value}')
    print(f'Speed: {speed}')
    print(f'Density: {density}')
    print(f'Chord Length: {chord_length}')
    if speed == 0 or density == 0:
        return float('nan')  # Return NaN if speed or density is zero to avoid division by zero
    return value / (0.5 * density * (speed ** 2) * chord_length)


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

def process_sheet(data_dict):
    processed_data_frames = []

    for degree, df in data_dict.items():
        df.dropna(subset=['Drag', 'Lift', 'Momentum', 'm/s', 'Density'], inplace=True)
        # Skip further processing if DataFrame is empty
        if df.empty:
            logging.warning(f"DataFrame for degree {degree} is empty after dropping NaN values.")
            continue
        # Group by the 'm/s' column
        grouped = df.groupby('m/s')
        # Calculate the mean for 'Drag', 'Lift', 'Momentum', and take the first 'Density' value
        # for each unique 'm/s' value
        aggregated_data = grouped.agg(
            {
                'Drag': 'mean',
                'Lift': 'mean',
                'Momentum': 'mean',
                'Density': 'first'
            }
        ).reset_index()

        # In your process_sheets function, replace the lambda functions with:
        aggregated_data['Coefficient of Lift'] = aggregated_data.apply(
            lambda row: try_calculate(row['Lift'], row['m/s'], row['Density'], CHORD_LENGTH)
            if row['Lift'] is not None and row['m/s'] is not None and row['Density'] is not None else float('nan'), axis=1
            )
        
        aggregated_data['Coefficient of Drag'] = aggregated_data.apply(
            lambda row: try_calculate(row['Drag'], row['m/s'], row['Density'], CHORD_LENGTH)
            if row['Drag'] is not None and row['m/s'] is not None and row['Density'] is not None else float('nan'), axis=1)

        aggregated_data['Degree'] = degree

        processed_data_frames.append(aggregated_data)   
    
    # Concatenate all DataFrames into a single DataFrame
    final_df = pd.concat(processed_data_frames)

    # Reorder columns
    final_df = final_df[['Degree', 'm/s', 'Drag', 'Lift', 'Momentum', 'Density', 'Coefficient of Lift', 'Coefficient of Drag']]

    return final_df 




if __name__ == "__main__":  # Change "main" to "__main__"
    file_path = 'data/Sangeetha_JP_Data.xlsx'
    data_dict = read_excel(file_path)
    logging.info(data_dict)
    logging.info("Data loaded successfully")
    processed_data_df = process_sheet(data_dict)
    # Ensure the 'output' directory exists or handle it appropriately
    output_file_path = 'output/Expected_Output.xlsx'  # or 'output/Expected_Output.xlsx'
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w') as writer:
        processed_data_df.to_excel(writer, sheet_name='Expected Output', index=False)
    logging.info(f"Processed data written to {output_file_path}")
