import os
from flask import Flask, request, render_template, send_from_directory
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

app = Flask(__name__)

# Path to save uploaded files
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def extract_college_data(input_file, college_name, streams, output_file):
    # Load the main scrutiny file
    xl = pd.ExcelFile(input_file)
    df = xl.parse("Sheet1")  # Adjust sheet name if different
    
    # Strip whitespace from column names
    df.columns = df.columns.str.strip()

    # Check if required columns exist
    required_columns = ["CollegeName", "Stream", "ReservationCategory", "Percentage"]
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        print(f"Missing columns: {missing_columns}")
        return

    # If 'everything' is in the streams list, get all unique streams from the dataset
    if 'everything' in streams:
        streams = df['Stream'].unique().tolist()

    # Create an Excel writer object to write multiple tables
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        row_start = 0  # Initial starting row

        for stream in streams:
            # Filter data for the specified college and stream
            college_df = df[(df["CollegeName"].str.contains(college_name, case=False, na=False)) & 
                            (df["Stream"].str.contains(stream, case=False, na=False))]

            if college_df.empty:
                print(f"No data found for college: {college_name} and stream: {stream}")
                continue

            # Group by ReservationCategory and find min/max marks
            result = college_df.groupby("ReservationCategory")["Percentage"].agg(["max", "min"]).reset_index()
            result.columns = ["Category", "High", "Low"]

            # Create a row for the subject (stream name) and merge it across the three columns
            subject_row = pd.DataFrame([[stream, "", ""]], columns=["Category", "High", "Low"])

            # Write the subject row and result table starting from the current row
            subject_row.to_excel(writer, index=False, header=True, startrow=row_start)
            result.to_excel(writer, index=False, startrow=row_start + 1)

            # Load the workbook and get the active sheet
            workbook = writer.book
            sheet = workbook.active

            # Merge cells for the subject row (stream name)
            sheet.merge_cells(f'A{row_start + 1}:C{row_start + 1}')  # Merging cells for the subject row
            cell = sheet[f'A{row_start + 1}']
            cell.value = stream  # Set the stream name
            cell.alignment = Alignment(horizontal='center', vertical='center')  # Center align the text
            cell.font = Font(bold=True)  # Make the text bold

            # Update row_start for the next table (add 4 extra rows for spacing)
            row_start += len(result) + 4

    print(f"Data extracted and saved to {output_file}")


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/extract_data', methods=['POST'])
def extract_data():
    # Get the form data
    college_name = request.form['college_name']
    streams_input = request.form['streams']
    streams = [stream.strip() for stream in streams_input.split(',')]  # Convert input into a list of streams

    # Handle file upload
    file = request.files['input_excel']
    input_file = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(input_file)

    # Output file name
    output_file = os.path.join(app.config['UPLOAD_FOLDER'], f"{college_name}_multiple_streams_marks.xlsx")

    # Call the extract_college_data function
    extract_college_data(input_file, college_name, streams, output_file)

    # Return the result file to the user
    return send_from_directory(app.config['UPLOAD_FOLDER'], os.path.basename(output_file), as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
