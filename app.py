from io import BytesIO
from flask import Flask, render_template, request, send_file
import pandas as pd

app = Flask(__name__)


@app.route('/')
def show_main():  # put application's code here
    return render_template('upload.html')


def combine_dfs(df1, df2):
    return pd.concat([df1, df2], ignore_index=True)


@app.route('/upload', methods=['POST'])
def upload():
    if 'first_file' not in request.files or 'second_file' not in request.files:
        return "Upload both files"

    first_file = request.files['first_file']
    second_file = request.files['second_file']

    # Check if both files are valid Excel files
    if first_file.filename == '' or second_file.filename == '':
        return "No selected file(s)"

    if first_file and first_file.filename.endswith('.xlsx') and second_file and second_file.filename.endswith('.xlsx'):
        # Read both Excel files into pandas DataFrames
        df1 = pd.read_excel(first_file)
        df2 = pd.read_excel(second_file)
        
        df_combined = combine_dfs(df1, df2)

        # Get the action the user selected
        action = request.form.get('action')

        # If the user chose to download the file
        if action == 'download':
            # Create an in-memory buffer to hold the new Excel file
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_combined.to_excel(writer, index=False)

            # Rewind the buffer
            output.seek(0)

            # Send the processed file back to the user as a download
            return send_file(output, download_name="processed_file.xlsx", as_attachment=True)

        # If the user chose to view the first 10 rows
        elif action == 'show':

            # Convert the DataFrame rows to a list of dictionaries (for easy rendering in Jinja2)
            table_data = df_combined.to_dict(orient='records')

            # Pass the table data and columns to the template
            return render_template('show_data.html', columns=df_combined.columns, rows=table_data)

    return "Invalid file format"


if __name__ == '__main__':
    app.run(debug=True)
