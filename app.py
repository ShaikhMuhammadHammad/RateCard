import os
import pandas as pd
from flask import Flask, request, render_template, redirect, send_file, flash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xls', 'xlsx','csv'}
app.secret_key = 'your_secret_key'  # Required for flashing messages

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        if 'file1' not in request.files or 'file2' not in request.files:
            flash("Missing file(s) in request.")
            return redirect(request.url)
        
        file1 = request.files['file1']
        file2 = request.files['file2']
        
        if file1.filename == '' or file2.filename == '':
            flash("One or both files have no filename.")
            return redirect(request.url)
        
        if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename):
            filename1 = secure_filename(file1.filename)
            filename2 = secure_filename(file2.filename)

            if not os.path.exists(app.config['UPLOAD_FOLDER']):
                os.makedirs(app.config['UPLOAD_FOLDER'])

            file1_path = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], filename1))
            file2_path = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], filename2))
            
            file1.save(file1_path)
            file2.save(file2_path)
   
            df1 = pd.read_excel(file1_path)
            df2 = pd.read_excel(file2_path)
            
            print("df1 columns:", df1.columns)
            print("df2 columns:", df2.columns)
            
            try:
                df3 = pd.read_excel(file1_path, sheet_name='Sheet1')
                df4 = pd.read_excel(file1_path, sheet_name='Sheet2')
            except Exception as e:
                print(f"Error loading df3 and df4: {e}")
                df3 = pd.DataFrame()
                df4 = pd.DataFrame()

            try:
                df5 = pd.read_excel(file1_path, sheet_name='Sheet3')  # Adjust as necessary
                df5.columns = df5.columns.str.strip()  # Remove trailing spaces from column names
            except Exception as e:
                print(f"Error loading df5: {e}")
                df5 = pd.DataFrame()

            print("df3 columns:", df3.columns)
            print("df4 columns:", df4.columns)
            print("df5 columns:", df5.columns)

            # Normal rates
            for index, row in df2.iterrows():
                condition = (
                    (df1['Channel'].str.lower() == row['Channel'].lower()) &
                    (row['AdStart'] >= df1['Starttime']) &
                    (row['AdEnd'] <= df1['EndTime'])
                )
                
                if condition.any():
                    matching_rate = df1.loc[condition, 'Rate'].values[0]
                    df2.at[index, 'RPM'] = matching_rate
            
            na_rows = pd.isnull(df2['RPM'])
            for index, row in df2[na_rows].iterrows():
                condition = (
                    (df1['Channel'].str.lower() == row['Channel'].lower()) &
                    (row['AdStart'] <= df1['EndTime'])
                )
                if condition.any():
                    matching_rate = df1.loc[condition, 'Rate'].values[0]
                    df2.at[index, 'RPM'] = matching_rate

            # Applying conditions separately for Special Rates
            if not df3.empty:
                for index, row in df2.iterrows():
                    condition = (
                        (df3['Channel'].str.lower() == row['Channel'].lower()) &
                        (row['AdStart'] >= df3['StartTime']) &
                        (row['AdEnd'] <= df3['EndTime']) &
                        (df3['programName'] == row['programName']) &
                        (df3['Day'] == row['Day'])
                    )
                    if condition.any():
                        matching_rate = df3.loc[condition, 'Rate'].values[0]
                        df2.at[index, 'RPM'] = matching_rate

            if not df4.empty:
                for index, row in df2.iterrows():
                    condition = (
                        (df4['Channel'].str.lower() == row['Channel'].lower()) &
                        (row['AdStart'] >= df4['StartTime']) &
                        (row['AdEnd'] <= df4['EndTime']) &
                        (df4['Day'] == row['Day'])
                    )
                    if condition.any():
                        matching_rate = df4.loc[condition, 'Rate'].values[0]
                        df2.at[index, 'RPM'] = matching_rate

            # Apply zero rates for channels in df5 before applying special rates
            if not df5.empty:
                # Convert df5['Channel'] to lowercase to match df2['Channel']
                channels_in_df5 = df5['Channel'].str.lower().unique()
                df2['Channel_lower'] = df2['Channel'].str.lower()

                # Set RPM to zero for channels in df5
                df2.loc[df2['Channel_lower'].isin(channels_in_df5), 'RPM'] = 0

                # Apply special rates from df5
                for index, row in df2.iterrows():
                    condition = (
                        (df5['Channel'].str.lower() == row['Channel'].lower()) &
                        (df5['programName'] == row['programName']) 
                    )
                    if condition.any():
                        matching_rate = df5.loc[condition, 'Rate'].values[0]
                        df2.at[index, 'RPM'] = matching_rate
                    # Remove the temporary 'Channel_lower' column
                df2.drop(columns=['Channel_lower'], inplace=True)

            # Apply specific hardcoded conditions
            condition = (
                (df2['Channel'] == 'ARY DIGITAL') &
                (df2['TransmissionHour'] >= 19) &
                (df2['TransmissionHour'] <= 22) &
                (df2['programName'] == 'Jeeto Pakistan')
            )
            df2.loc[condition, 'RPM'] = 255500

            original_filename = secure_filename(file2.filename)
            output_path = os.path.normpath(os.path.join(app.config['UPLOAD_FOLDER'], original_filename))
            df2.to_excel(output_path, index=False) 
            
            return send_file(output_path, as_attachment=True, download_name=original_filename)
        
        flash("One or both files are not allowed types.")
        return redirect(request.url)

    except Exception as e:
        print(f"Error occurred: {e}")
        flash(f"An error occurred during processing: {e}")
        return redirect('/') 

if __name__ == "__main__":
    app.run(debug=True, port=5000, host='0.0.0.0')
