from utilities import AnalysisUtilities
from flask import Flask, render_template, request, send_file, abort
import os

app = Flask(__name__)

def generate_files_phl_bl1_yemk_vl1(uploaded_file_path, sheet1, sheet2, final_sheet):
    
    # Specify the path to your Excel file
    file_name = AnalysisUtilities.getfile_name(uploaded_file_path)
    file_path = "uploads/" + file_name + ".xlsx"
    
    #get functions to retreave the desired data
    renamed_column_names_list = AnalysisUtilities.get_renamed_column_names()
    new_column_names_list = AnalysisUtilities.get_new_column_names()
    sheet1_name = AnalysisUtilities.getsheet1_name(sheet1)
    sheet2_name = AnalysisUtilities.getsheet2_name(sheet2)
    new_sheet_name = AnalysisUtilities.get_new_sheet_name(final_sheet)
    removed_columns_names_list = AnalysisUtilities.remove_columns_names_list()    
    
    # Prepare analysis DataFrame by combining data from Samples and High Controls and removing mean and SD rows
    combined_df = AnalysisUtilities.prepare_analysis_df(file_path, sheet1_name, sheet2_name, removed_columns_names_list)

    # get current combined_df column names
    old_column_name_list = AnalysisUtilities.get_old_column_names(combined_df)

    # rewrite columns and add new columns
    analysis_df = AnalysisUtilities.rewrite_column_names(combined_df, old_column_name_list,  renamed_column_names_list, new_column_names_list)

    # Calculate 'pHL_VL2_BL1' column
    analysis_df = AnalysisUtilities.calculate_pHL_VL2_BL1(analysis_df)

    # Calculate 'yemk_vl2_bl1' column
    analysis_df = AnalysisUtilities.calculate_yemk_vl2_bl1(analysis_df)

    # Calculate 'relative_well_number' column
    analysis_df = AnalysisUtilities.calculate_relative_well_number(analysis_df)

    # calculate the slopes
    slope_phl_vl2_phl_bl1 = AnalysisUtilities.calculate_slope_phl_vl2_phl_bl1(analysis_df)
    slope_yemk_vl2_yemk_bl1 = AnalysisUtilities.calculate_slope_yemk_vl2_bl1(analysis_df)

    # Calculate and populate the corrected slope for ph1_vl2_bl1
    analysis_df = AnalysisUtilities.calculate_slope_corrected_phl_vl2_bl1(analysis_df, slope_phl_vl2_phl_bl1)
    analysis_df = AnalysisUtilities.calculate_slope_corrected_yemk_vl2_bl1(analysis_df, slope_yemk_vl2_yemk_bl1)

    # calculate the means
    mean_phl_vl2_phl_bl1 = AnalysisUtilities.calculate_mean_phl_vl2_phl_bl1(analysis_df)
    mean_yemk_vl2_yemk_bl1 = AnalysisUtilities.calculate_mean_yemk_vl2_yemk_bl1(analysis_df)

    # calculate the standard deviations
    sd_phl_vl2_phl_bl1 = AnalysisUtilities.calculate_sd_phl_vl2_phl_bl1(analysis_df)
    sd_yemk_vl2_yemk_bl1 = AnalysisUtilities.calculate_sd_yemk_vl2_yemk_bl1(analysis_df)

    # calculate the cuttoffs
    cuttoff_phl_vl2_phl_bl1 = AnalysisUtilities.calculate_cuttoff_phl_vl2_phl_bl1(mean_phl_vl2_phl_bl1, sd_phl_vl2_phl_bl1)
    cuttoff_yemk_vl2_yemk_bl1 = AnalysisUtilities.calculate_cuttoff_yemk_vl2_yemk_bl1(mean_yemk_vl2_yemk_bl1, sd_yemk_vl2_yemk_bl1)

    # populate the cutoff_PHL_VL2_BL1_below_cuttoff if above the cutoff don't include
    analysis_df = AnalysisUtilities.populate_cutoff_PHL_VL2_BL1_below_cuttoff(analysis_df, cuttoff_phl_vl2_phl_bl1)
    analysis_df = AnalysisUtilities.populate_cutoff_yemk_vl2_bl1_below_cuttoff(analysis_df, cuttoff_yemk_vl2_yemk_bl1)

    #calculate correct mean or phl_vl2_phl_bl1 and yemk_vl2_yemk_bl1
    corrected_mean_phl_vl2_phl_bl1 = AnalysisUtilities.calculate_corrected_mean_phl_vl2_phl_bl1(analysis_df)
    corrected_mean_yemk_vl2_yemk_bl1 = AnalysisUtilities.calculate_corrected_mean_yemk_vl2_yemk_bl1(analysis_df)

    # calculate correct standard deviation for ph1 and yemk
    corrected_sd_phl_vl2_phl_bl1 = AnalysisUtilities.calculate_corrected_sd_phl_vl2_phl_bl1(analysis_df)
    corrected_sd_yemk_vl2_yemk_bl1 = AnalysisUtilities.calculate_corrected_sd_yemk_vl2_yemk_bl1(analysis_df)

    #populate ph1 and yemk z score
    analysis_df = AnalysisUtilities.populate_phl_z_score(analysis_df, corrected_mean_phl_vl2_phl_bl1, corrected_sd_phl_vl2_phl_bl1)
    analysis_df = AnalysisUtilities.populate_yemk_z_score(analysis_df, corrected_mean_yemk_vl2_yemk_bl1, corrected_sd_yemk_vl2_yemk_bl1)

    #calculate live mean and standard deviation
    live_mean = AnalysisUtilities.calculate_live_mean(analysis_df)
    live_sd = AnalysisUtilities.calculate_live_sd(analysis_df)

    #populate live z score
    analysis_df = AnalysisUtilities.populate_live_z_score(analysis_df, live_mean, live_sd)

    #populate the hits
    analysis_df = AnalysisUtilities.populate_hits_phl_z_score(analysis_df)
    analysis_df = AnalysisUtilities.populate_hits_yemk_z_score(analysis_df)
    analysis_df = AnalysisUtilities.populate_hits_live_z_score(analysis_df)

    #write the All_Plates_YEMK_pHL_Live Excel file
    AnalysisUtilities.export_All_Plates_YEMK_pHL_Live(analysis_df, 'downloads/All_P_YEMK_pHL_Live.xlsx','All_P_YEMK_pHL_Live')

    #write All hits excel file exclude all rows that don't have a hit in at least 1 of the three z numbers
    AnalysisUtilities.export_All_hits(analysis_df, 'downloads/All_hits.xlsx','All_hits')

    # Write analysis sheet
    AnalysisUtilities.write_analysis_sheet(analysis_df, "downloads/LC2-032_KCP1 pHL-YEMK DC 20231030.xlsx", new_sheet_name)
    
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/combine_sheets', methods=['POST'])
def combine_sheets():
    # Handle file and form data
    input_file = request.files['input_file']
    file_type = request.form['selected_radio_id']

    validate(input_file,file_type)

    sheet1 = "Samples"
    sheet2 = "High Controls"
    final_sheet = "Analysis"

    # Save the uploaded Excel file
    uploaded_file_path = os.path.join("uploads", input_file.filename)
    input_file.save(uploaded_file_path)

    if file_type == "pl1":
        # Generate the Excel phl_bl1_yemk_vl1 files using your processing function
        generate_files_phl_bl1_yemk_vl1(uploaded_file_path, sheet1, sheet2, final_sheet)
    elif file_type == "XXXX":
        # Generate the Excel XXXX files using your processing function
        generate_files_phl_bl1_yemk_vl1(uploaded_file_path, sheet1, sheet2, final_sheet)

    # Clear the contents of the "uploads" folder
    clear_uploads_folder()

    # Provide download links on the webpage
    return render_template('download.html', 
                           input_file_name=input_file.filename,
                           output_file_names=["LC2-032_KCP1 pHL-YEMK DC 20231030.xlsx", "All_P_YEMK_pHL_Live.xlsx", "All_hits.xlsx"])

@app.route('/download_file')
def download_file():
    filename = request.args.get('filename')
    file_path = os.path.join("downloads", filename)

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        abort(404)

def clear_uploads_folder():
    # Define the path to the "uploads" folder
    uploads_folder_path = "uploads"

    # Check if the folder exists
    if os.path.exists(uploads_folder_path):
        # Iterate over the files in the folder and remove them
        for file_name in os.listdir(uploads_folder_path):
            file_path = os.path.join(uploads_folder_path, file_name)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")
    else:
        print("The 'uploads' folder does not exist.")

def validate(input_file,file_type):
    pass

if __name__ == "__main__":
    app.run(debug=True, port=5000)