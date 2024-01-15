import pandas as pd

from utilities import AnalysisUtilities

def main():
    # Specify the path to your Excel file
    file_name = AnalysisUtilities.getfile_name()
    file_path = "root/" + file_name

    #get functions to retreave the desired data
    renamed_column_names_list = AnalysisUtilities.get_renamed_column_names()
    new_column_names_list = AnalysisUtilities.get_new_column_names()
    sheet1_name = AnalysisUtilities.getsheet1_name()
    sheet2_name = AnalysisUtilities.getsheet2_name()
    new_sheet_name = AnalysisUtilities.get_new_sheet_name()

    # Prepare analysis DataFrame
    combined_df = AnalysisUtilities.prepare_analysis_df(file_path, sheet1_name, sheet2_name)

    old_column_name_list = AnalysisUtilities.get_old_column_names(combined_df)

    analysis_df = AnalysisUtilities.rewrite_column_names(combined_df, renamed_column_names_list, new_column_names_list)

    # Calculate 'pHL_VL2_BL1' column
    analysis_df = AnalysisUtilities.calculate_pHL_VL2_BL1(analysis_df)

    # Calculate 'yemk_vl2_bl1' column
    #analysis_df = AnalysisUtilities.calculate_yemk_vl2_bl1(analysis_df)

    # Calculate 'relative_well_number' column
    #analysis_df = AnalysisUtilities.calculate_relative_well_number(analysis_df)

    #slope_phl_vl2_phl_bl1 = AnalysisUtilities.calculate_slope(analysis_df, AnalysisUtilities.get_slope_X_column_name(), AnalysisUtilities.get_slope_Y_Column_name())

    # Write analysis sheet
    #AnalysisUtilities.write_analysis_sheet(analysis_df, file_path, new_sheet_name)

if __name__ == "__main__":
    main()
