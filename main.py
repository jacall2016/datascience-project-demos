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
    removed_columns_names_list = AnalysisUtilities.remove_columns_names_list()

    # Prepare analysis DataFrame
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
    AnalysisUtilities.export_All_Plates_YEMK_pHL_Live(analysis_df, 'root/All_P_YEMK_pHL_Live.xlsx','All_P_YEMK_pHL_Live')

    #write All hits excel file exclude all rows that don't have a hit in at least 1 of the three z numbers
    AnalysisUtilities.export_All_hits(analysis_df, 'root/All_hits.xlsx','All_hits')

    # Write analysis sheet
    AnalysisUtilities.write_analysis_sheet(analysis_df, file_path, new_sheet_name)

if __name__ == "__main__":
    main()
