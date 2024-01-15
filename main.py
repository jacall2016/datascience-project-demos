import pandas as pd

from utilities import AnalysisUtilities

def main():
    # Specify the path to your Excel file
    file_path = 'LC2-032_KCP1 pHL-YEMK DC 20231030.xlsx'

    # Prepare analysis DataFrame
    analysis_df = AnalysisUtilities.prepare_analysis_df(file_path)

    # Calculate 'pHL_VL2_BL1' column
    analysis_df = AnalysisUtilities.calculate_pHL_VL2_BL1(analysis_df)

    # Calculate 'yemk_vl2_bl1' column
    analysis_df = AnalysisUtilities.calculate_yemk_vl2_bl1(analysis_df)

    # Calculate 'relative_well_number' column
    analysis_df = AnalysisUtilities.calculate_relative_well_number(analysis_df)

    # Write analysis sheet
    AnalysisUtilities.write_analysis_sheet(analysis_df, file_path)

if __name__ == "__main__":
    main()
