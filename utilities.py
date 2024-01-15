import pandas as pd
from scipy.stats import linregress

class AnalysisUtilities:
    
    @staticmethod
    def getfile_name():
        return 'LC2-032_KCP1 pHL-YEMK DC 20231030.xlsx'

    @staticmethod
    def getsheet1_name():
        return "Samples"

    @staticmethod
    def getsheet2_name():
        return "High Controls"
    
    @staticmethod
    def get_new_sheet_name():
        return "Analysis"

    @staticmethod
    def get_old_column_names(combined_df):
        # Extract only the column names from the combined_df DataFrame
        old_column_name_list = combined_df.columns.tolist()
        
        return old_column_name_list

    @staticmethod
    def get_new_column_names():

        new_column_names_list = ["pHL_VL2_BL1","yemk_vl2_bl1","relative_well_number","slope_corrected_phl_vl2_bl1","slope_corrected_yemk_vl2_bl1","cutoff_PHL_VL2_BL1_below_cuttoff","cutoff_yemk_vl2_bl1_below_cuttoff","phl_z_score","yemk_z_score","live_z_score","hits_phl_z_score","hits_yemk_z_score","hits_live_z_score"]

        return new_column_names_list 

    @staticmethod
    def get_renamed_column_names():
        
        renamed_column_names_list = ["well_number","total_count","phl_count","yemk_count","live_percentage","dead_percentage","phl_vl2","phl_bl1","yemk_vl2","yemk_bl1"]

        return renamed_column_names_list

    @staticmethod
    def prepare_analysis_df(file_path, sheet1, sheet2):
        # Create DataFrames for "Samples" and "High Controls" sheets
        samples_df = pd.read_excel(file_path, sheet_name=sheet1)
        high_controls_df = pd.read_excel(file_path, sheet_name=sheet2)

        # Combine the DataFrames into one
        combined_df = pd.concat([samples_df, high_controls_df], ignore_index=True)

        # Remove empty rows
        combined_df = combined_df.dropna(how='all')

        # Convert values in "X1" column to lowercase and remove rows where the value is "mean" or "sd"
        combined_df = combined_df[~combined_df[combined_df.columns[0]].str.lower().isin(['mean', 'sd'])]

        return combined_df

    @staticmethod
    def rewrite_column_names(combined_df, renamed_column_names_list, new_column_names_list):
        analysis_df = combined_df.copy()  # Create a copy to avoid modifying the original DataFrame

        # Rename existing columns
        for old_name, new_name in zip(renamed_column_names_list, new_column_names_list):
            analysis_df.rename(columns={old_name: new_name}, inplace=True)

        # Add new empty columns
        for new_name in new_column_names_list:
            analysis_df[new_name] = ''

        return analysis_df


    @staticmethod
    def write_analysis_sheet(analysis_df, file_path, new_sheet_name):
        # Check if the "Analysis" sheet already exists and delete it
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            if 'Analysis' in writer.sheets:
                writer.book.remove(writer.sheets['Analysis'])

            # Write the DataFrame to a new sheet named "Analysis"
            analysis_df.to_excel(writer, sheet_name='Analysis', index=False)

    @staticmethod
    def calculate_pHL_VL2_BL1(analysis_df):
        """
        Calculate the values for the 'pHL_VL2_BL1' column in the analysis DataFrame.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - pd.DataFrame: The analysis DataFrame with the 'pHL_VL2_BL1' column calculated.
        """
        analysis_df['pHL_VL2_BL1'] = analysis_df['phl_vl2'] / analysis_df['phl_bl1']
        return analysis_df
    
    @staticmethod
    def calculate_yemk_vl2_bl1(analysis_df):
        """
        Calculate the values for the 'yemk_vl2_bl1' column in the analysis DataFrame.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - pd.DataFrame: The analysis DataFrame with the 'yemk_vl2_bl1' column calculated.
        """
        analysis_df['yemk_vl2_bl1'] = analysis_df['yemk_vl2'] / analysis_df['yemk_bl1']
        return analysis_df
    
    @staticmethod
    def calculate_relative_well_number(analysis_df):
        """
        Calculate the values for the 'relative_well_number' column in the analysis DataFrame.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - pd.DataFrame: The analysis DataFrame with the 'relative_well_number' column calculated.
        """
        analysis_df['relative_well_number'] = range(1, len(analysis_df) + 1)
        return analysis_df
    
    @staticmethod
    def calculate_slope(analysis_df, x_column, y_column):
        """
        Calculate the slope of 'pHL_VL2_BL1' vs 'relative_well_number' columns.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - float: The slope of the linear regression.
        """
        # Perform linear regression
        slope, _, _, _, _ = linregress(analysis_df['relative_well_number'], analysis_df['pHL_VL2_BL1'])

        return slope

    @staticmethod
    def calculate_slope_corrected_phl_vl2_bl1(analysis_df, slope_phl_vl2_phl_bl1):
        """
        Calculate the values for the 'slope_corrected_phl_vl2_bl1' column.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.
        - slope_phl_vl2_phl_bl1 (float): The calculated slope.

        Returns:
        - pd.DataFrame: The analysis DataFrame with the 'slope_corrected_phl_vl2_bl1' column calculated.
        """
        analysis_df['slope_corrected_phl_vl2_bl1'] = analysis_df['pHL_VL2_BL1'] - (analysis_df['relative_well_number'] * slope_phl_vl2_phl_bl1)
        return analysis_df
    
    @staticmethod
    def get_slope_X_column_name():
        pass

    @staticmethod
    def get_slope_Y_Column_name():
        pass