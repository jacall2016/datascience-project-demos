import pandas as pd
from scipy.stats import linregress

class AnalysisUtilities:
    @staticmethod
    def prepare_analysis_df(file_path):
        # Create DataFrames for "Samples" and "High Controls" sheets
        samples_df = pd.read_excel(file_path, sheet_name='Samples')
        high_controls_df = pd.read_excel(file_path, sheet_name='High Controls')

        # Combine the DataFrames into one
        combined_df = pd.concat([samples_df, high_controls_df], ignore_index=True)

        # Remove empty rows
        combined_df = combined_df.dropna(how='all')

        # Convert values in "X1" column to lowercase and remove rows where the value is "mean" or "sd"
        combined_df = combined_df[~combined_df['X1'].str.lower().isin(['mean', 'sd'])]

        # Rename existing columns
        analysis_df = combined_df.rename(columns={
            'X1': 'well_number',
            'Count': 'total_count',
            'Live/Cells/Singlet.Cells/pHL.|.Count': 'phl_count',
            'Live/Cells/Singlet.Cells/YEMK.|.Count': 'yemk_count',
            'Live.|.Freq..of.Total.(%)': 'live_percentage',
            'Dead.|.Freq..of.Total.(%)': 'dead_percentage',
            'Live/Cells/Singlet.Cells/pHL.|.Median.(VL2-H.::.VL2-H)': 'phl_vl2',
            'Live/Cells/Singlet.Cells/pHL.|.Median.(BL1-H.::.BL1-H)': 'phl_bl1',
            'Live/Cells/Singlet.Cells/YEMK.|.Median.(VL2-H.::.VL2-H)': 'yemk_vl2',
            'Live/Cells/Singlet.Cells/YEMK.|.Median.(BL1-H.::.BL1-H)': 'yemk_bl1'
        })

        # Add new empty columns
        analysis_df['pHL_VL2_BL1'] = ''
        analysis_df['yemk_vl2_bl1'] = ''
        analysis_df['relative_well_number'] = ''
        analysis_df['slope_corrected_phl_vl2_bl1'] = ''
        analysis_df['slope_corrected_yemk_vl2_bl1'] = ''
        analysis_df['cutoff_PHL_VL2_BL1_below_cuttoff'] = ''
        analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'] = ''
        analysis_df['phl_z_score'] = ''
        analysis_df['yemk_z_score'] = ''
        analysis_df['live_z_score'] = ''
        analysis_df['hits_phl_z_score'] = ''
        analysis_df['hits_yemk_z_score'] = ''
        analysis_df['hits_live_z_score'] = ''

        return analysis_df

    @staticmethod
    def write_analysis_sheet(analysis_df, file_path):
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
    def calculate_slope_phl_vl2_phl_bl1(analysis_df):
        """
        Calculate the slope of 'pHL_VL2_BL1' vs 'relative_well_number' columns.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - float: The slope of the linear regression.
        """
        # Perform linear regression
        slope, _, _, _, _ = linregress(analysis_df['relative_well_number'], analysis_df['pHL_VL2_BL1'])

        # Save the slope as a variable
        slope_phl_vl2_phl_bl1 = slope

        return slope_phl_vl2_phl_bl1

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