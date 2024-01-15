import pandas as pd
from scipy.stats import linregress
from datetime import datetime
import os
from flask import Flask, render_template, request, redirect

class AnalysisUtilities:
    
    @staticmethod
    def getfile_name(uploaded_file_path):
        # Extract the base name (file name with extension) from the full path
        file_name_with_extension = os.path.basename(uploaded_file_path)
        
        # Remove the extension to get only the file name
        file_name_without_extension = os.path.splitext(file_name_with_extension)[0]

        return file_name_without_extension

    @staticmethod
    def getsheet1_name(sheet1):
        return sheet1

    @staticmethod
    def getsheet2_name(sheet2):
        return sheet2
    
    @staticmethod
    def get_new_sheet_name(final_sheet):
        return final_sheet

    @staticmethod
    def remove_columns_names_list():
        return ['mean', 'sd']

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
    def prepare_analysis_df(file_path, sheet1, sheet2, remove_columns_names):

        # Create DataFrames for "Samples" and "High Controls" sheets
        samples_df = pd.read_excel(file_path, sheet_name=sheet1)

        high_controls_df = pd.read_excel(file_path, sheet_name=sheet2)

        # Combine the DataFrames into one
        combined_df = pd.concat([samples_df, high_controls_df], ignore_index=True)

        # Remove empty rows
        combined_df = combined_df.dropna(how='all')

        # Convert values in "X1" column to lowercase and remove rows where the value is "mean" or "sd"
        combined_df = combined_df[~combined_df[combined_df.columns[0]].str.lower().isin(remove_columns_names)]

        return combined_df

    @staticmethod
    def rewrite_column_names(combined_df, old_column_name_list, renamed_column_names_list, new_column_names_list):
        analysis_df = combined_df.copy()  # Create a copy to avoid modifying the original DataFrame

        # Rename existing columns
        for old_name, new_name in zip(old_column_name_list, renamed_column_names_list):
            analysis_df.rename(columns={old_name: new_name}, inplace=True)

        # Add new empty columns
        for new_name in new_column_names_list:
            analysis_df[new_name] = ''

        return analysis_df

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

        return slope
    
    @staticmethod
    def calculate_slope_yemk_vl2_bl1(analysis_df):
        """
        Calculate the slope of 'yemk_vl2_bl1' vs 'relative_well_number' columns.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - float: The slope of the linear regression.
        """
        # Perform linear regression
        slope, _, _, _, _ = linregress(analysis_df['relative_well_number'], analysis_df['yemk_vl2_bl1'])

        return slope

    @staticmethod
    def calculate_mean_phl_vl2_phl_bl1(analysis_df):
        """
        Calculate the mean of the 'pHL_VL2_BL1' column.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - float: The mean of the 'pHL_VL2_BL1' column.
        """
        # Calculate the mean
        mean_phl_vl2_phl_bl1 = analysis_df['slope_corrected_phl_vl2_bl1'].mean()

        return mean_phl_vl2_phl_bl1

    @staticmethod
    def calculate_mean_yemk_vl2_yemk_bl1(analysis_df):
        """
        Calculate the mean of the 'yemk_vl2_bl1' column.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - float: The mean of the 'yemk_vl2_bl1' column.
        """
        # Calculate the mean
        mean_yemk_vl2_yemk_bl1 = analysis_df['slope_corrected_yemk_vl2_bl1'].mean()

        return mean_yemk_vl2_yemk_bl1

    @staticmethod
    def calculate_sd_phl_vl2_phl_bl1(analysis_df):
        """
        Calculate the standard deviation of the 'pHL_VL2_BL1' column.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - float: The standard deviation of the 'pHL_VL2_BL1' column.
        """
        # Calculate the standard deviation
        sd_phl_vl2_phl_bl1 = analysis_df['slope_corrected_phl_vl2_bl1'].std()

        return sd_phl_vl2_phl_bl1

    @staticmethod
    def calculate_sd_yemk_vl2_yemk_bl1(analysis_df):
        """
        Calculate the standard deviation of the 'yemk_vl2_bl1' column.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.

        Returns:
        - float: The standard deviation of the 'yemk_vl2_bl1' column.
        """
        # Calculate the standard deviation
        sd_yemk_vl2_yemk_bl1 = analysis_df['slope_corrected_yemk_vl2_bl1'].std()

        return sd_yemk_vl2_yemk_bl1

    @staticmethod
    def calculate_cuttoff_phl_vl2_phl_bl1(mean_phl_vl2_phl_bl1, sd_phl_vl2_phl_bl1):

        cutoff = mean_phl_vl2_phl_bl1 + (1.5 * sd_phl_vl2_phl_bl1)

        return cutoff

    @staticmethod
    def calculate_cuttoff_yemk_vl2_yemk_bl1(mean_yemk_vl2_yemk_bl1, sd_yemk_vl2_yemk_bl1):
        
        cutoff = mean_yemk_vl2_yemk_bl1 + (1.5 * sd_yemk_vl2_yemk_bl1)

        return cutoff

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
    def calculate_slope_corrected_yemk_vl2_bl1(analysis_df, slope_yemk_vl2_yemk_bl1):
        """
        Calculate the values for the 'slope_corrected_yemk_vl2_bl1' column.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.
        - slope_yemk_vl2_yemk_bl1 (float): The calculated slope.

        Returns:
        - pd.DataFrame: The analysis DataFrame with the 'slope_corrected_yemk_vl2_bl1' column calculated.
        """
        analysis_df['slope_corrected_yemk_vl2_bl1'] = analysis_df['yemk_vl2_bl1'] - (analysis_df['relative_well_number'] * slope_yemk_vl2_yemk_bl1)
        return analysis_df
    
    @staticmethod
    def populate_cutoff_PHL_VL2_BL1_below_cuttoff(analysis_df, cuttoff_phl_vl2_phl_bl1):
        """
        Populate the 'cutoff_PHL_VL2_BL1_below_cuttoff' column based on the cutoff value.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.
        - cuttoff_phl_vl2_phl_bl1 (float): The cutoff value.

        Returns:
        - pd.DataFrame: The analysis DataFrame with the 'cutoff_PHL_VL2_BL1_below_cuttoff' column populated.
        """
        # Create a copy of the 'pHL_VL2_BL1' column
        analysis_df['cutoff_PHL_VL2_BL1_below_cuttoff'] = analysis_df['slope_corrected_phl_vl2_bl1']

        # Replace values with None where the condition is not met
        analysis_df.loc[analysis_df['slope_corrected_phl_vl2_bl1'] > cuttoff_phl_vl2_phl_bl1, 'cutoff_PHL_VL2_BL1_below_cuttoff'] = None

        return analysis_df

    @staticmethod
    def populate_cutoff_yemk_vl2_bl1_below_cuttoff(analysis_df, cuttoff_yemk_vl2_yemk_bl1):
        """
        Populate the 'cutoff_yemk_vl2_bl1_below_cuttoff' column based on the cutoff value.

        Parameters:
        - analysis_df (pd.DataFrame): The analysis DataFrame.
        - cuttoff_yemk_vl2_yemk_bl1 (float): The cutoff value.

        Returns:
        - pd.DataFrame: The analysis DataFrame with the 'cutoff_yemk_vl2_bl1_below_cuttoff' column populated.
        """
        # Create a copy of the 'yemk_vl2_bl1' column
        analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'] = analysis_df['slope_corrected_yemk_vl2_bl1']

        # Replace values with None where the condition is not met
        analysis_df.loc[analysis_df['slope_corrected_yemk_vl2_bl1'] > cuttoff_yemk_vl2_yemk_bl1, 'cutoff_yemk_vl2_bl1_below_cuttoff'] = None

        return analysis_df
    
    @staticmethod
    def calculate_corrected_mean_phl_vl2_phl_bl1(analysis_df):
        
        # calculate the corrected_mean_phl_vl2_phl_bl1 by the cutoff_PHL_VL2_BL1_below_cuttoff 
        corrected_mean_phl_vl2_phl_bl1 = analysis_df['cutoff_PHL_VL2_BL1_below_cuttoff'].mean()

        return corrected_mean_phl_vl2_phl_bl1

    @staticmethod
    def calculate_corrected_mean_yemk_vl2_yemk_bl1(analysis_df):
        
        # calculate the corrected_mean_yemk_vl2_yemk_bl1 by the cutoff_yemk_vl2_bl1_below_cuttoff 
        corrected_mean_yemk_vl2_yemk_bl1 = analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'].mean()

        return corrected_mean_yemk_vl2_yemk_bl1

    @staticmethod
    def calculate_corrected_sd_phl_vl2_phl_bl1(analysis_df):
        
        # calculate the corrected_mean_phl_vl2_phl_bl1 by the cutoff_PHL_VL2_BL1_below_cuttoff 
        corrected_sd_phl_vl2_phl_bl1 = analysis_df['cutoff_PHL_VL2_BL1_below_cuttoff'].std()

        return corrected_sd_phl_vl2_phl_bl1

    @staticmethod
    def calculate_corrected_sd_yemk_vl2_yemk_bl1(analysis_df):
        
        # calculate the corrected_mean_yemk_vl2_yemk_bl1 by the cutoff_yemk_vl2_bl1_below_cuttoff 
        corrected_sd_yemk_vl2_yemk_bl1 = analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'].std()

        return corrected_sd_yemk_vl2_yemk_bl1
    
    @staticmethod
    def populate_phl_z_score(analysis_df, corrected_mean_phl_vl2_phl_bl1, corrected_sd_phl_vl2_phl_bl1):

        if not analysis_df['cutoff_PHL_VL2_BL1_below_cuttoff'].isna().all():
            analysis_df['phl_z_score'] = (analysis_df['cutoff_PHL_VL2_BL1_below_cuttoff'] - corrected_mean_phl_vl2_phl_bl1)/corrected_sd_phl_vl2_phl_bl1

        return analysis_df
    
    @staticmethod
    def populate_yemk_z_score(analysis_df, corrected_mean_yemk_vl2_yemk_bl1, corrected_sd_yemk_vl2_yemk_bl1):

        if not analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'].isna().all():
            analysis_df['yemk_z_score'] = (analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'] - corrected_mean_yemk_vl2_yemk_bl1)/corrected_sd_yemk_vl2_yemk_bl1

        return analysis_df
    
    @staticmethod
    def calculate_live_mean(analysis_df):
        
        live_mean = analysis_df['live_percentage'].mean()

        return live_mean

    @staticmethod
    def calculate_live_sd(analysis_df):

        live_sd = analysis_df['live_percentage'].std()

        return live_sd
    
    @staticmethod
    def populate_live_z_score(analysis_df, live_mean, live_sd):
        
        if not analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'].isna().all():
            analysis_df['live_z_score'] = (analysis_df['live_percentage'] - live_mean)/live_sd

        return analysis_df
    
    @staticmethod
    def populate_hits_phl_z_score(analysis_df):
        # Use boolean indexing to filter rows based on conditions
        condition = (analysis_df['cutoff_PHL_VL2_BL1_below_cuttoff'].notna()) & (analysis_df['phl_z_score'] < -5)

        # Populate 'hits_phl_z_score' based on the condition
        analysis_df.loc[condition, 'hits_phl_z_score'] = analysis_df.loc[condition, 'phl_z_score']

        return analysis_df
    
    @staticmethod
    def populate_hits_yemk_z_score(analysis_df):
        # Use boolean indexing to filter rows based on conditions
        condition = (analysis_df['cutoff_yemk_vl2_bl1_below_cuttoff'].notna()) & (analysis_df['yemk_z_score'] < -5)

        # Populate 'hits_yemk_z_score' based on the condition
        analysis_df.loc[condition, 'hits_yemk_z_score'] = analysis_df.loc[condition, 'yemk_z_score']

        return analysis_df

    @staticmethod
    def populate_hits_live_z_score(analysis_df):
        # Use boolean indexing to filter rows based on conditions
        condition = analysis_df['live_z_score'] < -5

        # Populate 'hits_live_z_score' based on the condition
        analysis_df.loc[condition, 'hits_live_z_score'] = analysis_df.loc[condition, 'live_z_score']

        return analysis_df
    
    @staticmethod
    def export_All_Plates_YEMK_pHL_Live(analysis_df, excel_file_path, base_sheet_name):
        # Select relevant columns from analysis_df
        selected_columns = ['well_number', 'phl_z_score', 'yemk_z_score', 'live_z_score']
        export_df = analysis_df[selected_columns]

        # Format datetime for readability
        formatted_datetime = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

        # Append formatted datetime to the base sheet name
        sheet_name = f"{base_sheet_name}_{formatted_datetime}"
        # Check if the file exists
        file_exists = os.path.isfile(excel_file_path)

        # Create the file if it doesn't exist
        if not file_exists:
            export_df.to_excel(excel_file_path, sheet_name=sheet_name, index=False)
        else:
            # Export the selected columns to a new sheet if the file exists
            with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
                export_df.to_excel(writer, sheet_name=sheet_name, index=False)
        return export_df
    
    @staticmethod
    def export_All_hits(analysis_df, excel_file_path, base_sheet_name):
        
        # Select relevant columns from analysis_df
        selected_columns = ['well_number', 'hits_phl_z_score', 'hits_yemk_z_score', 'hits_live_z_score']

        # Filter out rows where all specified columns don't have any numeric values
        export_df = analysis_df[analysis_df[selected_columns].applymap(lambda x: isinstance(x, (int, float))).any(axis=1)]

        # Include only the selected columns in export_df
        export_df = export_df[selected_columns]

        # Format datetime for readability
        formatted_datetime = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

        # Append formatted datetime to the base sheet name
        sheet_name = f"{base_sheet_name}_{formatted_datetime}"
        # Check if the file exists
        file_exists = os.path.isfile(excel_file_path)

        # Create the file if it doesn't exist
        if not file_exists:
            export_df.to_excel(excel_file_path, sheet_name=sheet_name, index=False, engine='openpyxl')
        else:
            # Export the selected columns to a new sheet if the file exists
            with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
                export_df.to_excel(writer, sheet_name=sheet_name, index=False)

        return export_df
    
    @staticmethod
    def write_analysis_sheet(analysis_df, file_path, new_sheet_name):
        # Extract the file name from the original file path
        original_file_name = os.path.basename(file_path)

        # Construct the new file name by adding "downloads/" to the front
        new_file_path = os.path.join("downloads", f"{original_file_name}")

        # Check if the "Analysis" sheet already exists and delete it in the new file
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            if new_sheet_name in writer.sheets:
                writer.book.remove(writer.sheets[new_sheet_name])

            # Write the DataFrame to a new sheet named "Analysis" in the new file
            analysis_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

        return new_file_path
