import pandas as pd
import numpy as np
import pandas as pd
import datetime
import openpyxl
from pandas_market_calendars import get_calendar
from dateutil.relativedelta import relativedelta
import os
from datetime import timedelta
import glob
import xlrd
import re
from pandas.tseries.offsets import DateOffset
import math

class Forward_Curve:
    
    def __init__(self, commodities):
        
        self.commodities = commodities
        
        start_date = pd.Timestamp('2013-01-01')
        end_date = pd.Timestamp('2024-09-10')

        '''
        
        Usefull to open only relevant files considering the commodity
        
        '''
        
        europe_holidays = pd.read_csv('.venv/Petroleum/Platts Calandar/eu_holidays.csv')
        us_holidays = pd.read_csv('.venv/Petroleum/Platts Calandar/us_holidays.csv')
        singapore_holidays = pd.read_csv('.venv/Petroleum/Platts Calandar/singapore_holidays.csv')
        # Generate the business days DataFrame for the specified month
        business_days = pd.bdate_range(start=start_date, end=end_date)
        self.europe = business_days[~business_days.isin(europe_holidays['Date'])]
        self.us = business_days[~business_days.isin(us_holidays['Date'])]
        self.singapore = business_days[~business_days.isin(singapore_holidays['Date'])]
        
    def scale(self):
        
        Commodities = {}
        
        for commodity, product, location in commodities:
            
            if location == "Europe":
                dates = self.europe
            elif location == "US":
                dates = self.us
            elif location == "Singapore":
                dates = self.singapore
            else:
                print("No location specified")
            
            Dates = {}
            
            for date in dates:        
                
                #just serve for later line of code purpose below, formatting
                date_dt = pd.to_datetime(date, format='%d/%m/%Y')
                print("[SCALE] Commodity:", commodity)
                print("[SCALE] Date:", date_dt)
                
                file_path = self.find_valid_file_path(date_dt)
                
                
                if file_path is not None:
                
                    df = pd.read_excel(file_path, header=0, index_col=None, usecols=None, skiprows=0, nrows=None, dtype=None)

                    if df.empty:
                        print("df empty for:", date)
                    else:

                        # Expected column names
                        
                        expected_column_names = ['TRADE DATE', 'HUB', 'PRODUCT', 'STRIP', 'CONTRACT', 'CONTRACT TYPE', 'STRIKE', 'SETTLEMENT PRICE', 'NET CHANGE', 'EXPIRATION DATE', 'PRODUCT_ID']

                        # Check if the header matches expected column names
                        if not all(col in df.columns for col in expected_column_names):
                            print('switch')
                            # If not, set the next row as the header
                            df = pd.read_excel(file_path, header=1, index_col=None, usecols=None, skiprows=0, nrows=None, dtype=None)   
                    
                        if df.empty:
                            print('df is empty for this date:', date)
            
                        else:

                            forward_curve_df = self.generate_forward_curve(df, commodity, product, date_dt)
                            
                            if forward_curve_df.empty:
                                print(f"[SCALE] No processed forward curve for {commodity} at {date_dt}")
                            else:
                                Dates[date_dt] = forward_curve_df
                
                else:
                    print('No file path for this date:', date)
                
                    
            Commodities[commodity] = Dates

        return Commodities
 
    def generate_forward_curve(self, input_spreadsheet, commodity, product, date_ref):
        
        '''
            necessiate a date input, because need to macth correct month
            
            
            input_spreadsheet
            
            date ref is the date used to open the spreadsheet
            
        '''
        
        
        df = input_spreadsheet
        date_dt = date_ref
 
        # Create a new DataFrame with desired commodity and product
        df = df[((df['HUB'] == commodity) & (df['PRODUCT'] == product))]
        
        

        if df.empty:
            return df
            
        # Convert 'STRIP' column to datetime
        df['EXPIRATION DATE'] = pd.to_datetime(df['EXPIRATION DATE'], errors='coerce')

        # Check if the dates in 'STRIP' column are in US format (MM/DD/YYYY)
        us_date_format = df['EXPIRATION DATE'].dt.strftime('%d/%m/%Y').str.contains(r'\b\d{1,2}/\d{1,2}/\d{4}\b')

        # Convert US format (MM/DD/YYYY) dates to UK format (DD/MM/YYYY)
        df.loc[us_date_format, 'EXPIRATION DATE'] = pd.to_datetime(df.loc[us_date_format, 'EXPIRATION DATE'], format='%m/%d/%Y')

        # Calculate the difference in days between current date and 'EXPIRATION DATE'
        df['Days_Left'] = (df['EXPIRATION DATE'] - date_dt).dt.days
        df['Months_Left'] = round(df['Days_Left']/30, 1)
        
        df['Month Left diff'] = df['Months_Left'].diff().fillna(0)
        
        # Assuming df is your DataFrame
        df.reset_index(drop=True, inplace=True)
        
       
                # Custom round function for the first value
        def custom_round(number):
            if number - math.floor(number) == 0.5:
                return math.ceil(number)
            else:
                return round(number)

        '''
            INITIALISE LIQUID MONTH
        
        '''
        
        # Fetch the value



        # Sample DataFrame (assuming the DataFrame is already loaded as df)
        # Creating the 'contract' column
        df['liquid contract'] = 0  # Initialize the column

        # Initialize the first value of 'contract' column
        df.loc[0, 'liquid contract'] = custom_round(df.loc[0, 'Months_Left'])
        
        '''
        
        INITIALISE CONTRACT MONTH
        '''
        
        df['contract'] = 0  # Initialize the column
        
        value = df.loc[0, 'Months_Left']

        # Check if the value is greater than 0 and has a decimal part of 0
        if value > 0 and value == round(value, 1) and (value * 10) % 10 == 0:
            print("The value is greater than 0 and its decimal is 0.")
            print(value)
            df.loc[0, 'contract'] = math.floor(df.loc[0, 'Months_Left']) - 1
        else:
            df.loc[0, 'contract'] = math.floor(df.loc[0, 'Months_Left'])
        
        
        
        # Loop through the DataFrame and apply the logic
        for i in range(1, len(df)):
            month_diff = df.loc[i, 'Month Left diff']
           
            
            # Check if month_diff ends in 0.5
            if month_diff % 1 == 0.5:
                increment = math.floor(month_diff)
            else:
                increment = round(month_diff)
                
    
            # Set the contract value based on the increment
            df.loc[i, 'liquid contract'] = df.loc[i-1, 'liquid contract'] + increment
            df.loc[i, 'contract'] = df.loc[i-1, 'contract'] + increment

        
        
       
        df["Liquid Month"] = [num - 1 for num in df['liquid contract']]
        df["Contract Month"] = df['contract']
       
        
        forward_curve_df = pd.DataFrame({
            
            #"Current Date": df["TRADE DATE"],
            "Expiration Date": df['EXPIRATION DATE'],
            "Price": df["SETTLEMENT PRICE"],
            "Contract Month": df['Contract Month'],
            "Liquid Month": df["Liquid Month"],
            "Month Left diff": df["Month Left diff"],
            'Months Until Expiry': df['Months_Left'],
            
        })

        # Optionally, you can reset the options back to default
        pd.reset_option('display.max_rows')
        pd.reset_option('display.max_columns')
        
        print("Date:", date_dt)
        print("Commodity:", commodity)
        print(forward_curve_df)
        
        # Optionally, you can reset the options back to default
        pd.reset_option('display.max_rows')
        pd.reset_option('display.max_columns')

        return forward_curve_df
        
    def get_dataframe(self):
        # Dictionary to store final DataFrames
        nested_dict = self.scale()
        print("Scale done")
                # Dictionary to store final DataFrames
        final_dataframes = {}
        
        # Iterate over commodities
        for commodity, dates_dict in nested_dict.items():
            # Dictionary to store prices for each contract month
            contract_month_prices = {}
            # Dictionary to store prices for each liquid month
            liquid_month_prices = {}
            # Iterate over dates and their corresponding DataFrames
            for date, df in dates_dict.items():
                # Group by 'Contract Month' and get the corresponding price
                contract_month_grouped = df.groupby('Contract Month')['Price'].first()
                # Update contract_month_prices with prices for the current date
                contract_month_prices[date] = contract_month_grouped
                # Group by 'Liquid Month' and get the corresponding price
                liquid_month_grouped = df.groupby('Liquid Month')['Price'].first()
                # Update liquid_month_prices with prices for the current date
                liquid_month_prices[date] = liquid_month_grouped
                
                
            # Check if there are objects to concatenate
            if contract_month_prices:
                # Concatenate the dictionaries into DataFrames
                final_contract_month_df = pd.concat(contract_month_prices, axis=1)
                # Transpose the DataFrame to have dates as rows and contract months as columns
                final_contract_month_df = final_contract_month_df.T
                # Reset index to have date as a regular column and create a new integer index
                
                final_contract_month_df.insert(0, 'Date', final_contract_month_df.index)
                final_contract_month_df.reset_index(drop=True, inplace=True)
            else:
                # If no objects to concatenate, create an empty DataFrame
                final_contract_month_df = pd.DataFrame(columns=['Date'])

            if liquid_month_prices:
                # Concatenate the dictionaries into DataFrames
                final_liquid_month_df = pd.concat(liquid_month_prices, axis=1)
                # Transpose the DataFrame to have dates as rows and liquid months as columns
                final_liquid_month_df = final_liquid_month_df.T
                # Reset index to have date as a regular column and create a new integer index
                
                final_liquid_month_df.insert(0, 'Date', final_liquid_month_df.index)
                final_liquid_month_df.reset_index(drop=True, inplace=True)
            else:
                # If no objects to concatenate, create an empty DataFrame
                final_liquid_month_df = pd.DataFrame(columns=['Date'])
        
                
            # Add the final DataFrames to the dictionary
            final_dataframes[commodity] = {'Contract Month': final_contract_month_df, 'Liquid Month': final_liquid_month_df}

        # Print final DataFrames for each commodity
        for commodity, dfs in final_dataframes.items():
            print(f"Commodity: {commodity}")
            print("Contract Month:")
            print(dfs['Contract Month'])
            print("Liquid Month:")
            print(dfs['Liquid Month'])
        
        return final_dataframes

    def find_valid_file_path(self, date):
    
        '''
        Find a valid file path.

        If not valid, try one day before until the month or year no longer match.
        Then, try one day after the original date.
        If still doesn't work, give up and return None.
        '''

        formatted_date = date.strftime('%Y_%m_%d')
            
            # Search for both .xlsx and .xls files
        file_paths = glob.glob(f'/Users/tristanchorley/Documents/oil prices/icecleared_oil_{formatted_date}.*')
            
        for file_path in file_paths:
            if os.path.exists(file_path):
                if file_path.endswith('.xlsx'):
                    return file_path
                elif file_path.endswith('.xls'):
                    return file_path
            else:
                return None


commodities = [("3.5% FOB Rdam Bg", "Fuel Oil Futures", "Europe"), 
                ("Brent 1st Line", "Crude Futures", "Europe"), 
                ("380cst Sing", "Fuel Oil Futures", "Singapore"),
                ("180cst Sing", "Fuel Oil Futures", "Singapore"),
                ("Marine 0.5% FOB Sing (Platts)", "Fuel Oil Futures", "Singapore"),
                ("Marine 0.5% FOB Rdam Bg (Platts)", "Fuel Oil Futures", "Europe"),
                ("WTI 1st Line", "Crude Futures", "US"),
                ("Sing Mogas 92 Unl (Platts)", "Gasoline Futures", "Singapore"),
                ("Abu Dhabi", "Murban Crude Futures", "Singapore"),
                ("Argus Mars", "Crude Futures", "Singapore"),
                ("Dubai 1st Line", "Crude Futures", "Singapore"),
                ("Murban 1st Line", "Crude Futures", "Singapore"),
                ('Sing Gasoil 10ppm (Platts)', 'Gasoil Futures', 'Singapore'),
                ('Dated Brent', 'Crude Futures', 'Europe'),
                ("Dated Brent/Brent 1st Line", "Crude Diff Futures", "Europe"),
                ("TC2", "Freight Futures (USD)", "Singapore"),
                ("TD3C", "Freight Futures (USD)", "Singapore"),
                ("TC15", "Freight Futures (USD)", "Singapore"),
                ("North Sea", "Brent Crude Futures", "Europe"),
                ("TC20", "Freight Futures (USD)", "Singapore"),
                ("TD22", "Freight Futures (USD)", "Singapore"),
                ("Argus Eurobob Oxy FOB Rdam Bg", "Gasoline Futures", "Europe"),
               ("RBOB 1st Line", "Gasoline Futures", "US"),
               ("Middle East Mogas 92 FOB Arab Gulf (Platts)", "Gasoline Futures", "Singapore"),
               ("Diesel 10ppm FOB ARA Bg", "Diesel Futures", "Europe"),
               ("USGC ULSD", "Diesel Futures", "US"),
               ("ULSD 10ppm FOB Med Cg (Platts)", "Diesel Futures", "Singapore"),
               ("Jet FOB Rdam Bg (Platts)", "Jet Fuel Futures", "Europe"),
               ("Jet Mid East FOB Arab Gulf (Platts)", "Jet Fuel Futures", "Singapore"),
               ("Sing Kero", "Jet Fuel Futures", "Singapore"),
               ("Jet USGC", "Jet Fuel Futures", "US"),
               
               ("USGC HSFO (Platts)", "Fuel Oil Futures", 'US'),
               ("3.5% FOB Med Cg", "Fuel Oil Futures", 'Singapore'),
               
                ("0.1% FOB ARA Bg", "Gasoil Futures", 'Europe'),
                ('HO 1st Line', 'Heating Oil Futures', 'US'),
                ("WTI", "WTI Crude Futures", "US"),
               ("WTI 1st Line/Brent 1st Line", "Crude Diff Futures", "US"),
               ("Argus WTI Mid 1st Line", "Crude Futures", "US") 
               ]


#commodities = [("Brent Sing Marker 1st Line", "Crude Futures", "Europe")]

ST = Forward_Curve(commodities)

final_dataframes = ST.get_dataframe()

# Create a folder to store CSV files if it doesn't exist
folder_path = '.venv/Petroleum/Collect Forward Curves/temporary inventory'
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Iterate through each commodity and its corresponding DataFrames
for commodity, dfs in final_dataframes.items():
    commodity = commodity.replace('/', '-')
    # Create a subfolder for the commodity if it doesn't exist
    commodity_folder_path = os.path.join(folder_path, commodity)
    if not os.path.exists(commodity_folder_path):
        os.makedirs(commodity_folder_path)
    
    # Iterate through each DataFrame in the commodity's DataFrames
    for df_name, df in dfs.items():
        
        # Define the file path for the CSV file
        file_name = f"{df_name}.csv"
        file_path = os.path.join(commodity_folder_path, file_name)
        
        # Save the DataFrame as a CSV file
        df.to_csv(file_path, index=False)
        
        print(f"CSV file saved for {commodity} - {df_name}: {file_path}")

