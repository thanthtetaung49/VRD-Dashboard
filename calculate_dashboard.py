import pandas as pd
import datetime
import glob
from config import INPUT_BASE_DIR, OUTPUT_BASE_DIR

class GenerateReports: 
    def __init__(self, outboundFile):
        self.df = pd.read_csv(outboundFile)
        self.productOffer = pd.read_csv(f"{INPUT_BASE_DIR}\\Product_Offer\\ProductOffer.csv")
        self.datetime_mis = datetime.datetime(2025, 5, 10).strftime("%d%m%Y")
        
        file_pattern = f"{INPUT_BASE_DIR}\\MIS_Pack_Sale\\{self.datetime_mis}\\Daily_Pack_Sales_Report_*.csv"
        matched_files = glob.glob(file_pattern)
        
        if matched_files:
            if len(matched_files) > 1:
                print(f"Warning: Multiple files matched. Using the first one: {matched_files[0]}")
            self.mispackSale = pd.read_csv(matched_files[0], delimiter='|')
        else:
            raise FileNotFoundError(f"No file matched pattern: {file_pattern}")

    
    def total_attempts_call(self):
        df = self.df
        
        df = df.copy()
        df['OutboundNumber'] = pd.to_numeric(df['OutboundNumber'], errors='coerce')
        df = df.dropna(subset=['OutboundNumber'])
        df['OutboundNumber'] = df['OutboundNumber'].astype(int)
        df['OutboundNumberLength'] = df['OutboundNumber'].apply(lambda x: len(str(x)))
        df['checkPhoneNumber'] = df['OutboundNumber'].astype(str).str[:2]
        df['OutboundNumberLength'] = df['OutboundNumberLength'].astype(int)
        df['checkPhoneNumber'] = df['checkPhoneNumber'].astype(str)
        
        filtered_df = df[(df['OutboundNumberLength'] >= 10) & (df['checkPhoneNumber'] == '97')]
        
        total_attempts_call = filtered_df['OutboundNumber'].count()
        
        return total_attempts_call
    
    def total_login_agents(self):
        df = self.df
        
        df = df.copy()
        df['OutboundNumber'] = pd.to_numeric(df['OutboundNumber'], errors='coerce')
        df = df.dropna(subset=['OutboundNumber'])
        df['OutboundNumber'] = df['OutboundNumber'].astype(int)
        df['OutboundNumberLength'] = df['OutboundNumber'].apply(lambda x: len(str(x)))
        df['checkPhoneNumber'] = df['OutboundNumber'].astype(str).str[:2]
        df['OutboundNumberLength'] = df['OutboundNumberLength'].astype(int)
        df['checkPhoneNumber'] = df['checkPhoneNumber'].astype(str)
        
        filtered_df = df[(df['OutboundNumberLength'] >= 10) & (df['checkPhoneNumber'] == '97')]
        remove_duplicate_df = filtered_df.drop_duplicates(subset=['AgentCID'])
        total_login_agents = remove_duplicate_df['AgentCID'].count()
        
        
        return total_login_agents
    
    def total_success_calls(self):
        df = self.df
        
        df = df.copy()
        df['OutboundNumber'] = pd.to_numeric(df['OutboundNumber'], errors='coerce')
        df = df.dropna(subset=['OutboundNumber'])
        df['OutboundNumber'] = df['OutboundNumber'].astype(int)
        df['OutboundNumberLength'] = df['OutboundNumber'].apply(lambda x: len(str(x)))
        df['checkPhoneNumber'] = df['OutboundNumber'].astype(str).str[:2]
        df['OutboundNumberLength'] = df['OutboundNumberLength'].astype(int)
        df['checkPhoneNumber'] = df['checkPhoneNumber'].astype(str)
        
        filtered_df = df[(df['OutboundNumberLength'] >= 10) & (df['checkPhoneNumber'] == '97') & (df['TalkTime'] > 5)]
        
        total_success_calls = filtered_df['TalkTime'].count()
        
        return total_success_calls
    
    def product_counts(self):
        product_master_df = self.productOffer
        outbound_df = self.df
        mispackSale = self.mispackSale
        
        mispackSale['checkPhoneNumber'] = mispackSale['MSISDN'].astype(str).str[:2]
        mispackSale['msisdnsLength'] = mispackSale['MSISDN'].fillna(0).astype(int).apply(lambda x: len(str(x)))
        
        # print(mispackSale['checkPhoneNumber'], mispackSale['msisdnsLength'])
        
        mispackSale = mispackSale[
            (mispackSale['checkPhoneNumber'] == '97') & (mispackSale['msisdnsLength'] >= 10)
        ]
        
        product_matched_sales=mispackSale[
            mispackSale['OFFERID'].isin(product_master_df['Product ID'])
        ]
        
        outbound_df['OutboundNumber'] = pd.to_numeric(outbound_df['OutboundNumber'], errors='coerce').fillna(0).astype(int)
        
        # print(outbound_df['OutboundNumber'], product_matched_sales['MSISDN'])
        
        df_filtered_with_outbound=product_matched_sales[
            product_matched_sales['MSISDN'].isin(outbound_df['OutboundNumber'])
        ]
        
        group_df_based_offerid = df_filtered_with_outbound.groupby('OFFERID', as_index=False)['COUNT OF PACK SALES'].sum()
        
        # print(group_df_based_offerid)
        return group_df_based_offerid

class GenerateReportsVmd2(GenerateReports):
    def __init__(self, outputFile):
       super().__init__(outputFile)