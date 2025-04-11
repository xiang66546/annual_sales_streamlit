# -*- coding: utf-8 -*-
"""
Created on Sat Apr  5 21:20:12 2025

@author: Administrator
"""
import pandas as pd
import numpy as np


class Calculator:
    def _cal_division(df, up_index, down_index, result_index):
        '''
        up_index : 分子index
        down_index : 分母index
        result_index : 目標index

        '''
        up_df = df.loc[up_index]
        down_df = df.loc[down_index]
        down_df = down_df.replace(0, np.nan)
        result_df = up_df / down_df
        result_df = pd.DataFrame(result_df).T 
        result_df.index = [result_index]
        df.update(result_df)
        
    ### 銷貨收入 = 折扣金額 + 營業額 ###
    def _cal_sales_revenue(df):
        discount_amount = df.loc['折扣金額']
        sales = df.loc['營業額']
    
        # 正常加總，遇到單邊NaN時當0，但雙NaN時保持NaN
        sales_revenue = discount_amount.fillna(0) + sales.fillna(0)
    
        # 如果兩個都是NaN，要保持是NaN
        mask_all_nan = discount_amount.isna() & sales.isna()
        sales_revenue[mask_all_nan] = pd.NA  # 或 np.nan
    
        # 整理成DataFrame
        sales_revenue_df = pd.DataFrame(sales_revenue).T
        sales_revenue_df.index = ['銷貨收入']
        df.update(sales_revenue_df)
        
    
    ### 成長率 = (營業額 - 去年營業額) / 去年營業額 ###
    def _cal_growth_rate(df):
        sales_df = df.loc['營業額']
        last_year_sales_df = df.loc['去年營業額']
        last_year_sales_df = last_year_sales_df.replace(0, np.nan)
        growth_rate_df = (sales_df - last_year_sales_df) / last_year_sales_df
        growth_rate_df = pd.DataFrame(growth_rate_df).T 
        growth_rate_df.index = ['成長率']
        df.update(growth_rate_df)
    
    ### 營業額占比 = 營業額 / 總營業額 ###
    def _cal_sales_proportion(total_sales, df, result_index):
        sales_df = df.loc['營業額']
        total_sales = total_sales.replace(0, np.nan)
        sales_proportion_df = sales_df / total_sales
        sales_proportion_df = pd.DataFrame(sales_proportion_df).T 
        sales_proportion_df.index = [result_index]
        df.update(sales_proportion_df)
        
    #%%中廚
    ### 收入合計 = 銷貨收入 + 公務費收入 + 其他收入 - 其他支出 ###
    def _cal_total_income(df):
        sales_revenue = df.loc['銷貨收入']
        official_expenses_income = df.loc['公務費收入']
        other_income = df.loc['其他收入']
        other_expenses = df.loc['其他支出']
    
        # 把要加的部分先加起來（NaN當作0處理），然後單獨處理全NaN情況
        add_part = sales_revenue.fillna(0) + official_expenses_income.fillna(0) + other_income.fillna(0)
        sub_part = other_expenses.fillna(0)
    
        total_income = add_part - sub_part
    
        # 處理如果四個欄位原本都是NaN的情況，total_income該是NaN
        mask_all_nan = sales_revenue.isna() & official_expenses_income.isna() & other_income.isna() & other_expenses.isna()
        total_income[mask_all_nan] = pd.NA  # 或用 numpy 的 np.nan
    
        # 整理成新的DataFrame
        total_income_df = pd.DataFrame(total_income).T
        total_income_df.index = ['收入合計']
    
        # 更新回原df
        df.update(total_income_df)
        
        
        