# -*- coding: utf-8 -*-
"""
Created on Sat Apr  5 22:04:10 2025

@author: Administrator
"""



#%% [ sheet_name = 114年 ]


import pandas as pd
import numpy as np
import os




class DataLoader:
    def __init__(self, Config, store_name_list, all_data_dict, center_kitchen_df):
        self.year = Config['year']
        self.month = Config['month']
        self.company_name = Config['company_name']
        self.each_area_path = Config['each_area_path']
        self.path_one = Config['path_one']
        self.path_two = Config['path_two']
        self.last_year_path = Config['last_year_path']
        self.this_year_path = Config['this_year_path']
        self.path_four = Config['path_four']
        self.path_five = Config['path_five']
        self.store_name_list = store_name_list
        self.all_data_dict = all_data_dict
        self.center_kitchen_df = center_kitchen_df
        self.columns_list = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二', '合  計']

    def load_all(self):
        self.get_file_dict_one()
        self.get_file_dict_two()
        self.get_file_data_three()
        self.get_file_data_four()
        self.get_file_dict_five()
        
        return self.all_data_dict
    
    def _generate_file_names(self, prefix, year, month):
        names = []
        for i in range(1, month+1):
            month_str = f'0{i}' if i < 10 else str(i)
            if self.company_name == '斑鳩的窩':
                names.append(f'{prefix}{year}{month_str}.xlsx')
            elif self.company_name == '聚椒':
                names.append(f'聚椒-{prefix}{year}{month_str}.xlsx')
        return names
    
    #### 1. 折扣金額、來客數 (11401.xlsx)
    
    # 11401.xlsx ....
    # ui介面選取方式 選擇資料夾內含 11401.xlsx .....
    def get_file_dict_one(self):
        '''
        讀取excel資料
        dict
        { 月份 : dataframe }
    
        '''
        file_name_one = self._generate_file_names(prefix="", year=self.year, month=self.month)
        
        self.file_dict_one = {}
        for i in range(self.month):
            try:
                df = pd.read_excel(os.path.join(self.path_one, file_name_one[i]), sheet_name='總表')                     
            except FileNotFoundError:
                print(f'檔案不存在: {os.path.join(self.path_one, file_name_one[i])}')
            except ValueError as e:
                print(f'資料格式錯誤在檔案 {file_name_one[i]}: {str(e)}')
            except Exception as e:
                print(f'發生未知錯誤在檔案 {file_name_one[i]}: {str(e)}')
                
                
            df = df.iloc[:, :len(self.store_name_list)+2]
            #第一欄設為index
            df.set_index(df.columns[0], inplace=True)
            #第一列設為column
            df.columns = df.iloc[0]  
            df = df.iloc[1:, :]
            #存入字典
            self.file_dict_one[i+1] = df
            
            discount = df.loc['折扣金額']
            customers = df.loc['來客數']

            for name in self.store_name_list:
                if name in df.columns:
                    self.all_data_dict[name][i+1]['折扣金額'] = discount.loc[name]
                    self.all_data_dict[name][i+1]['來客數'] = customers.loc[name]
                    

    #### 2. 應發薪資、獎金、工時合計 (薪資11401.xlsx) 
    # 薪資11401.xlsx ....
    # ui介面選取方式 選擇資料夾內含 薪資11401.xlsx .....
    def get_file_dict_two(self):
        '''
        讀取excel資料
        dict
        { 月份 : dataframe }
    
        '''
        self.file_name_two = self._generate_file_names(prefix="薪資", year=self.year, month=self.month)
        
        for i in range(self.month):
            try:
                excel_dict = pd.read_excel(os.path.join(self.path_two, self.file_name_two[i]), sheet_name=None)
            except FileNotFoundError:
                print(f'檔案不存在: {os.path.join(self.path_two, self.file_name_two[i])}')
            except ValueError as e:
                print(f'資料格式錯誤在檔案 {self.file_name_two[i]}: {str(e)}')
            except Exception as e:
                print(f'發生未知錯誤在檔案 {self.file_name_two[i]}: {str(e)}')
                
            for name in self.store_name_list:
                sheet_name = f'薪資表-{name[:-1]}'
                if sheet_name in excel_dict.keys():
                    df = excel_dict[sheet_name]
                    self.all_data_dict[name][i+1]['應發薪資'] = self._get_salary(df)
                    self.all_data_dict[name][i+1]['奬金'] = self._get_bonus(df)
                    self.all_data_dict[name][i+1]['工時合計'] = self._get_work_hours(df)
                    
 
    #應發薪資
    def _get_salary(self, df):
        row_idx, col_idx = df.eq('應發薪資').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]
        val = df.iloc[positions[0], positions[1]+2]
        return val
    
    #獎金
    def _get_bonus(self, df):
        row_idx, col_idx = df.eq('奬金').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]
        val = df.iloc[positions[0], positions[1]+1]
        return val
    
    #工時合計
    def _get_work_hours(self, df):
        row_idx, col_idx = df.eq('合   計').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]
        val = df.iloc[positions[0], positions[1]+4]
        return val    
    
                              
    
    #### 3. 營業額、去年營業額、費用、毛利、單位淨利、租金占比 (114年度損益表.xlsx)
    # 113年度損益表.xlsx、114年度損益表.xlsx
    # ui介面選取方式 直接選擇excel檔案
    def get_file_data_three(self):
        
        index_list = ['營業額', '去年營業額', '費用', '毛利', '單位淨利']
        
        try:
            self.last_year_dict = pd.read_excel(self.last_year_path, sheet_name=None)
            
        except:
            last_year = self.year -1
            print(f'未讀取到{last_year}年度損益表.xlsx')
        try:    
            self.this_year_dict = pd.read_excel(self.this_year_path, sheet_name=None)
        except:
            print(f'未讀取到{self.year}年度損益表.xlsx')
        
        for name in self.store_name_list:
            if name in self.this_year_dict.keys():
                last_year_df = self.last_year_dict[name]                
                last_year_df.set_index(last_year_df.columns[0], inplace=True)
                this_year_df = self.this_year_dict[name]
                this_year_df.set_index(this_year_df.columns[0], inplace=True)
                for i in range(self.month):
                    self._update_data(this_year_df, name, index_list[0], i)
                    self._update_data(this_year_df, name, index_list[2], i)
                    self._update_data(this_year_df, name, index_list[3], i)
                    self._update_data(this_year_df, name, index_list[4], i)
                    if '營業額' in last_year_df.index:
                        last_year_sales = list(last_year_df.loc['營業額'])[2+i]  #從第3個值開始取
                        self.all_data_dict[name][i+1][index_list[1]] = last_year_sales
        
        #取得租金欄位
        self.rent_df = self._get_rent()
    
    
    def _update_data(self, df, name, index, num):
        if index in df.index:
            self.all_data_dict[name][num+1][index] = list(df.loc[index])[2 + num]
            
            
    def _get_rent(self):
        #總表 df
        summary_table_df = self.this_year_dict['總表']
        summary_table_df.columns = summary_table_df.iloc[0, :]
        summary_table_df.set_index(summary_table_df.columns[0], inplace=True)
        
        #取得index為租金的列數
        try:
            row_number = summary_table_df.index.get_loc('租金(全)')
        except:
            print('❌ 今年損益表的租金欄位有誤...')
        
        #租金 df
        all_rent_df = summary_table_df.iloc[row_number:row_number+10, :13]
        
        #只有excel要呈現的租金row
        index_list = ['租金佔比(全)', '租金佔比(路面店)', '租金抽成(美食街+店中店)', '租金抽成(美食街)', '租金抽成(路面店+店中店)']
        self.rent_df = pd.DataFrame(index=index_list, columns=all_rent_df.columns)
        for i in range(1, 6):
            rent = all_rent_df.iloc[i*2-2, :self.month].fillna(0)  
            ratio = all_rent_df.iloc[i*2-1, :self.month].fillna(0)  
            sales = (rent/ratio).fillna(0) 
            sum_ratio = (sum(rent)/sum(sales))
                 
            row_df = all_rent_df.iloc[i*2-1, :self.month]
            self.rent_df.iloc[i-1, :self.month] = row_df
            self.rent_df.iloc[i-1, -1] = sum_ratio
        
        self.rent_df = self.rent_df.applymap(lambda x: f"%{x}" if pd.notna(x) else x)
        self.rent_df.columns = self.columns_list
        return self.rent_df
        
        
        
    #### 4. 營業目標、淨利目標 (114年度預算(全區).xlsx)
    # 114年度預算(全區).xlsx
    # ui介面選取方式 直接選擇excel檔案  
    def get_file_data_four(self):
        try:
            df = pd.read_excel(self.path_four, sheet_name=f'{self.year}年度')
        except FileNotFoundError:
            print(f'檔案不存在: {self.path_four}')
        except ValueError as e:
            print(f'資料格式錯誤或找不到工作表 {self.year}年度，在檔案 {self.path_four}: {str(e)}')
        except Exception as e:
            print(f'發生未知錯誤在檔案 {self.path_four}: {str(e)}')
            
        #第一列設為col
        df.columns = df.iloc[0]
        df = df.drop(0).reset_index(drop=True)

        
        index_list = ['營業目標', '淨利目標']
        
        offset_num = 6
        for name in self.store_name_list:
            if name in df.columns:
                target_list = df[name].tolist()
                for i in range(self.month):
                    #營業目標
                    self.all_data_dict[name][i+1][index_list[0]] = target_list[1+offset_num*i]                    
                    #淨利目標 
                    self.all_data_dict[name][i+1][index_list[1]] = target_list[4+offset_num*i]
        return self.all_data_dict         
    
    
    
    ####5. 實際毛利率 (月報表11401.xlsx)
    # 月報表11401.xlsx ....
    # ui介面選取方式 選擇資料夾內含 月報表11401.xlsx .....   
    
    def get_file_dict_five(self):
        '''
        讀取excel資料
        dict
        { 月份 : dataframe }
    
        '''
        file_name_five = self._generate_file_names(prefix="月報表", year=self.year, month=self.month)
        
        for i in range(self.month):
            try:
                excel_dict = pd.read_excel(os.path.join(self.path_five, file_name_five[i]), sheet_name=None)
            except FileNotFoundError:
                print(f'檔案不存在: {os.path.join(self.path_five, file_name_five[i])}')
            except ValueError as e:
                print(f'資料格式錯誤在檔案 {file_name_five[i]}: {str(e)}')
            except Exception as e:
                print(f'發生未知錯誤在檔案 {file_name_five[i]}: {str(e)}')
            
            for name in self.store_name_list:
                sheet_name = f'成本費用-{name[:-1]}'
                if sheet_name in excel_dict.keys():
                    df = excel_dict[sheet_name]
                    #取得實際毛利率的位置
                    row_idx, col_idx = df.eq('實際\n毛利率').to_numpy().nonzero()
                    positions = list(zip(row_idx, col_idx))[0]
                    #取得數值
                    value = df.iloc[positions[0], positions[1]+2]
                    self.all_data_dict[name][i+1]['實際毛利率'] = value

        # return self.all_data_dict
    
    #### 中廚
    def load_center_kitchen_all(self):
        self._get_file_data_in_PLtable()
        self._get_file_data_in_salary_table()
        return self.center_kitchen_df
    
    #損益表    
    def _get_file_data_in_PLtable(self):
    
        index_list = ['銷貨收入', '食材銷貨收入', '雜項銷貨收入', '公務費收入', '其他收入', 
                      '其他支出', '費用', '實際毛利', '實際毛利率(%)', "單位淨利"]
        
        df = pd.read_excel(self.this_year_path, sheet_name='中廚', header=1)
        df.set_index(df.columns[0], inplace=True)
        df_1 = df.iloc[:, 1:14]
        #中間用不到的月份的值設為nan
        df_1.iloc[:, self.month:-1] = np.nan
            
        df_1.columns = self.columns_list
        
        for index in index_list:
            if index in df_1.index:
                self.center_kitchen_df.loc[index] = df_1.loc[index] 
    
        #取得年終獎金的值
        row_idx, col_idx = df.eq('年終奬金').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]   
        Year_end_bonus = pd.DataFrame(df.iloc[positions[0], :][1:14]).T
        Year_end_bonus.columns = self.columns_list
        #中間用不到的月份的值設為nan
        Year_end_bonus.iloc[:, self.month:-1] = np.nan
        self.center_kitchen_df.loc['年終奬金'] = Year_end_bonus.values[0]
        
        # return self.center_kitchen_df
    
    
    #薪資11401.xlsx ....
    def _get_file_data_in_salary_table(self):
        sheet_name = '薪資表-中廚'
        self.file_name_two = self._generate_file_names(prefix="薪資", year=self.year, month=self.month)

        for i in range(self.month):
            try:
                excel_dict = pd.read_excel(os.path.join(self.path_two, self.file_name_two[i]), sheet_name=None)
            except FileNotFoundError:
                print(f'檔案不存在: {os.path.join(self.path_two, self.file_name_two[i])}')
            except ValueError as e:
                print(f'資料格式錯誤在檔案 {self.file_name_two[i]}: {str(e)}')
            except Exception as e:
                print(f'發生未知錯誤在檔案 {self.file_name_two[i]}: {str(e)}')
                        
            df = excel_dict[sheet_name]            
            #應發薪資
            salary_payable = self._get_salary(df)
            self.center_kitchen_df.iloc[self.center_kitchen_df.index.get_loc('應發薪資'), i] = salary_payable
            
            #獎金
            bonus = self._get_bonus(df)
            self.center_kitchen_df.iloc[self.center_kitchen_df.index.get_loc('奬金'), i] = bonus
    
            #現場人員薪資
            staff_salary = self._get_staff_salary(df)
            self.center_kitchen_df.iloc[self.center_kitchen_df.index.get_loc('現場人員薪資'), i] = staff_salary
            
            #工讀生薪資比(佔食材進貨金額)
            pt_salary_ratio = self._get_pt_salary_ratio(df)
            self.center_kitchen_df.iloc[self.center_kitchen_df.index.get_loc('工讀生薪資比\n(佔食材進貨金額)'), i] = pt_salary_ratio
            
            #工時生產力
            hour_productivity = self._get_hour_productivity(df)
            self.center_kitchen_df.iloc[self.center_kitchen_df.index.get_loc('工時生產力'), i] = hour_productivity
            
            #薪資生產力
            salary_productivity = self._get_salary_productivity(df)
            self.center_kitchen_df.iloc[self.center_kitchen_df.index.get_loc('薪資生產力'), i] = salary_productivity
    
        # 加總每一列
        # skipna=True就是遇到NaN自動跳過，直接當0
        self.center_kitchen_df[self.columns_list[-1]] = self.center_kitchen_df[self.columns_list[:-1]].sum(axis=1, skipna=True)
        return self.center_kitchen_df
        
    #現場人員薪資(薪資11401.xlsx)
    def _get_staff_salary(self,df):
        row_idx, col_idx = df.eq('現場人員薪資').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]   
        result = df.iloc[positions[0], positions[1]+2]
        return  result   
    
    #工讀生薪資比(佔食材進貨金額)(薪資11401.xlsx)
    def _get_pt_salary_ratio(self,df):
        row_idx, col_idx = df.eq('現場人員薪資佔比\n(薪資佔營業店食材進貨金額)').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]   
        result = df.iloc[positions[0], positions[1]+3]
        return  result
    
    #工時生產力(薪資11401.xlsx)
    def _get_hour_productivity(self,df):
        row_idx, col_idx = df.eq('工時生產力').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]   
        result = df.iloc[positions[0], positions[1]+3]
        return  result
    
    #薪資生產力(薪資11401.xlsx)
    def _get_salary_productivity(self,df):
        row_idx, col_idx = df.eq('薪資生產力').to_numpy().nonzero()
        positions = list(zip(row_idx, col_idx))[0]   
        result = df.iloc[positions[0], positions[1]+3]
        return  result    
    
    
    
    #### 取得同期單位淨利
    def load_same_period_profit_dict(self):
        self._cal_same_period_profit()
        return self.same_period_profit_dict
    
    def _cal_same_period_profit(self):
        this_year_every_month_store = pd.read_excel(self.each_area_path, sheet_name = f'{self.year}年各月店家')
        last_year_every_month_store = pd.read_excel(self.each_area_path, sheet_name = f'{self.year-1}年各月店家')
        
        # 先準備空字典存結果
        intersection_result = {}
        
        for month in this_year_every_month_store.columns:
            # 取出當月的店家名單（去掉nan，再轉成set）
            this_year_set = set(this_year_every_month_store[month].dropna())
            last_year_set = set(last_year_every_month_store[month].dropna())            
            # 做交集
            intersection = this_year_set & last_year_set
            intersection_result[month] = intersection
        
        self.same_period_profit_dict = {month+1: {} for month in range(self.month)}
        for m in range(self.month):
            month_str = f'{m+1}月'
            store_name = intersection_result[month_str]
            cnt_last_profit = 0
            cnt_this_profit = 0
            for name in store_name:
                last_year_df = self.last_year_dict[name]
                cnt_last_profit = DataLoader._get_profit(last_year_df, m, cnt_last_profit)
                this_year_df = self.this_year_dict[name]
                cnt_this_profit = DataLoader._get_profit(this_year_df, m, cnt_this_profit)
            growth_rate = (cnt_last_profit - cnt_this_profit) / cnt_last_profit
            self.same_period_profit_dict[m+1]['去年同期營業額'] = cnt_last_profit
            self.same_period_profit_dict[m+1]['今年同期營業額'] = cnt_this_profit
            self.same_period_profit_dict[m+1]['去年同期成長率'] = growth_rate
            

        
    @staticmethod        
    def _get_profit(df, m, cnt):
        # new_df = df.set_index(df.columns[0])
        val = list(df.loc['單位淨利'])[2 + m]
        val = 0 if pd.isna(val) else val
        cnt = cnt + val
        return cnt
    
