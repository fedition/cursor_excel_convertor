import pandas as pd
from typing import Dict, List, Optional

class DataProcessor:
    def __init__(self):
        self.products_df: Optional[pd.DataFrame] = None
        self.suppliers_df: Optional[pd.DataFrame] = None
        self.inventory_df: Optional[pd.DataFrame] = None
        self.members_df: Optional[pd.DataFrame] = None

    def process_products(self, df: pd.DataFrame) -> str:
        """处理商品数据"""
        try:
            # 检查必需列
            required_columns = ['商品编号', '商品名称', '商品类别', '单价']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                return f"商品表缺少必需列：{', '.join(missing_columns)}"

            # 数据清理
            df = df.copy()
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].str.strip()
            
            # 类型转换
            df['单价'] = pd.to_numeric(df['单价'], errors='coerce')
            
            # 删除重复行
            df = df.drop_duplicates()
            
            # 处理缺失值
            df = df.fillna({
                '商品类别': '未分类',
                '单价': 0.0
            })
            
            self.products_df = df
            return f"成功处理商品数据，共 {len(df)} 条记录"
        except Exception as e:
            return f"处理商品数据时出错：{str(e)}"

    def process_suppliers(self, df: pd.DataFrame) -> str:
        """处理供应商数据"""
        try:
            # 检查必需列
            required_columns = ['供应商编号', '供应商名称', '联系人', '联系电话']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                return f"供应商表缺少必需列：{', '.join(missing_columns)}"

            # 数据清理
            df = df.copy()
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].str.strip()
            
            # 删除重复行
            df = df.drop_duplicates()
            
            # 处理缺失值
            df = df.fillna({
                '联系人': '未知',
                '联系电话': '未知'
            })
            
            self.suppliers_df = df
            return f"成功处理供应商数据，共 {len(df)} 条记录"
        except Exception as e:
            return f"处理供应商数据时出错：{str(e)}"

    def process_inventory(self, df: pd.DataFrame) -> str:
        """处理库存数据"""
        try:
            # 检查必需列
            required_columns = ['商品编号', '库存数量', '库存位置']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                return f"库存表缺少必需列：{', '.join(missing_columns)}"

            # 数据清理
            df = df.copy()
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].str.strip()
            
            # 类型转换
            df['库存数量'] = pd.to_numeric(df['库存数量'], errors='coerce')
            
            # 删除重复行
            df = df.drop_duplicates()
            
            # 处理缺失值
            df = df.fillna({
                '库存数量': 0,
                '库存位置': '未知'
            })
            
            self.inventory_df = df
            return f"成功处理库存数据，共 {len(df)} 条记录"
        except Exception as e:
            return f"处理库存数据时出错：{str(e)}"

    def process_members(self, df: pd.DataFrame) -> str:
        """处理会员数据"""
        try:
            # 检查必需列
            required_columns = ['会员编号', '会员姓名', '会员等级', '积分']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                return f"会员表缺少必需列：{', '.join(missing_columns)}"

            # 数据清理
            df = df.copy()
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].str.strip()
            
            # 类型转换
            df['积分'] = pd.to_numeric(df['积分'], errors='coerce')
            
            # 删除重复行
            df = df.drop_duplicates()
            
            # 处理缺失值
            df = df.fillna({
                '会员等级': '普通会员',
                '积分': 0
            })
            
            self.members_df = df
            return f"成功处理会员数据，共 {len(df)} 条记录"
        except Exception as e:
            return f"处理会员数据时出错：{str(e)}"

    def process_all_data(self, excel_file: pd.ExcelFile) -> Dict[str, str]:
        """处理所有数据"""
        results = {}
        
        # 按顺序处理各个表
        if '商品' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name='商品')
            results['商品'] = self.process_products(df)
            
        if '供应商' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name='供应商')
            results['供应商'] = self.process_suppliers(df)
            
        if '库存' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name='库存')
            results['库存'] = self.process_inventory(df)
            
        if '会员' in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name='会员')
            results['会员'] = self.process_members(df)
            
        return results 