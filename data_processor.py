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
            # 定义表头映射关系
            header_mapping = {
                '原系统商品编码': ['商品编码', '药品编码', '药品编号', '商品编号','商品代码'],
                '商品名称': ['商品名称', '药品名称', '名称', '品名'],
                '通用名': ['通用名', '药品通用名', '通用名称'],
                '商品规格': ['商品规格', '规格', '规格型号'],
                '包装规格': ['包装规格', '包装', '包装型号'],
                '单位': ['单位', '计量单位', '销售单位'],
                '剂型': ['剂型', '药品剂型', '制剂类型'],
                '生产厂家': ['生产厂家', '生产厂商', '生产商', '生产企业'],
                '商品产地': ['商品产地', '产地', '原产地'],
                '条码': ['条码', '条形码', '商品条码','助记符'],
                '药品本位码': ['药品本位码', '本位码'], 
                '批准文号': ['批准文号', '药品批准文号', '注册证号'],
                '零售价': ['零售价', '价格', '售价', '销售价','单价'],
                '会员价': ['会员价', '会员价格', '会员销售价'],
                '保质期': ['保质期'],
                '存储条件': ['存储条件', '储存条件', '贮藏条件'],
                '是否处方药': ['是否处方药', '处方药', '处方药类型','处方类型'],
                '是否医保药品': ['是否医保药品', '医保药品', '医保类型'],
                '是否含麻黄碱': ['是否含麻黄碱', '含麻', '含麻黄碱', '麻黄碱类型'],
                '是否中药材': ['是否中药材', '中药材', '中药材类型']
            }

            # 查找并重命名列
            new_columns = {}
            found_columns = set()  # 用于记录找到的列
            for new_col, possible_names in header_mapping.items():
                for old_col in df.columns:
                    if old_col in possible_names:
                        new_columns[old_col] = new_col
                        found_columns.add(new_col)
                        break

            # 重命名列
            df = df.rename(columns=new_columns)

            # 创建缺失的列并设置默认值
            for col in header_mapping.keys():
                if col not in df.columns:
                    df[col] = None

            # 数据清理
            df = df.copy()
            
            # 对字符串类型的列进行清理
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 将非字符串值转换为字符串
                    df[col] = df[col].astype(str)
                    # 清理字符串
                    df[col] = df[col].str.strip()
                    # 将空字符串和'nan'替换为空字符串
                    df[col] = df[col].replace(['nan', 'None', 'NULL', ''], '')
            
            # 类型转换（如果存在这些列）
            if '零售价' in df.columns:
                df['零售价'] = pd.to_numeric(df['零售价'], errors='coerce')
            if '会员价' in df.columns:
                df['会员价'] = pd.to_numeric(df['会员价'], errors='coerce')
            
            # 处理原系统商品编码
            if '原系统商品编码' in df.columns:
                # 转换为字符串类型
                df['原系统商品编码'] = df['原系统商品编码'].astype(str)
                # 清理空白字符
                df['原系统商品编码'] = df['原系统商品编码'].str.strip()
                # 截断超过20位的编码
                df['原系统商品编码'] = df['原系统商品编码'].str[:20]

            # 处理商品名称
            if '商品名称' in df.columns:
                # 转换为字符串类型
                df['商品名称'] = df['商品名称'].astype(str)
                # 清理空白字符
                df['商品名称'] = df['商品名称'].str.strip()
                # 截断超过100位的商品名称
                df['商品名称'] = df['商品名称'].str[:100]
            # 处理条码
            if '条码' in df.columns:
                # 转换为字符串类型
                df['条码'] = df['条码'].astype(str)
                # 清理空白字符
                df['条码'] = df['条码'].str.strip()
                # 只保留数字
                df['条码'] = df['条码'].str.extract('(\d+)', expand=False)
                # 截断超过50位的条码
                df['条码'] = df['条码'].str[:50]

            # 处理"是否处方药"字段
            if '是否处方药' in df.columns:
                # 创建映射字典
                otc_mapping = {
                    '非处方药': '非处方药',
                    'OTC': '非处方药',
                    '处方药': '处方药',
                    'RX': '处方药',
                    '甲类非处方药': '甲类非处方药', 
                    '甲类OTC': '甲类非处方药',
                    '乙类非处方药': '乙类非处方药',
                    '乙类OTC': '乙类非处方药'
                }
                
                # 转换为字符串并清理
                df['是否处方药'] = df['是否处方药'].astype(str).str.strip().str.upper()
                
                # 应用映射,未匹配项设为"其他"
                df['是否处方药'] = df['是否处方药'].map(otc_mapping).fillna('其他')

            # 处理是否中药材字段
            if '是否中药材' in df.columns:
                # 转换为字符串并清理
                df['是否中药材'] = df['是否中药材'].astype(str).str.strip().str.upper()
                # 将各种"是"的表达映射为"是"
                df['是否中药材'] = df['是否中药材'].map(lambda x: '是' if x in ['是','YES','Y','TRUE','1'] else '否')

            # 处理是否含麻黄碱字段
            if '是否含麻黄碱' in df.columns:
                # 转换为字符串并清理
                df['是否含麻黄碱'] = df['是否含麻黄碱'].astype(str).str.strip().str.upper()
                # 将各种"是"的表达映射为"是"
                df['是否含麻黄碱'] = df['是否含麻黄碱'].map(lambda x: '是' if x in ['是','YES','Y','TRUE','1'] else '否')

            # 处理是否医保药品字段
            if '是否医保药品' in df.columns:
                # 转换为字符串并清理
                df['是否医保药品'] = df['是否医保药品'].astype(str).str.strip().str.upper()
                # 将各种"是"的表达映射为"是"
                df['是否医保药品'] = df['是否医保药品'].map(lambda x: '是' if x in ['是','YES','Y','TRUE','1'] else '否')

            
            # 删除重复行
            df = df.drop_duplicates()
            
            # 处理缺失值
            fillna_dict = {
                '零售价': 0.01,
                '会员价': '',
                '生产厂家': '',
                '商品产地': '',
                '包装规格': '',
                '单位': '',
                '剂型': '',
                '是否处方药': '否',
                '是否含麻黄碱': '否',
                '是否医保药品': '否',
                '保质期': '',
                '存储条件': '常温',
                '通用名': '',
                '商品规格': '',
                '条码': '',
                '药品本位码': '',
                '批准文号': ''
            }
            
            df = df.fillna(fillna_dict)
            
            
            # 只保留目标列
            target_columns = list(header_mapping.keys())
            df = df[target_columns]
            
            self.products_df = df
            
            # 保存处理后的数据到Excel文件
            try:
                output_path = 'products_output.xlsx'
                df.to_excel(output_path, index=False)
                
                # 生成处理报告
                report = f"成功处理商品数据，共 {len(df)} 条记录\n"
                report += f"数据已保存到 {output_path}\n\n"
                report += "找到的列：\n"
                for col in found_columns:
                    report += f"- {col}\n"
                report += "\n未找到的列（已创建并填充默认值）：\n"
                for col in header_mapping.keys():
                    if col not in found_columns:
                        report += f"- {col}\n"
                
                return report
            except Exception as save_error:
                return f"成功处理商品数据，共 {len(df)} 条记录\n但保存文件时出错：{str(save_error)}"
                
        except Exception as e:
            return f"处理商品数据时出错：{str(e)}"

    def process_suppliers(self, df: pd.DataFrame) -> str:
        """处理供应商数据"""
        try:
            # 定义表头映射关系
            header_mapping = {
                '原系统供应商编码': ['供应商编码', '供应商编号', '系统编码','供应商ID', '编码','供商代码'],
                '单位名称': ['单位名称', '供应商名称', '公司名称', '企业名称','供商名称'],
                '税务登记/信用代码/营业执照号': ['统一社会信用代码', '营业执照号', '税务登记号', '信用代码'],
                '法人代表': ['法人代表', '法定代表人', '负责人','法人'],
                '联系人': ['联系人', '业务联系人', '经办人'],
                '电话': ['电话', '联系电话', '手机号码', '联系电话'],
                '地址': ['地址', '公司地址', '企业地址', '详细地址'],
                '网址': ['网址', '网站', '公司网站'],
                '电子邮箱': ['电子邮箱', '邮箱', 'E-mail', 'email'],
                '销售员': ['销售员', '业务员']
            }

            # 查找并重命名列
            new_columns = {}
            found_columns = set()  # 用于记录找到的列
            for new_col, possible_names in header_mapping.items():
                for old_col in df.columns:
                    if old_col in possible_names:
                        new_columns[old_col] = new_col
                        found_columns.add(new_col)
                        break

            # 重命名列
            df = df.rename(columns=new_columns)

            # 创建缺失的列并设置默认值
            for col in header_mapping.keys():
                if col not in df.columns:
                    df[col] = None

            # 数据清理
            df = df.copy()
            
            # 对字符串类型的列进行清理
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 将非字符串值转换为字符串
                    df[col] = df[col].astype(str)
                    # 清理字符串
                    df[col] = df[col].str.strip()
                    # 将空字符串和'nan'替换为空字符串
                    df[col] = df[col].replace(['nan', 'None', 'NULL', ''], '')
            
            # 处理缺失值
            fillna_dict = {
                '单位名称': '',
                '税务登记/信用代码/营业执照号': '',
                '法人代表': '',
                '联系人': '',
                '电话': '',
                '地址': '',
                '网址': '',
                '电子邮箱': '',
                '销售员': ''
            }
            
            df = df.fillna(fillna_dict)
            # 删除重复行
            df = df.drop_duplicates()
            
            
            # 只保留目标列
            target_columns = list(header_mapping.keys())
            df = df[target_columns]
            
            self.suppliers_df = df
            
            # 保存处理后的数据到Excel文件
            try:
                output_path = 'suppliers_output.xlsx'
                df.to_excel(output_path, index=False)
                
                # 生成处理报告
                report = f"成功处理供应商数据，共 {len(df)} 条记录\n"
                report += f"数据已保存到 {output_path}\n\n"
                report += "找到的列：\n"
                for col in found_columns:
                    report += f"- {col}\n"
                report += "\n未找到的列（已创建并填充默认值）：\n"
                for col in header_mapping.keys():
                    if col not in found_columns:
                        report += f"- {col}\n"
                
                return report
            except Exception as save_error:
                return f"成功处理供应商数据，共 {len(df)} 条记录\n但保存文件时出错：{str(save_error)}"
                
        except Exception as e:
            return f"处理供应商数据时出错：{str(e)}"

    def process_inventory(self, df: pd.DataFrame) -> str:
        """处理库存数据"""
        try:
            # 定义表头映射关系
            header_mapping = {
                'pro系统编码': ['pro系统编码', '系统编码', '编码', '商品编码'],
                '原系统商品编码': ['原系统商品编码', '商品编码', '药品编码', '商品编号'],
                '批号': ['批号', '生产批号', '批次号'],
                '生产日期': ['生产日期', '生产时间', '制造日期'],
                '有效期至': ['有效期至', '有效期', '过期日期', '失效日期'],
                '数量': ['数量', '库存数量', '库存量'],
                '单价': ['单价', '价格', '进价', '采购价'],
                '供应商': ['供应商', '供应商名称', '供货商', '供商名称']
            }

            # 查找并重命名列
            new_columns = {}
            found_columns = set()  # 用于记录找到的列
            for new_col, possible_names in header_mapping.items():
                for old_col in df.columns:
                    if old_col in possible_names:
                        new_columns[old_col] = new_col
                        found_columns.add(new_col)
                        break

            # 重命名列
            df = df.rename(columns=new_columns)

            # 创建缺失的列并设置默认值
            for col in header_mapping.keys():
                if col not in df.columns:
                    df[col] = None

            # 数据清理
            df = df.copy()
            
            # 对字符串类型的列进行清理
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 将非字符串值转换为字符串
                    df[col] = df[col].astype(str)
                    # 清理字符串
                    df[col] = df[col].str.strip()
                    # 将空字符串和'nan'替换为空字符串
                    df[col] = df[col].replace(['nan', 'None', 'NULL', ''], '')
            
            # 类型转换（如果存在这些列）
            if '数量' in df.columns:
                df['数量'] = pd.to_numeric(df['数量'], errors='coerce')
            if '单价' in df.columns:
                df['单价'] = pd.to_numeric(df['单价'], errors='coerce')
            
            # 处理日期字段
            if '生产日期' in df.columns:
                df['生产日期'] = pd.to_datetime(df['生产日期'], errors='coerce')
            if '有效期至' in df.columns:
                df['有效期至'] = pd.to_datetime(df['有效期至'], errors='coerce')
            
            # 删除重复行
            df = df.drop_duplicates()
            
            # 处理缺失值
            fillna_dict = {
                'pro系统编码': '',
                '原系统商品编码': '',
                '批号': '',
                '生产日期': pd.NaT,
                '有效期至': pd.NaT,
                '数量': 0,
                '单价': 0,
                '供应商': ''
            }
            
            df = df.fillna(fillna_dict)
            
            # 只保留目标列
            target_columns = list(header_mapping.keys())
            df = df[target_columns]
            
            self.inventory_df = df
            
            # 保存处理后的数据到Excel文件
            try:
                output_path = 'inventory_output.xlsx'
                df.to_excel(output_path, index=False)
                
                # 生成处理报告
                report = f"成功处理库存数据，共 {len(df)} 条记录\n"
                report += f"数据已保存到 {output_path}\n\n"
                report += "找到的列：\n"
                for col in found_columns:
                    report += f"- {col}\n"
                report += "\n未找到的列（已创建并填充默认值）：\n"
                for col in header_mapping.keys():
                    if col not in found_columns:
                        report += f"- {col}\n"
                
                return report
            except Exception as save_error:
                return f"成功处理库存数据，共 {len(df)} 条记录\n但保存文件时出错：{str(save_error)}"
                
        except Exception as e:
            return f"处理库存数据时出错：{str(e)}"

    def process_members(self, df: pd.DataFrame) -> str:
        """处理会员数据"""
        try:
            # 定义表头映射关系
            header_mapping = {
                '会员姓名': ['会员姓名', '姓名', '客户姓名', '顾客姓名','客户名称'],
                '手机号': ['手机号', '手机号码', '联系电话', '手机'],
                '座机号': ['座机号', '固定电话', '电话', '座机','电话号码'],
                '性别': ['性别', '性别类型'],
                '身份证号': ['身份证号', '身份证号码', '身份证','IDCardCode'],
                '出生年月日': ['出生年月日', '出生日期', '生日', '出生时间'],
                '联系地址': ['联系地址', '地址', '居住地址', '详细地址'],
                '会员卡号': ['会员卡号', '卡号', '会员编号', '会员ID'],
                '剩余积分': ['剩余积分', '积分', '当前积分'],
                '剩余充值金额': ['剩余充值金额', '充值金额', '账户余额', '余额'],
                '剩余赠送金额': ['剩余赠送金额', '赠送金额', '赠送余额']
            }

            # 查找并重命名列
            new_columns = {}
            found_columns = set()  # 用于记录找到的列
            for new_col, possible_names in header_mapping.items():
                for old_col in df.columns:
                    if old_col in possible_names:
                        new_columns[old_col] = new_col
                        found_columns.add(new_col)
                        break

            # 重命名列
            df = df.rename(columns=new_columns)

            # 创建缺失的列并设置默认值
            for col in header_mapping.keys():
                if col not in df.columns:
                    df[col] = None

            # 数据清理
            df = df.copy()
            
            # 对字符串类型的列进行清理
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 将非字符串值转换为字符串
                    df[col] = df[col].astype(str)
                    # 清理字符串
                    df[col] = df[col].str.strip()
                    # 将空字符串和'nan'替换为空字符串
                    df[col] = df[col].replace(['nan', 'None', 'NULL', ''], '')
            
            # 处理手机号和座机号
            if '手机号' in df.columns:
                # 转换为字符串并清理
                df['手机号'] = df['手机号'].astype(str)
                # 只保留数字
                df['手机号'] = df['手机号'].str.extract('(\d+)', expand=False)
                # 截断超过20位的手机号
                df['手机号'] = df['手机号'].str[:20]
                # 将无效值替换为空字符串
                df['手机号'] = df['手机号'].replace(['nan', 'None', 'NULL', ''], '')
                
            if '座机号' in df.columns:
                # 转换为字符串并清理
                df['座机号'] = df['座机号'].astype(str)
                # 只保留数字、-和()
                df['座机号'] = df['座机号'].str.replace(r'[^\d\-\(\)]', '')
                # 截断超过20位的座机号
                df['座机号'] = df['座机号'].str[:20]
                # 将无效值替换为空字符串
                df['座机号'] = df['座机号'].replace(['nan', 'None', 'NULL', ''], '')
            
            # 处理性别
            if '性别' in df.columns:
                # 转换为字符串并清理
                df['性别'] = df['性别'].astype(str).str.strip()
                # 统一性别格式
                gender_mapping = {
                    '男': '男',
                    'M': '男',
                    'MALE': '男',
                    '女': '女',
                    'F': '女',
                    'FEMALE': '女'
                }
                df['性别'] = df['性别'].map(gender_mapping)
                # 将无效值替换为默认值"男"
                df['性别'] = df['性别'].replace(['nan', 'None', 'NULL', ''], '男')
            
            # 处理身份证号
            if '身份证号' in df.columns:
                # 转换为字符串并清理
                df['身份证号'] = df['身份证号'].astype(str)
                # 只保留数字和X
                df['身份证号'] = df['身份证号'].str.replace(r'[^\dXx]', '')
                # 统一大写X
                df['身份证号'] = df['身份证号'].str.upper()
                # 将无效值替换为空字符串
                df['身份证号'] = df['身份证号'].replace(['nan', 'None', 'NULL', ''], '')
            
            # 处理出生年月日
            if '出生年月日' in df.columns:
                df['出生年月日'] = pd.to_datetime(df['出生年月日'], errors='coerce')
            
            # 处理会员卡号
            if '会员卡号' in df.columns:
                # 转换为字符串并清理
                df['会员卡号'] = df['会员卡号'].astype(str)
                # 只保留数字和字母
                df['会员卡号'] = df['会员卡号'].str.replace(r'[^\dA-Za-z]', '')
                # 截断超过20位的卡号
                df['会员卡号'] = df['会员卡号'].str[:20]
                # 将无效值替换为空字符串
                df['会员卡号'] = df['会员卡号'].replace(['nan', 'None', 'NULL', ''], '')
            
            # 处理剩余积分
            if '剩余积分' in df.columns:
                # 转换为数值类型，将无效值设为0
                df['剩余积分'] = pd.to_numeric(df['剩余积分'], errors='coerce').fillna(0)
                # 确保积分大于等于0
                df['剩余积分'] = df['剩余积分'].clip(lower=0)
                # 转换为整数
                df['剩余积分'] = df['剩余积分'].astype(int)
            
            # 处理金额字段
            if '剩余充值金额' in df.columns:
                # 转换为数值类型，将无效值设为0
                df['剩余充值金额'] = pd.to_numeric(df['剩余充值金额'], errors='coerce').fillna(0)
                # 确保金额大于等于0
                df['剩余充值金额'] = df['剩余充值金额'].clip(lower=0)
            
            if '剩余赠送金额' in df.columns:
                # 转换为数值类型，将无效值设为0
                df['剩余赠送金额'] = pd.to_numeric(df['剩余赠送金额'], errors='coerce').fillna(0)
                # 确保金额大于等于0
                df['剩余赠送金额'] = df['剩余赠送金额'].clip(lower=0)
            
            # 删除重复行
            df = df.drop_duplicates()
            
            # 处理缺失值
            fillna_dict = {
                '会员姓名': '',
                '手机号': '',
                '座机号': '',
                '性别': '男',  # 默认值设为男
                '身份证号': '',
                '出生年月日': pd.NaT,
                '联系地址': '',
                '会员卡号': '',
                '剩余积分': 0,
                '剩余充值金额': 0,
                '剩余赠送金额': 0
            }
            
            df = df.fillna(fillna_dict)
            
            # 只保留目标列
            target_columns = list(header_mapping.keys())
            df = df[target_columns]
            
            self.members_df = df
            
            # 保存处理后的数据到Excel文件
            try:
                output_path = 'members_output.xlsx'
                df.to_excel(output_path, index=False)
                
                # 生成处理报告
                report = f"成功处理会员数据，共 {len(df)} 条记录\n"
                report += f"数据已保存到 {output_path}\n\n"
                report += "找到的列：\n"
                for col in found_columns:
                    report += f"- {col}\n"
                report += "\n未找到的列（已创建并填充默认值）：\n"
                for col in header_mapping.keys():
                    if col not in found_columns:
                        report += f"- {col}\n"
                
                return report
            except Exception as save_error:
                return f"成功处理会员数据，共 {len(df)} 条记录\n但保存文件时出错：{str(save_error)}"
                
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