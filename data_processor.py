"""
数据处理模块
负责解析和转换三个Excel数据源
"""

import pandas as pd
import json
import os
from datetime import datetime
from pathlib import Path


class DataProcessor:
    """数据处理器类"""
    
    def __init__(self, data_dir: str = "./data"):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        
    def parse_reach_statistics(self, file_path: str) -> pd.DataFrame:
        """
        解析客户触达统计数据
        第0行是主标题行，第1行是子标题行，需跳过
        """
        df = pd.read_excel(file_path, header=None)
        
        # 获取有效数据（跳过第0行和第1行子标题）
        df_valid = df.iloc[2:].copy()
        
        # 设置列名 - 使用行0作为主标题
        columns = {
            0: '主AR大区',
            1: '主AR片区', 
            2: '主AR姓名',
            3: '客户数',
            4: '综合覆盖率',
            5: '服务次数',
            6: '服务客户数',
            7: '高质量触达',
            8: '高质量客户占比',
            9: '陪访客户数',
            10: '陪访覆盖率',
            11: '活动人数',
            12: '活动覆盖率',
            13: '存续客户数',
            14: '存续覆盖率',
            15: '服务记录人数',
            16: '服务记录覆盖率',
            17: '商机客户数',
            18: '商机/触达占比',
            19: 'PPL金额USD',
            20: '打款客户数',
            21: '打款/触达',
            22: '打款金额USD',
            23: '海投打款金额USD',
            24: '主AR_ID'
        }
        df_valid = df_valid.rename(columns=columns)
        
        # 转换数值类型
        numeric_cols = ['客户数', '服务次数', '服务客户数', '高质量触达', '陪访客户数',
                       '活动人数', '存续客户数', '服务记录人数', '商机客户数', 
                       '打款客户数', '主AR_ID']
        for col in numeric_cols:
            if col in df_valid.columns:
                df_valid[col] = pd.to_numeric(df_valid[col], errors='coerce')
        
        # 转换覆盖率百分比
        def parse_percentage(x):
            if pd.isna(x):
                return None
            if isinstance(x, str):
                # 处理 "% 覆盖率" 这种格式
                x_clean = x.replace('%', '').replace(' ', '').strip()
                try:
                    return float(x_clean) / 100 if x_clean and x_clean[0].isdigit() else None
                except:
                    return None
            try:
                return float(x)
            except:
                return None
        
        for col in ['综合覆盖率', '高质量客户占比', '陪访覆盖率', '活动覆盖率', 
                    '存续覆盖率', '服务记录覆盖率', '商机/触达占比', '打款/触达']:
            if col in df_valid.columns:
                df_valid[col] = df_valid[col].apply(parse_percentage)
        
        # 添加时间戳
        df_valid['数据时间'] = datetime.now().strftime('%Y-%m-%d')
        
        return df_valid[['主AR大区', '主AR片区', '主AR姓名', '客户数', '综合覆盖率',
                         '高质量触达', '高质量客户占比', '陪访客户数', '陪访覆盖率',
                         '活动人数', '活动覆盖率', '存续客户数', '存续覆盖率',
                         '服务记录人数', '服务记录覆盖率', '商机客户数', 
                         '商机/触达占比', 'PPL金额USD', '打款客户数', 
                         '打款/触达', '打款金额USD', '海投打款金额USD', '数据时间']]

    def parse_performance(self, file_path: str) -> pd.DataFrame:
        """
        解析出单业绩数据
        """
        df = pd.read_excel(file_path)
        
        # 提取关键字段
        df_valid = df[['日期时间', '业绩员工', '订单募集金额美元', '订单营销业绩人民币', 
                       '产品大类', '营销节点', '营销年']].copy()
        
        # 转换日期
        df_valid['日期时间'] = pd.to_datetime(df_valid['日期时间'], errors='coerce')
        
        # 转换数值
        df_valid['订单募集金额美元'] = pd.to_numeric(df_valid['订单募集金额美元'], errors='coerce')
        df_valid['订单营销业绩人民币'] = pd.to_numeric(df_valid['订单营销业绩人民币'], errors='coerce')
        
        # 添加时间戳
        df_valid['数据时间'] = datetime.now().strftime('%Y-%m-%d')
        
        return df_valid

    def parse_business_opportunity(self, file_path: str) -> pd.DataFrame:
        """
        解析商机线索汇总数据
        使用RAW sheet的明细数据
        """
        df = pd.read_excel(file_path, sheet_name='RAW')
        
        # 提取关键字段
        df_valid = df[['商机id', '商机创建时间', '当前全球主AR片区', '当前全球主AR',
                       '会员等级', '产品大类', '意向产品', '预计投资原币金额/万',
                       '币种', '商机温度', '商机状态', '转化总金额RMB/万',
                       '创建人', '商机进展', '最后活跃时间']].copy()
        
        # 转换日期
        df_valid['商机创建时间'] = pd.to_datetime(df_valid['商机创建时间'], errors='coerce')
        
        # 转换数值
        df_valid['预计投资原币金额/万'] = pd.to_numeric(df_valid['预计投资原币金额/万'], errors='coerce')
        df_valid['转化总金额RMB/万'] = pd.to_numeric(df_valid['转化总金额RMB/万'], errors='coerce')
        
        # 添加时间戳
        df_valid['数据时间'] = datetime.now().strftime('%Y-%m-%d')
        
        return df_valid

    def save_data(self, reach_df: pd.DataFrame, performance_df: pd.DataFrame, 
                  opportunity_df: pd.DataFrame, period_label: str = None):
        """
        保存数据到历史存储
        """
        if period_label is None:
            period_label = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # 保存为JSON文件
        reach_df.to_json(self.data_dir / f"reach_{period_label}.json", orient='records', force_ascii=False)
        performance_df.to_json(self.data_dir / f"performance_{period_label}.json", orient='records', force_ascii=False)
        opportunity_df.to_json(self.data_dir / f"opportunity_{period_label}.json", orient='records', force_ascii=False)
        
        return period_label

    def load_latest_data(self):
        """
        加载最新的历史数据
        """
        reach_files = sorted(self.data_dir.glob("reach_*.json"), reverse=True)
        perf_files = sorted(self.data_dir.glob("performance_*.json"), reverse=True)
        opp_files = sorted(self.data_dir.glob("opportunity_*.json"), reverse=True)
        
        reach_df = pd.read_json(reach_files[0]) if reach_files else pd.DataFrame()
        perf_df = pd.read_json(perf_files[0]) if perf_files else pd.DataFrame()
        opp_df = pd.read_json(opp_files[0]) if opp_files else pd.DataFrame()
        
        return reach_df, perf_df, opp_df

    def load_history_data(self, limit: int = 10):
        """
        加载历史数据用于趋势分析
        """
        reach_files = sorted(self.data_dir.glob("reach_*.json"), reverse=True)[:limit]
        perf_files = sorted(self.data_dir.glob("performance_*.json"), reverse=True)[:limit]
        opp_files = sorted(self.data_dir.glob("opportunity_*.json"), reverse=True)[:limit]
        
        reach_dfs = [pd.read_json(f) for f in reach_files]
        perf_dfs = [pd.read_json(f) for f in perf_files]
        opp_dfs = [pd.read_json(f) for f in opp_files]
        
        return reach_dfs, perf_dfs, opp_dfs

    def aggregate_team_metrics(self, reach_df: pd.DataFrame, perf_df: pd.DataFrame, 
                                opp_df: pd.DataFrame) -> dict:
        """
        聚合团队关键指标
        """
        metrics = {}
        
        # 触达统计汇总
        if not reach_df.empty:
            metrics['总客户数'] = int(reach_df['客户数'].sum())
            metrics['平均覆盖率'] = float(reach_df['综合覆盖率'].mean())
            metrics['平均高质量占比'] = float(reach_df['高质量客户占比'].mean())
            metrics['总服务客户数'] = int(reach_df['服务客户数'].sum()) if '服务客户数' in reach_df.columns else 0
            metrics['总商机数'] = int(reach_df['商机客户数'].sum())
        
        # 业绩汇总
        if not perf_df.empty:
            metrics['总出单金额美元'] = float(perf_df['订单募集金额美元'].sum())
            metrics['总营销业绩人民币'] = float(perf_df['订单营销业绩人民币'].sum())
            metrics['出单笔数'] = len(perf_df)
            metrics['出单人数'] = perf_df['业绩员工'].nunique()
        
        # 商机汇总
        if not opp_df.empty:
            metrics['商机总数'] = len(opp_df)
            metrics['热商机数'] = len(opp_df[opp_df['商机温度'] == '热'])
            metrics['温商机数'] = len(opp_df[opp_df['商机温度'] == '温'])
            metrics['商机转化金额'] = float(opp_df['转化总金额RMB/万'].sum())
            metrics['商机预计投资'] = float(opp_df['预计投资原币金额/万'].sum())
        
        return metrics

    def get_personal_metrics(self, reach_df: pd.DataFrame, perf_df: pd.DataFrame, 
                             opp_df: pd.DataFrame) -> pd.DataFrame:
        """
        获取个人指标数据
        """
        # 触达数据按人聚合
        reach_metrics = reach_df.groupby('主AR姓名').agg({
            '客户数': 'sum',
            '综合覆盖率': 'mean',
            '高质量触达': 'sum',
            '高质量客户占比': 'mean',
            '服务记录人数': 'sum',
            '商机客户数': 'sum',
            '打款金额USD': 'sum'
        }).reset_index()
        reach_metrics.columns = ['姓名', '客户数', '覆盖率', '高质量触达', 
                                 '高质量占比', '服务人数', '商机数', '打款金额']
        
        # 业绩数据按人聚合
        perf_metrics = perf_df.groupby('业绩员工').agg({
            '订单募集金额美元': 'sum',
            '订单营销业绩人民币': 'sum',
            '日期时间': 'count'
        }).reset_index()
        perf_metrics.columns = ['姓名', '出单金额美元', '营销业绩', '出单笔数']
        
        # 商机数据按人聚合
        opp_metrics = opp_df.groupby('当前全球主AR').agg({
            '商机id': 'count',
            '预计投资原币金额/万': 'sum',
            '转化总金额RMB/万': 'sum'
        }).reset_index()
        opp_metrics.columns = ['姓名', '商机数', '预计投资', '转化金额']
        
        # 合并所有指标
        result = reach_metrics.merge(perf_metrics, on='姓名', how='outer')
        result = result.merge(opp_metrics, on='姓名', how='outer')
        result = result.fillna(0)
        
        return result

    def get_period_comparison(self, current_df: pd.DataFrame, 
                               current_date: str, comparison_period: str = 'week') -> dict:
        """
        获取周期对比数据
        """
        # 获取对比数据
        reach_files = sorted(self.data_dir.glob("reach_*.json"), reverse=True)
        
        # 简单实现：返回当前数据与上一份数据的对比
        if len(reach_files) >= 2:
            previous_df = pd.read_json(reach_files[1])
            
            comparison = {
                'current_total': int(current_df['客户数'].sum()) if not current_df.empty else 0,
                'previous_total': int(previous_df['客户数'].sum()) if not previous_df.empty else 0,
                'current_coverage': float(current_df['综合覆盖率'].mean()) if not current_df.empty else 0,
                'previous_coverage': float(previous_df['综合覆盖率'].mean()) if not previous_df.empty else 0
            }
            
            comparison['coverage_change'] = comparison['current_coverage'] - comparison['previous_coverage']
            comparison['customer_change'] = comparison['current_total'] - comparison['previous_total']
        else:
            comparison = {
                'current_total': 0,
                'previous_total': 0,
                'current_coverage': 0,
                'previous_coverage': 0,
                'coverage_change': 0,
                'customer_change': 0
            }
        
        return comparison
