"""
预警系统模块
负责检测并生成人员预警信息
"""

import pandas as pd
from typing import List, Dict
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class Alert:
    """预警信息数据类"""
    level: str  # 'red', 'orange', 'yellow'
    person: str
    alert_type: str
    reason: str
    current_value: float
    threshold: float
    metric_name: str
    
    def to_dict(self) -> dict:
        return asdict(self)


class AlertSystem:
    """预警系统类"""
    
    # 预警阈值配置
    THRESHOLDS = {
        'coverage': 0.80,           # 触达覆盖率 < 80% 预警
        'high_quality_ratio': 0.50, # 高质量触达占比 < 50% 预警
        'cold_opportunity': 5,      # 冷商机数量阈值
        'rank_drop': 3,             # 排名下降超过3名预警
    }
    
    def __init__(self):
        self.alerts: List[Alert] = []
    
    def reset(self):
        """重置预警列表"""
        self.alerts = []
    
    def check_coverage_alert(self, reach_df: pd.DataFrame) -> List[Alert]:
        """
        检查触达覆盖率预警
        触达覆盖率 < 80% 预警
        """
        alerts = []
        for _, row in reach_df.iterrows():
            coverage = row.get('综合覆盖率', 0)
            if coverage is not None and coverage < self.THRESHOLDS['coverage']:
                alerts.append(Alert(
                    level='red' if coverage < 0.6 else 'orange',
                    person=row['主AR姓名'],
                    alert_type='覆盖率预警',
                    reason=f'触达覆盖率低于{self.THRESHOLDS["coverage"]*100:.0f}%阈值',
                    current_value=coverage * 100,
                    threshold=self.THRESHOLDS['coverage'] * 100,
                    metric_name='综合覆盖率'
                ))
        return alerts
    
    def check_no_order_alert(self, perf_df: pd.DataFrame, weeks: int = 2) -> List[Alert]:
        """
        检查连续无出单预警
        连续2周无出单预警
        """
        alerts = []
        
        # 获取有订单的员工
        employees_with_orders = perf_df['业绩员工'].unique() if not perf_df.empty else []
        
        # 获取所有员工名单（从触达数据）
        all_employees = perf_df['业绩员工'].unique() if not perf_df.empty else []
        
        # 这里简化处理，实际应该检查历史数据来判断连续无出单
        # 暂时标记为需要关注
        
        return alerts
    
    def check_high_quality_alert(self, reach_df: pd.DataFrame) -> List[Alert]:
        """
        检查高质量触达占比预警
        高质量触达占比 < 50% 预警
        """
        alerts = []
        for _, row in reach_df.iterrows():
            ratio = row.get('高质量客户占比', 0)
            if ratio is not None and ratio < self.THRESHOLDS['high_quality_ratio']:
                alerts.append(Alert(
                    level='yellow',
                    person=row['主AR姓名'],
                    alert_type='质量预警',
                    reason=f'高质量触达占比低于{self.THRESHOLDS["high_quality_ratio"]*100:.0f}%阈值',
                    current_value=ratio * 100,
                    threshold=self.THRESHOLDS['high_quality_ratio'] * 100,
                    metric_name='高质量客户占比'
                ))
        return alerts
    
    def check_opportunity_alert(self, opp_df: pd.DataFrame) -> List[Alert]:
        """
        检查商机温度预警
        商机温度"冷"数量过多预警
        """
        alerts = []
        
        # 按人统计商机温度
        if not opp_df.empty and '商机温度' in opp_df.columns:
            # 计算每个人的商机统计
            person_stats = opp_df.groupby('当前全球主AR').agg({
                '商机id': 'count',
                '预计投资原币金额/万': 'sum'
            }).reset_index()
            person_stats.columns = ['姓名', '商机总数', '预计投资金额']
            
            # 统计冷商机（假设商机温度只有 热/温，没有冷）
            # 如果某人的商机数少且预计投资金额低，可视为"冷"
            cold_threshold = 3  # 少于3个商机视为冷
            
            for _, row in person_stats.iterrows():
                if row['商机总数'] < cold_threshold:
                    alerts.append(Alert(
                        level='yellow',
                        person=row['姓名'],
                        alert_type='商机预警',
                        reason=f'商机数量过少（少于{cold_threshold}个）',
                        current_value=row['商机总数'],
                        threshold=cold_threshold,
                        metric_name='商机数量'
                    ))
        
        return alerts
    
    def check_rank_change_alert(self, current_df: pd.DataFrame, 
                                previous_df: pd.DataFrame = None) -> List[Alert]:
        """
        检查排名变化预警
        排名下降超过3名预警
        """
        alerts = []
        
        if previous_df is None or previous_df.empty:
            return alerts
        
        # 按覆盖率排名
        current_rank = current_df.sort_values('综合覆盖率', ascending=False).reset_index(drop=True)
        current_rank['当前排名'] = current_rank.index + 1
        
        previous_rank = previous_df.sort_values('综合覆盖率', ascending=False).reset_index(drop=True)
        previous_rank['上期排名'] = previous_rank.index + 1
        
        # 合并排名
        merged = current_rank.merge(previous_rank[['主AR姓名', '上期排名']], on='主AR姓名', how='left')
        merged['排名变化'] = merged['上期排名'] - merged['当前排名']  # 正数表示下降
        
        # 检查排名下降
        for _, row in merged.iterrows():
            if pd.notna(row['上期排名']) and row['排名变化'] > self.THRESHOLDS['rank_drop']:
                alerts.append(Alert(
                    level='orange',
                    person=row['主AR姓名'],
                    alert_type='排名预警',
                    reason=f'覆盖率排名下降{int(row["排名变化"])}位（从上期第{int(row["上期排名"])}名降至第{int(row["当前排名"])}名）',
                    current_value=row['当前排名'],
                    threshold=self.THRESHOLDS['rank_drop'],
                    metric_name='综合覆盖率排名'
                ))
        
        return alerts
    
    def check_performance_alert(self, perf_df: pd.DataFrame, 
                                 reach_df: pd.DataFrame) -> List[Alert]:
        """
        检查业绩预警
        无出单人员预警
        """
        alerts = []
        
        if perf_df.empty:
            # 如果没有业绩数据，标记所有触达人员
            for _, row in reach_df.iterrows():
                alerts.append(Alert(
                    level='orange',
                    person=row['主AR姓名'],
                    alert_type='业绩预警',
                    reason='本期暂无出单记录',
                    current_value=0,
                    threshold=1,
                    metric_name='出单笔数'
                ))
            return alerts
        
        # 找出没有出单的人员
        employees_with_orders = set(perf_df['业绩员工'].unique())
        all_employees = set(reach_df['主AR姓名'].unique())
        no_order_employees = all_employees - employees_with_orders
        
        for name in no_order_employees:
            alerts.append(Alert(
                level='orange',
                person=name,
                alert_type='业绩预警',
                reason='本期暂无出单记录',
                current_value=0,
                threshold=1,
                metric_name='出单笔数'
            ))
        
        return alerts
    
    def generate_alerts(self, reach_df: pd.DataFrame, perf_df: pd.DataFrame, 
                        opp_df: pd.DataFrame, previous_df: pd.DataFrame = None) -> List[Alert]:
        """
        生成所有预警
        """
        self.reset()
        
        # 覆盖率预警
        self.alerts.extend(self.check_coverage_alert(reach_df))
        
        # 高质量占比预警
        self.alerts.extend(self.check_high_quality_alert(reach_df))
        
        # 商机预警
        self.alerts.extend(self.check_opportunity_alert(opp_df))
        
        # 业绩预警
        self.alerts.extend(self.check_performance_alert(perf_df, reach_df))
        
        # 排名变化预警
        if previous_df is not None and not previous_df.empty:
            self.alerts.extend(self.check_rank_change_alert(reach_df, previous_df))
        
        # 按级别排序
        level_order = {'red': 0, 'orange': 1, 'yellow': 2}
        self.alerts.sort(key=lambda x: (level_order.get(x.level, 3), str(x.person)))
        
        return self.alerts
    
    def get_alert_summary(self) -> Dict:
        """
        获取预警汇总信息
        """
        summary = {
            'total': len(self.alerts),
            'red': len([a for a in self.alerts if a.level == 'red']),
            'orange': len([a for a in self.alerts if a.level == 'orange']),
            'yellow': len([a for a in self.alerts if a.level == 'yellow']),
            'by_type': {}
        }
        
        for alert in self.alerts:
            if alert.alert_type not in summary['by_type']:
                summary['by_type'][alert.alert_type] = 0
            summary['by_type'][alert.alert_type] += 1
        
        return summary
    
    def get_alerts_by_person(self, person: str = None) -> List[Alert]:
        """
        获取指定人员的预警
        """
        if person:
            return [a for a in self.alerts if a.person == person]
        return self.alerts


def format_alert_message(alert: Alert) -> str:
    """格式化预警消息"""
    emoji = {
        'red': '🔴',
        'orange': '🟠',
        'yellow': '🟡'
    }.get(alert.level, '⚪')
    
    return f"{emoji} **{alert.person}** - {alert.alert_type}\n{alert.reason}\n当前值: {alert.current_value:.1f}, 阈值: {alert.threshold:.1f}"
