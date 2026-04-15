"""
理财师团队管理数据看板 - Streamlit主程序
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
from io import BytesIO
import json

from data_processor import DataProcessor
from alert_system import AlertSystem, format_alert_message

# 页面配置
st.set_page_config(
    page_title="理财师团队数据看板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        color: #1f77b4;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #6c757d;
    }
    .alert-red { background-color: #ffebee; border-left: 4px solid #f44336; }
    .alert-orange { background-color: #fff3e0; border-left: 4px solid #ff9800; }
    .alert-yellow { background-color: #fffde7; border-left: 4px solid #ffeb3b; }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
    }
</style>
""", unsafe_allow_html=True)


# 初始化数据处理器和预警系统
@st.cache_resource
def get_processor():
    return DataProcessor(data_dir="./理财师数据看板/data")

@st.cache_resource
def get_alert_system():
    return AlertSystem()

processor = get_processor()
alert_system = get_alert_system()


def load_uploaded_data(reach_file, perf_file, opp_file):
    """加载并处理上传的文件"""
    try:
        reach_df = processor.parse_reach_statistics(reach_file)
        perf_df = processor.parse_performance(perf_file)
        opp_df = processor.parse_business_opportunity(opp_file)
        return reach_df, perf_df, opp_df, None
    except Exception as e:
        return None, None, None, str(e)


def load_sample_data():
    """加载示例数据"""
    DATA_DIR = "/app/data/所有对话/主对话/用户上传"
    try:
        reach_df = processor.parse_reach_statistics(f"{DATA_DIR}/客户触达统计（by全球主RM）_20260415_1811_1776268318466_0_eijl.xlsx")
        perf_df = processor.parse_performance(f"{DATA_DIR}/出单业绩_1776268318466_1_i8zz.xlsx")
        opp_df = processor.parse_business_opportunity(f"{DATA_DIR}/广州ARK-LTC商机线索汇总_1776268318467_2_obsf.xlsx")
        return reach_df, perf_df, opp_df, None
    except Exception as e:
        return None, None, None, str(e)


def render_kpi_cards(metrics: dict):
    """渲染KPI指标卡片"""
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="团队总客户数",
            value=f"{metrics.get('总客户数', 0):,}",
            delta=None
        )
    
    with col2:
        coverage = metrics.get('平均覆盖率', 0) * 100 if metrics.get('平均覆盖率') else 0
        st.metric(
            label="平均覆盖率",
            value=f"{coverage:.1f}%",
            delta=None
        )
    
    with col3:
        amount = metrics.get('总出单金额美元', 0) / 10000
        st.metric(
            label="总出单金额(万美元)",
            value=f"{amount:,.2f}",
            delta=None
        )
    
    with col4:
        st.metric(
            label="商机总数",
            value=metrics.get('商机总数', 0),
            delta=None
        )


def render_coverage_chart(reach_df: pd.DataFrame):
    """渲染覆盖率图表"""
    if reach_df.empty:
        st.warning("暂无触达数据")
        return
    
    # 过滤有效数据并排序
    df = reach_df[reach_df['主AR姓名'].notna()].copy()
    df = df.sort_values('综合覆盖率', ascending=True)
    
    # 创建水平条形图
    fig = go.Figure()
    
    colors = ['#4CAF50' if x >= 0.8 else '#FFC107' if x >= 0.6 else '#F44336' 
              for x in df['综合覆盖率']]
    
    fig.add_trace(go.Bar(
        y=df['主AR姓名'],
        x=df['综合覆盖率'] * 100,
        orientation='h',
        marker_color=colors,
        text=[f"{x*100:.1f}%" for x in df['综合覆盖率']],
        text_position='outside'
    ))
    
    # 添加80%阈值线
    fig.add_vline(x=80, line_dash="dash", line_color="red", annotation_text="80%阈值")
    
    fig.update_layout(
        title="理财师触达覆盖率排名",
        xaxis_title="覆盖率 (%)",
        yaxis_title="理财师",
        height=max(400, len(df) * 25),
        margin=dict(l=100, r=50, t=50, b=30),
        xaxis_range=[0, 105]
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_high_quality_chart(reach_df: pd.DataFrame):
    """渲染高质量触达占比图表"""
    if reach_df.empty:
        return
    
    df = reach_df[reach_df['主AR姓名'].notna()].copy()
    df = df.sort_values('高质量客户占比', ascending=True)
    
    fig = go.Figure()
    
    colors = ['#4CAF50' if x >= 0.5 else '#FFC107' if x >= 0.3 else '#F44336' 
              for x in df['高质量客户占比']]
    
    fig.add_trace(go.Bar(
        y=df['主AR姓名'],
        x=df['高质量客户占比'] * 100,
        orientation='h',
        marker_color=colors,
        text=[f"{x*100:.1f}%" for x in df['高质量客户占比']],
        text_position='outside'
    ))
    
    fig.add_vline(x=50, line_dash="dash", line_color="orange", annotation_text="50%阈值")
    
    fig.update_layout(
        title="高质量触达占比排名",
        xaxis_title="高质量占比 (%)",
        yaxis_title="理财师",
        height=max(400, len(df) * 25),
        margin=dict(l=100, r=50, t=50, b=30),
        xaxis_range=[0, 105]
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_performance_chart(perf_df: pd.DataFrame):
    """渲染业绩图表"""
    if perf_df.empty:
        st.warning("暂无业绩数据")
        return
    
    # 按人聚合业绩
    perf_by_person = perf_df.groupby('业绩员工').agg({
        '订单募集金额美元': 'sum',
        '订单营销业绩人民币': 'sum',
        '日期时间': 'count'
    }).reset_index()
    perf_by_person.columns = ['理财师', '出单金额(万美元)', '营销业绩(万)', '出单笔数']
    perf_by_person['出单金额(万美元)'] = perf_by_person['出单金额(万美元)'] / 10000
    perf_by_person['营销业绩(万)'] = perf_by_person['营销业绩(万)'] / 10000
    perf_by_person = perf_by_person.sort_values('出单金额(万美元)', ascending=True)
    
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("出单金额排名 (万美元)", "营销业绩排名 (万人民币)"),
        horizontal_spacing=0.15
    )
    
    # 出单金额图
    fig.add_trace(
        go.Bar(
            y=perf_by_person['理财师'],
            x=perf_by_person['出单金额(万美元)'],
            orientation='h',
            marker_color='#1f77b4',
            text=[f"{x:.2f}" for x in perf_by_person['出单金额(万美元)']],
            textposition='outside',
            name='出单金额'
        ),
        row=1, col=1
    )
    
    # 营销业绩图
    fig.add_trace(
        go.Bar(
            y=perf_by_person['理财师'],
            x=perf_by_person['营销业绩(万)'],
            orientation='h',
            marker_color='#2ca02c',
            text=[f"{x:.2f}" for x in perf_by_person['营销业绩(万)']],
            textposition='outside',
            name='营销业绩'
        ),
        row=1, col=2
    )
    
    fig.update_layout(
        height=max(400, len(perf_by_person) * 25),
        showlegend=False,
        margin=dict(l=100, r=50, t=50, b=30)
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_opportunity_chart(opp_df: pd.DataFrame):
    """渲染商机图表"""
    if opp_df.empty:
        st.warning("暂无商机数据")
        return
    
    # 按人统计商机
    opp_by_person = opp_df.groupby('当前全球主AR').agg({
        '商机id': 'count',
        '预计投资原币金额/万': 'sum',
        '转化总金额RMB/万': 'sum'
    }).reset_index()
    opp_by_person.columns = ['理财师', '商机数', '预计投资(万)', '转化金额(万)']
    opp_by_person = opp_by_person.sort_values('商机数', ascending=True)
    
    # 商机温度分布
    temp_dist = opp_df.groupby(['当前全球主AR', '商机温度']).size().unstack(fill_value=0)
    
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("商机数量排名", "商机温度分布"),
        specs=[[{"type": "bar"}, {"type": "pie"}]]
    )
    
    # 商机数量图
    fig.add_trace(
        go.Bar(
            y=opp_by_person['理财师'],
            x=opp_by_person['商机数'],
            orientation='h',
            marker_color='#ff7f0e',
            text=opp_by_person['商机数'],
            textposition='outside',
            name='商机数'
        ),
        row=1, col=1
    )
    
    # 温度饼图
    temp_counts = opp_df['商机温度'].value_counts()
    fig.add_trace(
        go.Pie(
            labels=temp_counts.index,
            values=temp_counts.values,
            marker_colors=['#F44336', '#FFC107'],  # 热-红, 温-黄
            name="商机温度"
        ),
        row=1, col=2
    )
    
    fig.update_layout(
        height=400,
        showlegend=False,
        margin=dict(l=100, r=50, t=50, b=30)
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_trend_chart(perf_df: pd.DataFrame, reach_df: pd.DataFrame):
    """渲染趋势图表"""
    if perf_df.empty:
        st.info("暂无趋势数据，请上传多期数据进行对比")
        return
    
    # 按日期聚合
    perf_df['日期'] = pd.to_datetime(perf_df['日期时间']).dt.date
    daily_perf = perf_df.groupby('日期').agg({
        '订单募集金额美元': 'sum',
        '日期时间': 'count'
    }).reset_index()
    daily_perf.columns = ['日期', '出单金额', '出单笔数']
    daily_perf['出单金额'] = daily_perf['出单金额'] / 10000
    daily_perf = daily_perf.sort_values('日期')
    
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=("每日出单金额趋势 (万美元)", "每日出单笔数趋势"),
        vertical_spacing=0.15
    )
    
    fig.add_trace(
        go.Scatter(
            x=daily_perf['日期'],
            y=daily_perf['出单金额'],
            mode='lines+markers',
            name='出单金额',
            line=dict(color='#1f77b4', width=2),
            marker=dict(size=8)
        ),
        row=1, col=1
    )
    
    fig.add_trace(
        go.Bar(
            x=daily_perf['日期'],
            y=daily_perf['出单笔数'],
            name='出单笔数',
            marker_color='#2ca02c'
        ),
        row=2, col=1
    )
    
    fig.update_layout(
        height=500,
        showlegend=False,
        margin=dict(l=50, r=50, t=50, b=50)
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_comparison_chart(current_metrics: dict, previous_metrics: dict):
    """渲染环比对比图表"""
    if not previous_metrics or previous_metrics.get('总客户数', 0) == 0:
        st.info("暂无上期数据，无法进行环比对比")
        return
    
    categories = ['总客户数', '平均覆盖率\n(%)', '出单金额\n(万美元)', '商机总数']
    
    current_values = [
        current_metrics.get('总客户数', 0),
        current_metrics.get('平均覆盖率', 0) * 100,
        current_metrics.get('总出单金额美元', 0) / 10000,
        current_metrics.get('商机总数', 0)
    ]
    
    previous_values = [
        previous_metrics.get('总客户数', 0),
        previous_metrics.get('平均覆盖率', 0) * 100,
        previous_metrics.get('总出单金额美元', 0) / 10000,
        previous_metrics.get('商机总数', 0)
    ]
    
    # 计算变化率
    changes = []
    for curr, prev in zip(current_values, previous_values):
        if prev != 0:
            change = (curr - prev) / prev * 100
        else:
            change = 0
        changes.append(change)
    
    fig = go.Figure()
    
    x = list(range(len(categories)))
    width = 0.35
    
    fig.add_trace(go.Bar(
        x=[i - width/2 for i in x],
        y=previous_values,
        width=width,
        name='上期',
        marker_color='#90caf9'
    ))
    
    fig.add_trace(go.Bar(
        x=[i + width/2 for i in x],
        y=current_values,
        width=width,
        name='本期',
        marker_color='#1f77b4'
    ))
    
    # 添加变化率标注
    for i, change in enumerate(changes):
        color = '#4CAF50' if change >= 0 else '#F44336'
        symbol = '↑' if change >= 0 else '↓'
        fig.add_annotation(
            x=i,
            y=max(current_values[i], previous_values[i]) * 1.1,
            text=f"{symbol}{abs(change):.1f}%",
            showarrow=False,
            font=dict(color=color, size=12)
        )
    
    fig.update_layout(
        title="本期 vs 上期 环比对比",
        xaxis=dict(tickmode='array', tickvals=x, ticktext=categories),
        yaxis_title="数值",
        barmode='group',
        height=400,
        legend=dict(yanchor="top", y=0.99, xanchor="right", x=0.99)
    )
    
    st.plotly_chart(fig, use_container_width=True)


def render_alerts(alerts: list):
    """渲染预警列表"""
    if not alerts:
        st.success("🎉 暂无预警，所有指标正常！")
        return
    
    # 按级别分组
    red_alerts = [a for a in alerts if a.level == 'red']
    orange_alerts = [a for a in alerts if a.level == 'orange']
    yellow_alerts = [a for a in alerts if a.level == 'yellow']
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 🔴 严重预警")
        if red_alerts:
            for alert in red_alerts:
                st.error(f"**{alert.person}**\n\n{alert.reason}\n\n当前: {alert.current_value:.1f} | 阈值: {alert.threshold:.1f}")
        else:
            st.info("无严重预警")
    
    with col2:
        st.markdown("### 🟠 中度预警")
        if orange_alerts:
            for alert in orange_alerts:
                st.warning(f"**{alert.person}**\n\n{alert.reason}\n\n当前: {alert.current_value:.1f} | 阈值: {alert.threshold:.1f}")
        else:
            st.info("无中度预警")
    
    with col3:
        st.markdown("### 🟡 轻度预警")
        if yellow_alerts:
            for alert in yellow_alerts:
                st.info(f"**{alert.person}**\n\n{alert.reason}\n\n当前: {alert.current_value:.1f} | 阈值: {alert.threshold:.1f}")
        else:
            st.info("无轻度预警")


def render_person_detail(reach_df: pd.DataFrame, perf_df: pd.DataFrame, 
                         opp_df: pd.DataFrame, person: str):
    """渲染个人详情"""
    st.subheader(f"📋 {person} 详细数据")
    
    # 触达数据
    reach_data = reach_df[reach_df['主AR姓名'] == person]
    if not reach_data.empty:
        row = reach_data.iloc[0]
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("客户数", int(row['客户数']) if pd.notna(row['客户数']) else 0)
        with col2:
            coverage = row['综合覆盖率'] * 100 if pd.notna(row['综合覆盖率']) else 0
            st.metric("综合覆盖率", f"{coverage:.1f}%")
        with col3:
            hq = row['高质量触达'] if pd.notna(row['高质量触达']) else 0
            st.metric("高质量触达", int(hq))
        with col4:
            hq_ratio = row['高质量客户占比'] * 100 if pd.notna(row['高质量客户占比']) else 0
            st.metric("高质量占比", f"{hq_ratio:.1f}%")
        
        col5, col6, col7, col8 = st.columns(4)
        with col5:
            st.metric("服务记录", int(row['服务记录人数']) if pd.notna(row['服务记录人数']) else 0)
        with col6:
            st.metric("商机客户数", int(row['商机客户数']) if pd.notna(row['商机客户数']) else 0)
        with col7:
            st.metric("打款金额(USD)", f"{row['打款金额USD']:.2f}" if pd.notna(row['打款金额USD']) else "0.00")
        with col8:
            st.metric("PPL金额(USD)", f"{row['PPL金额USD']:.2f}" if pd.notna(row['PPL金额USD']) else "0.00")
    
    # 业绩数据
    st.markdown("#### 业绩数据")
    perf_data = perf_df[perf_df['业绩员工'] == person]
    if not perf_data.empty:
        perf_agg = perf_data.agg({
            '订单募集金额美元': 'sum',
            '订单营销业绩人民币': 'sum',
            '日期时间': 'count'
        })
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("出单金额(USD)", f"{perf_agg['订单募集金额美元']:,.2f}")
        with col2:
            st.metric("营销业绩(RMB)", f"{perf_agg['订单营销业绩人民币']:,.2f}")
        with col3:
            st.metric("出单笔数", int(perf_agg['日期时间']))
        
        # 显示最近订单
        st.markdown("##### 最近订单")
        recent_orders = perf_data.nlargest(5, '日期时间')[['日期时间', '订单募集金额美元', '产品大类', '营销节点']]
        st.dataframe(recent_orders, use_container_width=True)
    else:
        st.info("暂无业绩记录")
    
    # 商机数据
    st.markdown("#### 商机数据")
    opp_data = opp_df[opp_df['当前全球主AR'] == person]
    if not opp_data.empty:
        opp_agg = opp_data.agg({
            '商机id': 'count',
            '预计投资原币金额/万': 'sum',
            '转化总金额RMB/万': 'sum'
        })
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("商机总数", int(opp_agg['商机id']))
        with col2:
            st.metric("预计投资(万)", f"{opp_agg['预计投资原币金额/万']:,.2f}")
        with col3:
            st.metric("转化金额(万)", f"{opp_agg['转化总金额RMB/万']:,.2f}")
        
        # 温度分布
        temp_dist = opp_data['商机温度'].value_counts()
        if not temp_dist.empty:
            fig = go.Figure(data=[go.Pie(
                labels=temp_dist.index,
                values=temp_dist.values,
                marker_colors=['#F44336', '#FFC107'] if '热' in temp_dist.index else ['#FFC107'],
            )])
            fig.update_layout(title="商机温度分布", height=250)
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("暂无商机记录")


def main():
    """主函数"""
    
    # 标题
    st.markdown('<h1 class="main-header">📊 理财师团队数据看板</h1>', unsafe_allow_html=True)
    
    # 侧边栏 - 数据上传
    st.sidebar.header("📁 数据上传")
    
    # 模式选择
    mode = st.sidebar.radio("选择数据加载模式", ["上传文件", "加载示例数据"])
    
    reach_df, perf_df, opp_df = None, None, None
    
    if mode == "上传文件":
        st.sidebar.markdown("### 上传Excel文件")
        
        reach_file = st.sidebar.file_uploader(
            "客户触达统计",
            type=['xlsx'],
            help="客户触达统计数据"
        )
        
        perf_file = st.sidebar.file_uploader(
            "出单业绩",
            type=['xlsx'],
            help="出单业绩数据"
        )
        
        opp_file = st.sidebar.file_uploader(
            "商机线索汇总",
            type=['xlsx'],
            help="商机线索汇总数据"
        )
        
        if reach_file and perf_file and opp_file:
            with st.spinner("正在解析数据..."):
                reach_df, perf_df, opp_df, error = load_uploaded_data(
                    reach_file, perf_file, opp_file
                )
                if error:
                    st.error(f"数据解析错误: {error}")
    else:
        st.sidebar.info("正在加载示例数据...")
        with st.spinner("正在加载示例数据..."):
            reach_df, perf_df, opp_df, error = load_sample_data()
            if error:
                st.error(f"加载示例数据失败: {error}")
                st.sidebar.error(f"错误: {error}")
            else:
                st.sidebar.success("示例数据加载成功！")
    
    # 如果有数据，保存并显示
    if reach_df is not None and perf_df is not None and opp_df is not None:
        # 保存数据
        period_label = datetime.now().strftime('%Y%m%d_%H%M%S')
        processor.save_data(reach_df, perf_df, opp_df, period_label)
        
        # 加载上期数据用于对比
        reach_dfs, perf_dfs, opp_dfs = processor.load_history_data(limit=2)
        previous_reach = reach_dfs[1] if len(reach_dfs) > 1 else pd.DataFrame()
        
        # 计算指标
        metrics = processor.aggregate_team_metrics(reach_df, perf_df, opp_df)
        previous_metrics = processor.aggregate_team_metrics(
            previous_reach, 
            perf_dfs[1] if len(perf_dfs) > 1 else pd.DataFrame(),
            opp_dfs[1] if len(opp_dfs) > 1 else pd.DataFrame()
        ) if not previous_reach.empty else {}
        
        # 生成预警
        alerts = alert_system.generate_alerts(reach_df, perf_df, opp_df, previous_reach)
        alert_summary = alert_system.get_alert_summary()
        
        # 主内容区
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📈 团队总览", 
            "👤 个人详情", 
            "📉 趋势分析",
            "🏆 排名榜单",
            "⚠️ 预警中心"
        ])
        
        with tab1:
            st.header("团队关键指标")
            render_kpi_cards(metrics)
            
            st.divider()
            
            # 环比对比
            col1, col2 = st.columns([2, 1])
            with col1:
                render_comparison_chart(metrics, previous_metrics)
            with col2:
                st.markdown("### 指标变化")
                if previous_metrics:
                    for key in ['总客户数', '平均覆盖率', '总出单金额美元', '商机总数']:
                        if key in metrics and key in previous_metrics:
                            curr = metrics[key]
                            prev = previous_metrics[key]
                            if prev != 0 and key == '平均覆盖率':
                                change = (curr - prev) * 100
                                emoji = "📈" if change >= 0 else "📉"
                                st.metric(key, f"{curr*100:.1f}%", f"{emoji} {change:+.1f}%")
                            elif key == '总出单金额美元':
                                change = (curr - prev) / prev * 100 if prev != 0 else 0
                                emoji = "📈" if change >= 0 else "📉"
                                st.metric(key, f"{curr/10000:.2f}万", f"{emoji} {change:+.1f}%")
                            elif key == '商机总数':
                                change = curr - prev
                                emoji = "📈" if change >= 0 else "📉"
                                st.metric(key, int(curr), f"{emoji} {change:+d}")
                            else:
                                change = curr - prev
                                emoji = "📈" if change >= 0 else "📉"
                                st.metric(key, int(curr), f"{emoji} {change:+d}")
            
            st.divider()
            
            # 覆盖率和高质量触达
            col1, col2 = st.columns(2)
            with col1:
                render_coverage_chart(reach_df)
            with col2:
                render_high_quality_chart(reach_df)
        
        with tab2:
            st.header("个人详情查看")
            
            # 选择人员
            persons = reach_df['主AR姓名'].dropna().unique().tolist()
            selected_person = st.selectbox("选择理财师", sorted(persons))
            
            if selected_person:
                render_person_detail(reach_df, perf_df, opp_df, selected_person)
        
        with tab3:
            st.header("趋势分析")
            
            # 时间范围选择
            period = st.radio("时间维度", ["日度", "周度", "月度"], horizontal=True)
            
            render_trend_chart(perf_df, reach_df)
            
            st.divider()
            
            # 业绩趋势
            col1, col2 = st.columns(2)
            with col1:
                # 产品分布
                if not perf_df.empty:
                    product_dist = perf_df.groupby('产品大类')['订单募集金额美元'].sum().reset_index()
                    product_dist = product_dist.sort_values('订单募集金额美元', ascending=False)
                    product_dist['金额(万)'] = product_dist['订单募集金额美元'] / 10000
                    
                    fig = px.pie(
                        product_dist, 
                        values='金额(万)', 
                        names='产品大类',
                        title='产品大类分布 (金额占比)',
                        hole=0.4
                    )
                    fig.update_layout(height=350)
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # 商机状态分布
                if not opp_df.empty:
                    status_dist = opp_df['商机状态'].value_counts()
                    fig = px.pie(
                        values=status_dist.values,
                        names=status_dist.index,
                        title='商机状态分布',
                        hole=0.4,
                        color=status_dist.index,
                        color_discrete_map={
                            '进行中': '#FFC107',
                            '已成交': '#4CAF50',
                            '已放弃': '#9E9E9E'
                        }
                    )
                    fig.update_layout(height=350)
                    st.plotly_chart(fig, use_container_width=True)
        
        with tab4:
            st.header("🏆 排名榜单")
            
            # 综合排名
            personal_metrics = processor.get_personal_metrics(reach_df, perf_df, opp_df)
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### 覆盖率排名 TOP10")
                top_coverage = personal_metrics.nlargest(10, '覆盖率')[['姓名', '覆盖率', '客户数']]
                top_coverage['覆盖率'] = top_coverage['覆盖率'].apply(lambda x: f"{x*100:.1f}%" if pd.notna(x) else "N/A")
                st.dataframe(top_coverage, use_container_width=True, hide_index=True)
            
            with col2:
                st.markdown("#### 出单金额排名 TOP10")
                top_perf = personal_metrics.nlargest(10, '出单金额美元')[['姓名', '出单金额美元', '出单笔数']]
                top_perf['出单金额(万)'] = top_perf['出单金额美元'].apply(lambda x: f"{x/10000:.2f}" if pd.notna(x) else "0")
                st.dataframe(top_perf[['姓名', '出单金额(万)', '出单笔数']], use_container_width=True, hide_index=True)
            
            col3, col4 = st.columns(2)
            with col3:
                st.markdown("#### 商机数量排名 TOP10")
                top_opp = personal_metrics.nlargest(10, '商机数')[['姓名', '商机数', '预计投资']]
                top_opp['预计投资(万)'] = top_opp['预计投资'].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "0")
                st.dataframe(top_opp[['姓名', '商机数', '预计投资(万)']], use_container_width=True, hide_index=True)
            
            with col4:
                st.markdown("#### 营销业绩排名 TOP10")
                top_revenue = personal_metrics.nlargest(10, '营销业绩')[['姓名', '营销业绩', '出单笔数']]
                top_revenue['营销业绩(万)'] = top_revenue['营销业绩'].apply(lambda x: f"{x/10000:.2f}" if pd.notna(x) else "0")
                st.dataframe(top_revenue[['姓名', '营销业绩(万)', '出单笔数']], use_container_width=True, hide_index=True)
        
        with tab5:
            st.header("⚠️ 预警中心")
            
            # 预警概览
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("预警总数", alert_summary['total'])
            with col2:
                st.markdown(f"<h3 style='color:red'>🔴 {alert_summary['red']}</h3>", unsafe_allow_html=True)
            with col3:
                st.markdown(f"<h3 style='color:orange'>🟠 {alert_summary['orange']}</h3>", unsafe_allow_html=True)
            with col4:
                st.markdown(f"<h3 style='color:#FFC107'>🟡 {alert_summary['yellow']}</h3>", unsafe_allow_html=True)
            
            st.divider()
            
            # 预警详情
            render_alerts(alerts)
        
        # 预警数量显示
        if alert_summary['total'] > 0:
            st.sidebar.markdown("---")
            st.sidebar.markdown(f"### ⚠️ 预警信息")
            st.sidebar.warning(f"共 {alert_summary['total']} 条预警")
            for alert in alerts[:5]:
                emoji = {'red': '🔴', 'orange': '🟠', 'yellow': '🟡'}.get(alert.level, '⚪')
                st.sidebar.text(f"{emoji} {alert.person}: {alert.alert_type}")
    
    else:
        # 无数据时的提示
        st.info("👈 请先上传三个Excel数据文件，或选择「加载示例数据」开始使用")
        
        st.markdown("""
        ## 📋 数据文件说明
        
        ### 1. 客户触达统计
        - 文件名包含"客户触达统计"
        - 包含字段：主AR姓名、客户数、综合覆盖率、高质量触达、陪访等
        
        ### 2. 出单业绩
        - 文件名包含"出单业绩"
        - 包含字段：业绩员工、订单募集金额、订单营销业绩、产品大类等
        
        ### 3. 商机线索汇总
        - 文件名包含"商机线索汇总"
        - 包含字段：商机id、当前全球主AR、商机温度、商机状态等
        
        ## 🎯 功能模块
        
        - **团队总览**: 关键指标、覆盖率排名、环比对比
        - **个人详情**: 每个理财师的详细数据
        - **趋势分析**: 日/周/月度趋势、产品分布
        - **排名榜单**: 各维度TOP10排名
        - **预警中心**: 自动检测并展示预警信息
        """)


if __name__ == "__main__":
    main()
