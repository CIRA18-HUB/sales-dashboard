import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import traceback

# 设置页面配置
st.set_page_config(
    page_title="销售数据分析仪表盘",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 定义一些更美观的自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.8rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 2rem;
        padding: 1.5rem;
        background-color: #f8f9fa;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .sub-header {
        font-size: 1.8rem;
        color: #0D47A1;
        padding-top: 1.5rem;
        padding-bottom: 1rem;
        margin-top: 1rem;
        border-bottom: 2px solid #E3F2FD;
    }
    .card {
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        background-color: white;
        transition: transform 0.3s;
    }
    .card:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
    }
    .metric-value {
        font-size: 2.2rem;
        font-weight: bold;
        color: #1E88E5;
        margin: 0.5rem 0;
    }
    .metric-label {
        font-size: 1.1rem;
        color: #424242;
        font-weight: 500;
    }
    .highlight {
        background-color: #E3F2FD;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1.5rem 0;
        border-left: 5px solid #1E88E5;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        border-radius: 5px 5px 0 0;
    }
    .stTabs [aria-selected="true"] {
        background-color: #E3F2FD;
        border-bottom: 3px solid #1E88E5;
    }
    .stExpander {
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }
    .download-button {
        text-align: center;
        margin-top: 2rem;
    }
    .section-gap {
        margin-top: 2.5rem;
        margin-bottom: 2.5rem;
    }
    /* 调整图表容器的样式 */
    .st-emotion-cache-1wrcr25 {
        margin-top: 2rem !important;
        margin-bottom: 3rem !important;
        padding: 1rem !important;
    }
    /* 设置侧边栏样式 */
    .st-emotion-cache-6qob1r {
        background-color: #f5f7fa;
        border-right: 1px solid #e0e0e0;
    }
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
    }
    [data-testid="stSidebarNav"] {
        padding-top: 2rem;
    }
    .sidebar-header {
        font-size: 1.3rem;
        color: #0D47A1;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #e0e0e0;
    }
    /* 调整图表字体大小 */
    .js-plotly-plot .plotly .ytick text, 
    .js-plotly-plot .plotly .xtick text {
        font-size: 14px !important;
    }
    .js-plotly-plot .plotly .gtitle {
        font-size: 18px !important;
    }
    /* 错误消息样式 */
    .error-message {
        color: #721c24;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    /* 信息消息样式 */
    .info-message {
        color: #0c5460;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# 标题
st.markdown('<div class="main-header">销售数据分析仪表盘</div>', unsafe_allow_html=True)


# 格式化数值的函数
def format_yuan(value):
    if value >= 10000:
        return f"{value / 10000:.1f}万元"
    return f"{value:,.0f}元"


# 安全筛选函数
def safe_filter(df, filter_func):
    """安全应用筛选条件，确保筛选后的数据不为空"""
    try:
        filtered = filter_func(df)
        if filtered.empty:
            st.warning("当前筛选条件下没有匹配的数据。请尝试放宽筛选条件。")
            return df  # 返回原始数据而不是空数据
        return filtered
    except Exception as e:
        st.error(f"筛选数据时出错: {str(e)}")
        return df


# 加载数据函数
@st.cache_data
def load_data(file_path=None):
    # 如果提供了文件路径，从文件加载
    try:
        if file_path is not None:
            try:
                # 检查是否是FileUploader对象还是字符串路径
                if hasattr(file_path, 'read'):
                    # 是上传的文件对象
                    df = pd.read_excel(file_path)
                    st.sidebar.success(f"文件加载成功！")
                else:
                    # 是文件路径字符串
                    df = pd.read_excel(file_path)
            except Exception as e:
                st.error(f"文件加载失败: {str(e)}。使用示例数据进行演示。")
                df = load_sample_data()
        else:
            df = load_sample_data()

        # 数据预处理
        df['销售额'] = df['单价（箱）'] * df['数量（箱）']

        # 确保发运月份是日期类型
        try:
            df['发运月份'] = pd.to_datetime(df['发运月份'])
        except Exception as e:
            st.info(f"发运月份转换为日期类型时出错。原因：{str(e)}。将保持原格式。")

        # 添加简化产品名称列
        df['简化产品名称'] = df.apply(lambda row: get_simplified_product_name(row['产品代码'], row['产品名称']), axis=1)

        return df
    except Exception as e:
        st.error(f"加载数据时出现未预期的错误: {str(e)}")
        st.write("错误详情:")
        st.write(traceback.format_exc())
        return load_sample_data()


# 创建产品代码到简化产品名称的映射函数 (修复版)
def get_simplified_product_name(product_code, product_name):
    try:
        # 从产品名称中提取关键部分
        if '口力' in product_name:
            # 提取"口力"之后的产品类型
            name_parts = product_name.split('口力')[1].split('-')[0].strip()
            # 进一步简化，只保留主要部分（去掉规格和包装形式）
            for suffix in ['G分享装袋装', 'G盒装', 'G袋装', 'KG迷你包', 'KG随手包']:
                name_parts = name_parts.split(suffix)[0]

            # 去掉可能的数字和单位
            import re
            simple_name = re.sub(r'\d+\w*\s*', '', name_parts).strip()

            # 始终包含产品代码以确保唯一性
            return f"{simple_name} ({product_code})"
        else:
            # 如果无法提取，则返回产品代码
            return product_code
    except Exception as e:
        # 出错时返回产品代码
        return product_code


# 创建示例数据（以防用户没有上传文件）
@st.cache_data
def load_sample_data():
    # 创建简化版示例数据
    data = {
        '客户简称': ['广州佳成行', '广州佳成行', '广州佳成行', '广州佳成行', '广州佳成行',
                     '广州佳成行', '河南甜丰號', '河南甜丰號', '河南甜丰號', '河南甜丰號',
                     '河南甜丰號', '广州佳成行', '河南甜丰號', '广州佳成行', '河南甜丰號',
                     '广州佳成行'],
        '所属区域': ['南', '南', '南', '南', '南', '南', '中', '中', '中', '中', '中',
                     '南', '中', '南', '中', '南'],
        '发运月份': ['2025-03', '2025-03', '2025-03', '2025-03', '2025-03', '2025-03',
                     '2025-03', '2025-03', '2025-03', '2025-03', '2025-03', '2025-03',
                     '2025-03', '2025-03', '2025-03', '2025-03'],
        '申请人': ['梁洪泽', '梁洪泽', '梁洪泽', '梁洪泽', '梁洪泽', '梁洪泽',
                   '胡斌', '胡斌', '胡斌', '胡斌', '胡斌', '梁洪泽', '胡斌', '梁洪泽',
                   '胡斌', '梁洪泽'],
        '产品代码': ['F3415D', 'F3421D', 'F0104J', 'F0104L', 'F3411A', 'F01E4B',
                     'F01L4C', 'F01C2P', 'F01E6D', 'F3450B', 'F3415B', 'F0110C',
                     'F0183F', 'F01K8A', 'F0183K', 'F0101P'],
        '产品名称': ['口力酸小虫250G分享装袋装-中国', '口力可乐瓶250G分享装袋装-中国',
                     '口力比萨XXL45G盒装-中国', '口力比萨68G袋装-中国', '口力午餐袋77G袋装-中国',
                     '口力汉堡108G袋装-中国', '口力扭扭虫2KG迷你包-中国', '口力字节软糖2KG迷你包-中国',
                     '口力西瓜1.5KG随手包-中国', '口力七彩熊1.5KG随手包-中国', '口力酸小虫1.5KG随手包-中国',
                     '口力软糖新品A-中国', '口力软糖新品B-中国', '口力软糖新品C-中国', '口力软糖新品D-中国',
                     '口力软糖新品E-中国'],
        '订单类型': ['订单-正常产品'] * 16,
        '单价（箱）': [121.44, 121.44, 216.96, 126.72, 137.04, 137.04, 127.2, 127.2,
                     180, 180, 180, 150, 160, 170, 180, 190],
        '数量（箱）': [10, 10, 20, 50, 252, 204, 7, 2, 6, 6, 6, 30, 20, 15, 10, 5]
    }

    df = pd.DataFrame(data)
    return df


# 侧边栏 - 上传文件区域
st.sidebar.markdown('<div class="sidebar-header">数据导入</div>', unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("上传Excel销售数据文件", type=["xlsx", "xls"])

# 加载数据
try:
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        st.sidebar.success(f"已成功加载文件: {uploaded_file.name}")
    else:
        # 使用示例数据进行演示
        df = load_data()
        st.sidebar.info("正在使用示例数据。请上传您的数据文件获取真实分析。")
except Exception as e:
    st.error(f"加载数据时出错: {str(e)}")
    df = load_sample_data()
    st.sidebar.warning("由于错误，使用示例数据进行演示。请检查您的数据文件格式。")

# 显示数据预览
with st.expander("数据预览", expanded=False):
    st.write("以下是加载的数据的前几行：")
    st.dataframe(df.head())
    st.write(f"总行数: {len(df)}")
    st.write(f"列名: {', '.join(df.columns)}")

# 定义新品产品代码
new_products = ['F0110C', 'F0183F', 'F01K8A', 'F0183K', 'F0101P']
new_products_df = df[df['产品代码'].isin(new_products)]

# 创建产品代码到简化名称的映射字典（用于图表显示）
product_name_mapping = {
    code: df[df['产品代码'] == code]['简化产品名称'].iloc[0] if len(df[df['产品代码'] == code]) > 0 else code
    for code in df['产品代码'].unique()
}

# 侧边栏 - 筛选器
st.sidebar.markdown('<div class="sidebar-header">筛选数据</div>', unsafe_allow_html=True)

# 区域筛选器
all_regions = sorted(df['所属区域'].astype(str).unique())
selected_regions = st.sidebar.multiselect("选择区域", all_regions, default=all_regions)

# 客户筛选器
all_customers = sorted(df['客户简称'].astype(str).unique())
selected_customers = st.sidebar.multiselect("选择客户", all_customers, default=[])

# 产品代码筛选器
all_products = sorted(df['产品代码'].astype(str).unique())
product_options = [(code, product_name_mapping[code]) for code in all_products]
selected_products = st.sidebar.multiselect(
    "选择产品",
    options=all_products,
    format_func=lambda x: f"{x} ({product_name_mapping[x]})",
    default=[]
)

# 申请人筛选器
all_applicants = sorted(df['申请人'].astype(str).unique())
selected_applicants = st.sidebar.multiselect("选择申请人", all_applicants, default=[])

# 应用筛选条件
filtered_df = df.copy()

if selected_regions:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['所属区域'].isin(selected_regions)])

if selected_customers:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['客户简称'].isin(selected_customers)])

if selected_products:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['产品代码'].isin(selected_products)])

if selected_applicants:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['申请人'].isin(selected_applicants)])

# 检查筛选后是否还有数据
if filtered_df.empty:
    st.error("应用所有筛选条件后没有匹配的数据。请调整筛选条件。")
    # 重置为原始数据
    filtered_df = df.copy()
    st.warning("已重置为原始数据。")

# 根据筛选后的数据筛选新品数据
filtered_new_products_df = filtered_df[filtered_df['产品代码'].isin(new_products)]

# 导航栏
st.markdown('<div class="sub-header">导航</div>', unsafe_allow_html=True)
tabs = st.tabs(["销售概览", "新品分析", "客户细分", "产品组合", "市场渗透率"])

with tabs[0]:  # 销售概览
    # KPI指标行
    st.markdown('<div class="sub-header">🔑 关键绩效指标</div>', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)

    try:
        total_sales = filtered_df['销售额'].sum()
        with col1:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">总销售额</div>
                <div class="metric-value">{format_yuan(total_sales)}</div>
            </div>
            """, unsafe_allow_html=True)

        total_customers = filtered_df['客户简称'].nunique()
        with col2:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">客户数量</div>
                <div class="metric-value">{total_customers}</div>
            </div>
            """, unsafe_allow_html=True)

        total_products = filtered_df['产品代码'].nunique()
        with col3:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">产品数量</div>
                <div class="metric-value">{total_products}</div>
            </div>
            """, unsafe_allow_html=True)

        avg_price = filtered_df['单价（箱）'].mean()
        with col4:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">平均单价</div>
                <div class="metric-value">{avg_price:.2f}元</div>
            </div>
            """, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"计算KPI指标时出错: {str(e)}")

    # 区域销售分析
    st.markdown('<div class="sub-header section-gap">📊 区域销售分析</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    try:
        # 区域销售额柱状图
        region_sales = filtered_df.groupby('所属区域')['销售额'].sum().reset_index()

        if not region_sales.empty:
            with col1:
                fig_region = px.bar(
                    region_sales,
                    x='所属区域',
                    y='销售额',
                    color='所属区域',
                    title='各区域销售额',
                    labels={'销售额': '销售额 (元)', '所属区域': '区域'},
                    height=500,
                    color_discrete_sequence=px.colors.qualitative.Bold
                )
                # 添加文本标签
                fig_region.update_traces(
                    text=[format_yuan(val) for val in region_sales['销售额']],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_region.update_layout(
                    xaxis_title=dict(text="区域", font=dict(size=16)),
                    yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_region.update_yaxes(
                    range=[0, region_sales['销售额'].max() * 1.2]
                )
                st.plotly_chart(fig_region, use_container_width=True)

            with col2:
                # 区域销售占比饼图
                fig_region_pie = px.pie(
                    region_sales,
                    values='销售额',
                    names='所属区域',
                    title='各区域销售占比',
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Bold
                )
                fig_region_pie.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    textfont=dict(size=14)
                )
                fig_region_pie.update_layout(
                    margin=dict(t=60, b=60, l=60, r=60),
                    font=dict(size=14)
                )
                st.plotly_chart(fig_region_pie, use_container_width=True)
        else:
            st.warning("没有足够的区域销售数据来创建图表。")
    except Exception as e:
        st.error(f"创建区域销售分析图表时出错: {str(e)}")

    # 产品销售分析
    st.markdown('<div class="sub-header section-gap">📦 产品销售分析</div>', unsafe_allow_html=True)


    # 提取包装类型
    def extract_packaging(product_name):
        try:
            if '袋装' in product_name:
                return '袋装'
            elif '盒装' in product_name:
                return '盒装'
            elif '随手包' in product_name:
                return '随手包'
            elif '迷你包' in product_name:
                return '迷你包'
            elif '分享装' in product_name:
                return '分享装'
            else:
                return '其他'
        except:
            return '其他'


    try:
        filtered_df['包装类型'] = filtered_df['产品名称'].apply(extract_packaging)
        packaging_sales = filtered_df.groupby('包装类型')['销售额'].sum().reset_index()

        col1, col2 = st.columns(2)

        if not packaging_sales.empty:
            with col1:
                # 包装类型销售额柱状图
                fig_packaging = px.bar(
                    packaging_sales.sort_values(by='销售额', ascending=False),
                    x='包装类型',
                    y='销售额',
                    color='包装类型',
                    title='不同包装类型销售额',
                    labels={'销售额': '销售额 (元)', '包装类型': '包装类型'},
                    height=500
                )
                # 添加文本标签
                fig_packaging.update_traces(
                    text=[format_yuan(val) for val in packaging_sales['销售额']],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_packaging.update_layout(
                    xaxis_title=dict(text="包装类型", font=dict(size=16)),
                    yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_packaging.update_yaxes(
                    range=[0, packaging_sales['销售额'].max() * 1.2]
                )
                st.plotly_chart(fig_packaging, use_container_width=True)

        with col2:
            # 价格-销量散点图
            try:
                fig_price_qty = px.scatter(
                    filtered_df,
                    x='单价（箱）',
                    y='数量（箱）',
                    size='销售额',
                    color='所属区域',
                    hover_name='简化产品名称',  # 使用简化产品名称
                    title='价格与销售数量关系',
                    labels={'单价（箱）': '单价 (元/箱)', '数量（箱）': '销售数量 (箱)'},
                    height=500
                )

                # 添加趋势线
                fig_price_qty.update_layout(
                    xaxis_title=dict(text="单价 (元/箱)", font=dict(size=16)),
                    yaxis_title=dict(text="销售数量 (箱)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                st.plotly_chart(fig_price_qty, use_container_width=True)
            except Exception as e:
                st.error(f"创建价格-销量散点图时出错: {str(e)}")
    except Exception as e:
        st.error(f"创建产品销售分析图表时出错: {str(e)}")

    # 申请人销售业绩
    st.markdown('<div class="sub-header section-gap">👨‍💼 申请人销售业绩</div>', unsafe_allow_html=True)

    try:
        applicant_performance = filtered_df.groupby('申请人')['销售额'].sum().sort_values(ascending=False).reset_index()

        if not applicant_performance.empty:
            fig_applicant = px.bar(
                applicant_performance,
                x='申请人',
                y='销售额',
                color='申请人',
                title='申请人销售业绩排名',
                labels={'销售额': '销售额 (元)', '申请人': '申请人'},
                height=500
            )
            # 添加文本标签
            fig_applicant.update_traces(
                text=[format_yuan(val) for val in applicant_performance['销售额']],
                textposition='outside',
                textfont=dict(size=14)
            )
            fig_applicant.update_layout(
                xaxis_title=dict(text="申请人", font=dict(size=16)),
                yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                xaxis_tickfont=dict(size=14),
                yaxis_tickfont=dict(size=14),
                margin=dict(t=60, b=80, l=80, r=60),
                plot_bgcolor='rgba(0,0,0,0)'
            )
            # 确保Y轴有足够空间显示数据标签
            fig_applicant.update_yaxes(
                range=[0, applicant_performance['销售额'].max() * 1.2]
            )
            st.plotly_chart(fig_applicant, use_container_width=True)
        else:
            st.warning("没有足够的申请人销售数据来创建图表。")
    except Exception as e:
        st.error(f"创建申请人销售业绩图表时出错: {str(e)}")

    # 原始数据表
    with st.expander("查看筛选后的原始数据"):
        st.dataframe(filtered_df)

with tabs[1]:  # 新品分析
    st.markdown('<div class="sub-header">🆕 新品销售分析</div>', unsafe_allow_html=True)

    # 检查新品数据是否为空
    if filtered_new_products_df.empty:
        st.warning("当前筛选条件下没有新品销售数据。请调整筛选条件或确认产品代码是否正确。")
    else:
        # 新品KPI指标
        col1, col2, col3 = st.columns(3)

        try:
            new_products_sales = filtered_new_products_df['销售额'].sum()
            with col1:
                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">新品销售额</div>
                    <div class="metric-value">{format_yuan(new_products_sales)}</div>
                </div>
                """, unsafe_allow_html=True)

            new_products_percentage = (new_products_sales / total_sales * 100) if total_sales > 0 else 0
            with col2:
                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">新品销售占比</div>
                    <div class="metric-value">{new_products_percentage:.2f}%</div>
                </div>
                """, unsafe_allow_html=True)

            new_products_customers = filtered_new_products_df['客户简称'].nunique()
            with col3:
                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">购买新品的客户数</div>
                    <div class="metric-value">{new_products_customers}</div>
                </div>
                """, unsafe_allow_html=True)
        except Exception as e:
            st.error(f"计算新品KPI指标时出错: {str(e)}")

        # 新品销售详情
        st.markdown('<div class="sub-header section-gap">各新品销售额对比</div>', unsafe_allow_html=True)

        try:
            # 使用简化产品名称
            product_sales = filtered_new_products_df.groupby(['产品代码', '简化产品名称'])['销售额'].sum().reset_index()
            product_sales = product_sales.sort_values('销售额', ascending=False)

            if not product_sales.empty:
                fig_product_sales = px.bar(
                    product_sales,
                    x='简化产品名称',  # 使用简化产品名称
                    y='销售额',
                    color='简化产品名称',  # 使用简化产品名称
                    title='新品产品销售额对比',
                    labels={'销售额': '销售额 (元)', '简化产品名称': '产品名称'},
                    height=500
                )
                # 添加文本标签
                fig_product_sales.update_traces(
                    text=[format_yuan(val) for val in product_sales['销售额']],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_product_sales.update_layout(
                    xaxis_title=dict(text="产品名称", font=dict(size=16)),
                    yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_product_sales.update_yaxes(
                    range=[0, product_sales['销售额'].max() * 1.2]
                )
                st.plotly_chart(fig_product_sales, use_container_width=True)
            else:
                st.warning("没有足够的新品销售数据来创建图表。")
        except Exception as e:
            st.error(f"创建新品销售对比图表时出错: {str(e)}")

        # 区域新品销售分析
        st.markdown('<div class="sub-header section-gap">区域新品销售分析</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)

        try:
            with col1:
                # 区域新品销售额堆叠柱状图
                region_product_sales = filtered_new_products_df.groupby(['所属区域', '简化产品名称'])[
                    '销售额'].sum().reset_index()

                if not region_product_sales.empty:
                    fig_region_product = px.bar(
                        region_product_sales,
                        x='所属区域',
                        y='销售额',
                        color='简化产品名称',  # 使用简化产品名称
                        title='各区域新品销售额分布',
                        labels={'销售额': '销售额 (元)', '所属区域': '区域', '简化产品名称': '产品名称'},
                        height=500
                    )
                    fig_region_product.update_layout(
                        xaxis_title=dict(text="区域", font=dict(size=16)),
                        yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                        xaxis_tickfont=dict(size=14),
                        yaxis_tickfont=dict(size=14),
                        margin=dict(t=60, b=80, l=80, r=60),
                        plot_bgcolor='rgba(0,0,0,0)',
                        legend_title="产品名称",
                        legend_font=dict(size=12)
                    )
                    st.plotly_chart(fig_region_product, use_container_width=True)
                else:
                    st.warning("没有足够的区域新品销售数据来创建图表。")

            with col2:
                # 新品占比饼图
                fig_new_vs_old = px.pie(
                    values=[new_products_sales, total_sales - new_products_sales],
                    names=['新品', '非新品'],
                    title='新品销售额占总销售额比例',
                    hole=0.4,
                    color_discrete_sequence=['#ff9999', '#66b3ff']
                )
                fig_new_vs_old.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    textfont=dict(size=14)
                )
                fig_new_vs_old.update_layout(
                    margin=dict(t=60, b=60, l=60, r=60),
                    font=dict(size=14)
                )
                st.plotly_chart(fig_new_vs_old, use_container_width=True)
        except Exception as e:
            st.error(f"创建区域新品销售分析图表时出错: {str(e)}")

        # 区域内新品销售占比热力图
        st.markdown('<div class="sub-header section-gap">各区域内新品销售占比</div>', unsafe_allow_html=True)

        try:
            # 计算各区域的新品总销售额
            region_total_sales = filtered_new_products_df.groupby('所属区域')['销售额'].sum().reset_index()

            # 计算各区域各新品的销售占比
            region_product_sales = filtered_new_products_df.groupby(['所属区域', '产品代码', '简化产品名称'])[
                '销售额'].sum().reset_index()

            if not region_total_sales.empty and not region_product_sales.empty:
                region_product_sales = region_product_sales.merge(region_total_sales, on='所属区域',
                                                                  suffixes=('', '_区域总计'))
                region_product_sales['销售占比'] = region_product_sales['销售额'] / region_product_sales[
                    '销售额_区域总计'] * 100

                # 创建显示名称列（简化产品名称）
                region_product_sales['显示名称'] = region_product_sales['简化产品名称']

                # 透视表
                pivot_percentage = pd.pivot_table(
                    region_product_sales,
                    values='销售占比',
                    index='所属区域',
                    columns='显示名称',  # 使用简化名称作为列名
                    fill_value=0
                )

                # 使用Plotly创建热力图
                fig_heatmap = px.imshow(
                    pivot_percentage,
                    labels=dict(x="产品名称", y="区域", color="销售占比 (%)"),
                    x=pivot_percentage.columns,
                    y=pivot_percentage.index,
                    color_continuous_scale="YlGnBu",
                    title="各区域内新品销售占比 (%)",
                    height=500
                )

                fig_heatmap.update_layout(
                    xaxis_title=dict(text="产品名称", font=dict(size=16)),
                    yaxis_title=dict(text="区域", font=dict(size=16)),
                    margin=dict(t=80, b=80, l=100, r=100),
                    font=dict(size=14)
                )

                # 添加注释
                for i in range(len(pivot_percentage.index)):
                    for j in range(len(pivot_percentage.columns)):
                        fig_heatmap.add_annotation(
                            x=j,
                            y=i,
                            text=f"{pivot_percentage.iloc[i, j]:.1f}%",
                            showarrow=False,
                            font=dict(color="black" if pivot_percentage.iloc[i, j] < 50 else "white", size=14)
                        )

                st.plotly_chart(fig_heatmap, use_container_width=True)
            else:
                st.warning("没有足够的区域内新品销售数据来创建热力图。")
        except Exception as e:
            st.error(f"创建区域内新品销售占比热力图时出错: {str(e)}")

        # 新品数据表
        with st.expander("查看新品销售数据"):
            display_columns = [col for col in filtered_new_products_df.columns if
                               col != '产品代码' or col != '产品名称']
            st.dataframe(filtered_new_products_df[display_columns])

with tabs[2]:  # 客户细分
    st.markdown('<div class="sub-header">👥 客户细分分析</div>', unsafe_allow_html=True)

    try:
        # 检查是否有足够的数据进行分析
        if filtered_df.empty:
            st.warning("没有数据可供分析。请调整筛选条件。")
        else:
            # 计算客户特征
            customer_features = filtered_df.groupby('客户简称').agg({
                '销售额': 'sum',  # 总销售额
                '产品代码': lambda x: len(set(x)),  # 购买的不同产品数量
                '数量（箱）': 'sum',  # 总购买数量
                '单价（箱）': 'mean'  # 平均单价
            }).reset_index()

            # 确保new_products_df不为空
            if not filtered_new_products_df.empty:
                # 添加新品购买指标
                new_products_by_customer = filtered_new_products_df.groupby('客户简称')['销售额'].sum().reset_index()
                customer_features = customer_features.merge(new_products_by_customer, on='客户简称', how='left',
                                                            suffixes=('', '_新品'))
                customer_features['销售额_新品'] = customer_features['销售额_新品'].fillna(0)
                customer_features['新品占比'] = customer_features['销售额_新品'] / customer_features['销售额'] * 100
            else:
                # 如果没有新品数据，添加默认值
                st.info("当前筛选条件下没有新品销售数据。将使用默认值0进行分析。")
                customer_features['销售额_新品'] = 0
                customer_features['新品占比'] = 0

            # 简单客户分类
            customer_features['客户类型'] = pd.cut(
                customer_features['新品占比'],
                bins=[0, 10, 30, 100],
                labels=['保守型客户', '平衡型客户', '创新型客户']
            )

            # 客户分类展示
            st.markdown('<div class="sub-header section-gap">客户类型分布</div>', unsafe_allow_html=True)

            simple_segments = customer_features.groupby('客户类型').agg({
                '客户简称': 'count',
                '销售额': 'mean',
                '新品占比': 'mean'
            }).reset_index()

            simple_segments.columns = ['客户类型', '客户数量', '平均销售额', '平均新品占比']

            # 确保有数据再创建图表
            if not simple_segments.empty:
                # 创建图表代码...
                # 使用Plotly绘制客户类型分布
                fig_customer_types = px.bar(
                    simple_segments,
                    x='客户类型',
                    y='客户数量',
                    color='客户类型',
                    title='客户类型分布',
                    text='客户数量',
                    height=500
                )

                fig_customer_types.update_traces(
                    texttemplate='%{text}',
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_customer_types.update_layout(
                    xaxis_title=dict(text="客户类型", font=dict(size=16)),
                    yaxis_title=dict(text="客户数量", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_customer_types.update_yaxes(
                    range=[0, simple_segments['客户数量'].max() * 1.2]
                )

                st.plotly_chart(fig_customer_types, use_container_width=True)

                # 客户类型特征对比
                st.markdown('<div class="sub-header section-gap">不同客户类型的特征对比</div>', unsafe_allow_html=True)

                # 创建子图 - 优化版
                fig = make_subplots(rows=1, cols=2,
                                    subplot_titles=("客户类型平均销售额", "客户类型平均新品占比"),
                                    specs=[[{"type": "bar"}, {"type": "bar"}]])

                # 添加平均销售额柱状图
                fig.add_trace(
                    go.Bar(
                        x=simple_segments['客户类型'],
                        y=simple_segments['平均销售额'],
                        name='平均销售额',
                        marker_color='rgb(55, 83, 109)',
                        text=[format_yuan(val) for val in simple_segments['平均销售额']],  # 添加文本标签
                        textposition='outside',  # 标签位置设为外部
                        textfont=dict(size=14)
                    ),
                    row=1, col=1
                )

                # 添加平均新品占比柱状图
                fig.add_trace(
                    go.Bar(
                        x=simple_segments['客户类型'],
                        y=simple_segments['平均新品占比'],
                        name='平均新品占比',
                        marker_color='rgb(26, 118, 255)',
                        text=[f"{x:.1f}%" for x in simple_segments['平均新品占比']],  # 添加文本标签
                        textposition='outside',  # 标签位置设为外部
                        textfont=dict(size=14)
                    ),
                    row=1, col=2
                )

                # 优化图表布局
                fig.update_layout(
                    height=500,  # 增加高度
                    showlegend=False,
                    margin=dict(t=80, b=80, l=80, r=80),  # 增加边距
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(
                        family="Arial, sans-serif",
                        size=14,  # 增加字体大小
                        color="rgb(50, 50, 50)"
                    ),
                    title_font=dict(size=18)  # 标题字体大小
                )

                # 优化X轴和Y轴
                fig.update_xaxes(
                    title_text="客户类型",
                    title_font=dict(size=16),
                    tickfont=dict(size=14),
                    row=1, col=1
                )

                fig.update_yaxes(
                    title_text="平均销售额 (元)",
                    title_font=dict(size=16),
                    tickfont=dict(size=14),
                    tickformat=",",  # 添加千位分隔符
                    row=1, col=1
                )

                fig.update_xaxes(
                    title_text="客户类型",
                    title_font=dict(size=16),
                    tickfont=dict(size=14),
                    row=1, col=2
                )

                fig.update_yaxes(
                    title_text="平均新品占比 (%)",
                    title_font=dict(size=16),
                    tickfont=dict(size=14),
                    row=1, col=2
                )

                # 确保Y轴有足够空间显示数据标签
                fig.update_yaxes(range=[0, simple_segments['平均销售额'].max() * 1.3], row=1, col=1)
                fig.update_yaxes(range=[0, simple_segments['平均新品占比'].max() * 1.3], row=1, col=2)

                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("无法创建客户类型分布图：分类后的数据为空。")

            # 散点图检查必要的列是否存在
            if not customer_features.empty and '新品占比' in customer_features.columns and '销售额' in customer_features.columns:
                # 客户销售额和新品占比散点图
                st.markdown('<div class="sub-header section-gap">客户销售额与新品占比关系</div>',
                            unsafe_allow_html=True)

                fig_scatter = px.scatter(
                    customer_features,
                    x='销售额',
                    y='新品占比',
                    color='客户类型',
                    size='产品代码',  # 购买的产品种类数量
                    hover_name='客户简称',
                    title='客户销售额与新品占比关系',
                    labels={
                        '销售额': '销售额 (元)',
                        '新品占比': '新品销售占比 (%)',
                        '产品代码': '购买产品种类数'
                    },
                    height=500
                )

                fig_scatter.update_layout(
                    xaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                    yaxis_title=dict(text="新品销售占比 (%)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)',
                    legend_font=dict(size=14)
                )

                st.plotly_chart(fig_scatter, use_container_width=True)

                # 新品接受度最高的客户
                st.markdown('<div class="sub-header section-gap">新品接受度最高的客户</div>', unsafe_allow_html=True)

                top_acceptance = customer_features.sort_values('新品占比', ascending=False).head(10)

                if not top_acceptance.empty:
                    fig_top_acceptance = px.bar(
                        top_acceptance,
                        x='客户简称',
                        y='新品占比',
                        color='新品占比',
                        title='新品接受度最高的前10名客户',
                        labels={'新品占比': '新品销售占比 (%)', '客户简称': '客户'},
                        height=500,
                        color_continuous_scale=px.colors.sequential.Viridis
                    )
                    # 添加文本标签
                    fig_top_acceptance.update_traces(
                        text=[f"{x:.1f}%" for x in top_acceptance['新品占比']],
                        textposition='outside',
                        textfont=dict(size=14)
                    )
                    fig_top_acceptance.update_layout(
                        xaxis_title=dict(text="客户", font=dict(size=16)),
                        yaxis_title=dict(text="新品销售占比 (%)", font=dict(size=16)),
                        xaxis_tickfont=dict(size=14),
                        yaxis_tickfont=dict(size=14),
                        margin=dict(t=60, b=80, l=80, r=60),
                        plot_bgcolor='rgba(0,0,0,0)'
                    )
                    # 确保Y轴有足够空间显示数据标签
                    fig_top_acceptance.update_yaxes(
                        range=[0, top_acceptance['新品占比'].max() * 1.2]
                    )

                    st.plotly_chart(fig_top_acceptance, use_container_width=True)
                else:
                    st.warning("没有足够的数据来显示新品接受度最高的客户。")
            else:
                st.warning("无法创建散点图：数据不足或缺少必要列。")

            # 客户表格
            with st.expander("查看客户细分数据"):
                st.dataframe(customer_features)
    except Exception as e:
        st.error(f"客户细分分析出错: {str(e)}")
        st.info("请尝试调整筛选条件或检查数据格式。")

with tabs[3]:  # 产品组合
    st.markdown('<div class="sub-header">🔄 产品组合分析</div>', unsafe_allow_html=True)

    try:
        # 共现矩阵分析
        st.markdown('<div class="sub-header section-gap">产品共现矩阵分析</div>', unsafe_allow_html=True)
        st.info("共现矩阵显示不同产品一起被同一客户购买的频率，有助于发现产品间的关联。")

        # 准备数据 - 创建交易矩阵
        if not filtered_df.empty and filtered_df['客户简称'].nunique() > 1 and filtered_df['产品代码'].nunique() > 1:
            transaction_data = filtered_df.groupby(['客户简称', '产品代码'])['销售额'].sum().unstack().fillna(0)
            # 转换为二进制格式（是否购买）
            transaction_binary = transaction_data.applymap(lambda x: 1 if x > 0 else 0)

            # 创建产品共现矩阵
            co_occurrence = pd.DataFrame(0, index=transaction_binary.columns, columns=transaction_binary.columns)

            # 创建产品代码到简化名称的映射
            name_mapping = {code: df[df['产品代码'] == code]['简化产品名称'].iloc[0]
            if len(df[df['产品代码'] == code]) > 0 else code
                            for code in transaction_binary.columns}

            # 计算共现次数
            for _, row in transaction_binary.iterrows():
                bought_products = row.index[row == 1].tolist()
                for p1 in bought_products:
                    for p2 in bought_products:
                        if p1 != p2:
                            co_occurrence.loc[p1, p2] += 1

            # 筛选新品的共现情况
            new_product_co_occurrence = pd.DataFrame()
            valid_new_products = [p for p in new_products if p in co_occurrence.index]

            if valid_new_products:
                for np_code in valid_new_products:
                    top_co = co_occurrence.loc[np_code].sort_values(ascending=False).head(5)
                    new_product_co_occurrence[np_code] = top_co

                # 可视化每个新品的前5个共现产品
                for np_code in valid_new_products:
                    np_name = name_mapping.get(np_code, np_code)  # 获取新品的简化名称
                    st.markdown(f'<div class="sub-header">与"{np_name}"共同购买最多的产品</div>',
                                unsafe_allow_html=True)

                    co_data = co_occurrence.loc[np_code].sort_values(ascending=False).head(5).reset_index()
                    co_data.columns = ['产品代码', '共现次数']

                    # 添加简化产品名称
                    co_data['简化产品名称'] = co_data['产品代码'].map(name_mapping)

                    if not co_data.empty and co_data['共现次数'].max() > 0:
                        fig_co = px.bar(
                            co_data,
                            x='简化产品名称',  # 使用简化产品名称
                            y='共现次数',
                            color='简化产品名称',
                            title=f'与{np_name}共同购买最多的产品',
                            labels={'共现次数': '共同购买次数', '简化产品名称': '产品名称'},
                            height=500
                        )
                        # 添加文本标签
                        fig_co.update_traces(
                            text=co_data['共现次数'],
                            textposition='outside',
                            textfont=dict(size=14)
                        )
                        fig_co.update_layout(
                            xaxis_title=dict(text="产品名称", font=dict(size=16)),
                            yaxis_title=dict(text="共同购买次数", font=dict(size=16)),
                            xaxis_tickfont=dict(size=14),
                            yaxis_tickfont=dict(size=14),
                            margin=dict(t=60, b=80, l=80, r=60),
                            plot_bgcolor='rgba(0,0,0,0)'
                        )
                        # 确保Y轴有足够空间显示数据标签
                        fig_co.update_yaxes(
                            range=[0, co_data['共现次数'].max() * 1.2]
                        )

                        st.plotly_chart(fig_co, use_container_width=True)
                    else:
                        st.info(f"没有与{np_name}共同购买的产品记录。")

                # 热力图展示所有产品的共现关系
                st.markdown('<div class="sub-header section-gap">产品共现热力图</div>', unsafe_allow_html=True)
                st.info("热力图显示产品之间的共现关系，颜色越深表示两个产品一起购买的频率越高。")

                # 筛选主要产品以避免图表过于复杂
                top_products = filtered_df.groupby('产品代码')['销售额'].sum().sort_values(ascending=False).head(
                    10).index.tolist()
                # 确保所有新品都包含在内
                for np in valid_new_products:
                    if np not in top_products:
                        top_products.append(np)

                # 创建简化名称映射的列表
                top_product_names = [name_mapping.get(code, code) for code in top_products]

                # 创建热力图数据
                heatmap_data = co_occurrence.loc[top_products, top_products].copy()

                # 创建热力图
                fig_co_heatmap = px.imshow(
                    heatmap_data,
                    labels=dict(x="产品名称", y="产品名称", color="共现次数"),
                    x=top_product_names,  # 使用简化名称
                    y=top_product_names,  # 使用简化名称
                    color_continuous_scale="Viridis",
                    title="产品共现热力图",
                    height=600  # 增加高度以容纳更多数据
                )

                fig_co_heatmap.update_layout(
                    margin=dict(t=80, b=80, l=100, r=100),
                    font=dict(size=14),
                    xaxis_tickangle=-45  # 倾斜x轴标签以防重叠
                )

                # 添加数值注释
                for i in range(len(top_products)):
                    for j in range(len(top_products)):
                        if heatmap_data.iloc[i, j] > 0:  # 只显示非零值
                            fig_co_heatmap.add_annotation(
                                x=j,
                                y=i,
                                text=str(heatmap_data.iloc[i, j]),
                                showarrow=False,
                                font=dict(color="white" if heatmap_data.iloc[
                                                               i, j] > heatmap_data.max().max() / 2 else "black",
                                          size=12)
                            )

                st.plotly_chart(fig_co_heatmap, use_container_width=True)
            else:
                st.warning("在当前筛选条件下，未找到新品数据或共现关系。")

            # 产品购买模式
            st.markdown('<div class="sub-header section-gap">产品购买模式分析</div>', unsafe_allow_html=True)

            # 计算平均每单购买的产品种类数
            avg_products_per_order = transaction_binary.sum(axis=1).mean()

            col1, col2 = st.columns(2)

            with col1:
                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">平均每客户购买产品种类</div>
                    <div class="metric-value">{avg_products_per_order:.2f}</div>
                </div>
                """, unsafe_allow_html=True)

            with col2:
                # 计算含有新品的订单比例
                orders_with_new_products = transaction_binary[valid_new_products].any(
                    axis=1).sum() if valid_new_products else 0
                total_orders = len(transaction_binary)
                percentage_orders_with_new = (orders_with_new_products / total_orders * 100) if total_orders > 0 else 0

                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">含新品的客户比例</div>
                    <div class="metric-value">{percentage_orders_with_new:.2f}%</div>
                </div>
                """, unsafe_allow_html=True)

            # 购买产品种类数分布
            products_per_order = transaction_binary.sum(axis=1).value_counts().sort_index().reset_index()
            products_per_order.columns = ['产品种类数', '客户数']

            if not products_per_order.empty:
                fig_products_dist = px.bar(
                    products_per_order,
                    x='产品种类数',
                    y='客户数',
                    title='客户购买产品种类数分布',
                    labels={'产品种类数': '购买产品种类数', '客户数': '客户数量'},
                    height=500
                )
                # 添加文本标签
                fig_products_dist.update_traces(
                    text=products_per_order['客户数'],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_products_dist.update_layout(
                    xaxis_title=dict(text="购买产品种类数", font=dict(size=16)),
                    yaxis_title=dict(text="客户数量", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_products_dist.update_yaxes(
                    range=[0, products_per_order['客户数'].max() * 1.2]
                )

                st.plotly_chart(fig_products_dist, use_container_width=True)
            else:
                st.warning("没有足够的数据来显示客户购买产品种类数分布。")

            # 产品组合表格
            with st.expander("查看产品共现矩阵"):
                # 转换产品代码为简化名称
                display_co_occurrence = co_occurrence.copy()
                display_co_occurrence.index = [name_mapping.get(code, code) for code in display_co_occurrence.index]
                display_co_occurrence.columns = [name_mapping.get(code, code) for code in display_co_occurrence.columns]
                st.dataframe(display_co_occurrence)
        else:
            st.warning("当前筛选条件下的数据不足以进行产品组合分析。需要多个客户和多个产品。")
    except Exception as e:
        st.error(f"产品组合分析出错: {str(e)}")
        st.info("请尝试调整筛选条件或检查数据格式。")

with tabs[4]:  # 市场渗透率
    st.markdown('<div class="sub-header">🌐 新品市场渗透率分析</div>', unsafe_allow_html=True)

    try:
        # 计算总体渗透率
        total_customers = filtered_df['客户简称'].nunique()
        new_product_customers = filtered_new_products_df['客户简称'].nunique()
        penetration_rate = (new_product_customers / total_customers * 100) if total_customers > 0 else 0

        # KPI指标
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">总客户数</div>
                <div class="metric-value">{total_customers}</div>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">购买新品的客户数</div>
                <div class="metric-value">{new_product_customers}</div>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">新品市场渗透率</div>
                <div class="metric-value">{penetration_rate:.2f}%</div>
            </div>
            """, unsafe_allow_html=True)

        # 区域渗透率分析
        st.markdown('<div class="sub-header section-gap">各区域新品渗透率</div>', unsafe_allow_html=True)

        if 'selected_regions' in locals() and selected_regions:
            # 按区域计算渗透率
            region_customers = filtered_df.groupby('所属区域')['客户简称'].nunique().reset_index()
            region_customers.columns = ['所属区域', '客户总数']

            new_region_customers = filtered_new_products_df.groupby('所属区域')['客户简称'].nunique().reset_index()
            new_region_customers.columns = ['所属区域', '购买新品客户数']

            region_penetration = region_customers.merge(new_region_customers, on='所属区域', how='left')
            region_penetration['购买新品客户数'] = region_penetration['购买新品客户数'].fillna(0)
            region_penetration['渗透率'] = (
                    region_penetration['购买新品客户数'] / region_penetration['客户总数'] * 100).round(2)

            if not region_penetration.empty:
                # 创建区域渗透率条形图
                fig_region_penetration = px.bar(
                    region_penetration,
                    x='所属区域',
                    y='渗透率',
                    color='所属区域',
                    text='渗透率',
                    title='各区域新品市场渗透率',
                    labels={'渗透率': '渗透率 (%)', '所属区域': '区域'},
                    height=500
                )

                fig_region_penetration.update_traces(
                    texttemplate='%{text:.2f}%',
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_region_penetration.update_layout(
                    xaxis_title=dict(text="区域", font=dict(size=16)),
                    yaxis_title=dict(text="渗透率 (%)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_region_penetration.update_yaxes(
                    range=[0, region_penetration['渗透率'].max() * 1.2]
                )

                st.plotly_chart(fig_region_penetration, use_container_width=True)

                # 区域渗透率表格
                st.markdown('<div class="sub-header section-gap">区域渗透率详细数据</div>', unsafe_allow_html=True)
                st.dataframe(region_penetration)

                # 渗透率和销售额关系
                st.markdown('<div class="sub-header section-gap">渗透率与销售额的关系</div>', unsafe_allow_html=True)

                # 计算每个区域的新品销售额
                region_new_sales = filtered_new_products_df.groupby('所属区域')['销售额'].sum().reset_index()
                region_new_sales.columns = ['所属区域', '新品销售额']

                # 合并渗透率和销售额数据
                region_analysis = region_penetration.merge(region_new_sales, on='所属区域', how='left')
                region_analysis['新品销售额'] = region_analysis['新品销售额'].fillna(0)

                # 创建气泡图
                fig_bubble = px.scatter(
                    region_analysis,
                    x='渗透率',
                    y='新品销售额',
                    size='客户总数',
                    color='所属区域',
                    hover_name='所属区域',
                    text='所属区域',
                    title='区域渗透率与新品销售额关系',
                    labels={
                        '渗透率': '渗透率 (%)',
                        '新品销售额': '新品销售额 (元)',
                        '客户总数': '客户总数'
                    },
                    height=500
                )

                fig_bubble.update_traces(
                    textposition='top center',
                    marker=dict(sizemode='diameter', sizeref=0.1),
                    textfont=dict(size=14)
                )

                fig_bubble.update_layout(
                    xaxis_title=dict(text="渗透率 (%)", font=dict(size=16)),
                    yaxis_title=dict(text="新品销售额 (元)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )

                st.plotly_chart(fig_bubble, use_container_width=True)
            else:
                st.warning("没有足够的数据来计算区域渗透率。")
        else:
            st.warning("请在侧边栏选择至少一个区域以查看区域渗透率分析。")

        # 渗透率趋势分析（如果有时间数据）
        if '发运月份' in filtered_df.columns:
            st.markdown('<div class="sub-header section-gap">新品渗透率趋势</div>', unsafe_allow_html=True)

            try:
                # 确保发运月份是日期类型
                if not pd.api.types.is_datetime64_dtype(filtered_df['发运月份']):
                    filtered_df['发运月份'] = pd.to_datetime(filtered_df['发运月份'])
                if not pd.api.types.is_datetime64_dtype(filtered_new_products_df['发运月份']):
                    filtered_new_products_df['发运月份'] = pd.to_datetime(filtered_new_products_df['发运月份'])

                # 按月分组
                monthly_customers = filtered_df.groupby(pd.Grouper(key='发运月份', freq='M'))[
                    '客户简称'].nunique().reset_index()
                monthly_customers.columns = ['月份', '客户总数']

                monthly_new_customers = filtered_new_products_df.groupby(pd.Grouper(key='发运月份', freq='M'))[
                    '客户简称'].nunique().reset_index()
                monthly_new_customers.columns = ['月份', '购买新品客户数']

                # 合并月度数据
                monthly_penetration = monthly_customers.merge(monthly_new_customers, on='月份', how='left')
                monthly_penetration['购买新品客户数'] = monthly_penetration['购买新品客户数'].fillna(0)
                monthly_penetration['渗透率'] = (
                        monthly_penetration['购买新品客户数'] / monthly_penetration['客户总数'] * 100).round(2)
                monthly_penetration['月份_str'] = monthly_penetration['月份'].dt.strftime('%Y-%m')

                if not monthly_penetration.empty and len(monthly_penetration) > 1:
                    # 创建趋势线图
                    fig_trend = px.line(
                        monthly_penetration,
                        x='月份',
                        y='渗透率',
                        markers=True,
                        title='新品渗透率月度趋势',
                        labels={'渗透率': '渗透率 (%)', '月份': '月份'},
                        height=500
                    )
                    # 添加数据标签
                    fig_trend.update_traces(
                        text=[f"{x:.1f}%" for x in monthly_penetration['渗透率']],
                        textposition='top center',
                        textfont=dict(size=14)
                    )
                    fig_trend.update_layout(
                        xaxis_title=dict(text="月份", font=dict(size=16)),
                        yaxis_title=dict(text="渗透率 (%)", font=dict(size=16)),
                        xaxis_tickfont=dict(size=14),
                        yaxis_tickfont=dict(size=14),
                        margin=dict(t=60, b=80, l=80, r=60),
                        plot_bgcolor='rgba(0,0,0,0)'
                    )

                    st.plotly_chart(fig_trend, use_container_width=True)
                else:
                    st.warning("没有足够的月度数据来显示渗透率趋势。需要多个月份的数据。")
            except Exception as e:
                st.error(f"处理月份数据进行趋势分析时出错: {str(e)}")
                st.info("请确保发运月份格式正确，应为YYYY-MM格式或标准日期格式。")
    except Exception as e:
        st.error(f"市场渗透率分析出错: {str(e)}")
        st.info("请尝试调整筛选条件或检查数据格式。")

# 底部下载区域
st.markdown("---")
st.markdown('<div class="sub-header">📊 导出分析结果</div>', unsafe_allow_html=True)


# 创建Excel报告
def generate_excel_report(df, new_products_df):
    try:
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')

        # 销售概览表
        df.to_excel(writer, sheet_name='销售数据总览', index=False)

        # 新品分析表
        if not new_products_df.empty:
            new_products_df.to_excel(writer, sheet_name='新品销售数据', index=False)

        # 区域销售汇总
        region_summary = df.groupby('所属区域').agg({
            '销售额': 'sum',
            '客户简称': pd.Series.nunique,
            '产品代码': pd.Series.nunique,
            '数量（箱）': 'sum'
        }).reset_index()
        region_summary.columns = ['区域', '销售额', '客户数', '产品数', '销售数量']
        region_summary.to_excel(writer, sheet_name='区域销售汇总', index=False)

        # 产品销售汇总
        product_summary = df.groupby(['产品代码', '简化产品名称']).agg({
            '销售额': 'sum',
            '客户简称': pd.Series.nunique,
            '数量（箱）': 'sum'
        }).sort_values('销售额', ascending=False).reset_index()
        product_summary.columns = ['产品代码', '产品名称', '销售额', '购买客户数', '销售数量']
        product_summary.to_excel(writer, sheet_name='产品销售汇总', index=False)

        # 保存Excel
        writer.close()

        return output.getvalue()
    except Exception as e:
        st.error(f"生成Excel报告时出错: {str(e)}")
        # 返回一个简单的错误报告
        error_output = BytesIO()
        with pd.ExcelWriter(error_output, engine='xlsxwriter') as writer:
            pd.DataFrame({'错误': [f"生成报告时出错: {str(e)}"]}).to_excel(writer, sheet_name='错误信息', index=False)
        return error_output.getvalue()


# 下载按钮
try:
    excel_report = generate_excel_report(filtered_df, filtered_new_products_df)

    st.markdown('<div class="download-button">', unsafe_allow_html=True)
    st.download_button(
        label="下载Excel分析报告",
        data=excel_report,
        file_name="销售数据分析报告.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown('</div>', unsafe_allow_html=True)
except Exception as e:
    st.error(f"创建下载按钮时出错: {str(e)}")

# 底部注释
st.markdown("""
<div style="text-align: center; margin-top: 30px; color: #666;">
    <p>销售数据分析仪表盘 © 2025</p>
</div>
""", unsafe_allow_html=True)
