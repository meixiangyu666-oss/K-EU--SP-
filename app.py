import streamlit as st
import pandas as pd
from collections import defaultdict
import re
import os

# 设置页面配置
st.set_page_config(page_title="K EU 小赖版-SP广告批量模版工具", page_icon="📊", layout="centered")

# 自定义 CSS 样式
st.markdown("""
    <style>
    /* 主标题样式 */
    .main-title {
        font-size: 2.5em;
        font-weight: bold;
        color: #2C3E50;
        text-align: center;
        margin-bottom: 20px;
    }
    /* 规则说明样式 */
    .rules {
        font-size: 0.9em;
        color: #34495E;
        background-color: #F8F9FA;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    /* 按钮样式 */
    .stButton>button {
        background-color: #3498DB;
        color: white;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 1em;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #2980B9;
    }
    /* 文件上传框样式 */
    .stFileUploader label {
        font-size: 1.1em;
        color: #2C3E50;
        font-weight: bold;
    }
    /* 成功和错误消息样式 */
    .stSuccess {
        background-color: #E8F5E9;
        border-left: 5px solid #4CAF50;
        padding: 10px;
        border-radius: 5px;
    }
    .stError {
        background-color: #FFEBEE;
        border-left: 5px solid #F44336;
        padding: 10px;
        border-radius: 5px;
    }
    .stWarning {
        background-color: #FFF3E0;
        border-left: 5px solid #FF9800;
        padding: 10px;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# 规则说明
st.markdown('<div class="main-title">K EU 小赖版-SP广告批量模版工具</div>', unsafe_allow_html=True)
st.markdown("""
<div class="rules">
<b>使用规则 / Usage Rules:</b><br>
1. 上传任意 .xlsx 文件（文件名不限），需包含 "广告活动名称"、"CPC"、"SKU"、"广告组默认竞价"、"预算" 列。<br>
2. H-Q 列为关键词列（精准/广泛）。<br>
3. 支持精准/广泛/ASIN 活动，自动生成否定关键词和商品定向。<br>
4. 输出文件: header-K EU.xlsx（包含广告活动、广告组、关键词、否定关键词、商品定向等）。<br>
5. 如有重复关键词，生成中止，请清理后重试。<br><br>
<b>Upload any .xlsx file (filename flexible), must include "广告活动名称", "CPC", "SKU", "广告组默认竞价", "预算" columns.</b><br>
<b>H-Q columns for keywords (exact/broad).</b><br>
<b>Supports exact/broad/ASIN campaigns, auto-generates negatives and product targeting.</b><br>
<b>Output: header-K EU.xlsx (includes campaigns, groups, keywords, negatives, product targeting).</b><br>
<b>If duplicate keywords, generation stops; clean and retry.</b>
</div>
""", unsafe_allow_html=True)

# 文件上传
uploaded_file = st.file_uploader("上传 Excel 文件 (任意 .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # 保存上传的文件
    with open("temp_survey.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # 运行按钮
    if st.button("生成 Header 文件"):
        output_file = 'header-K EU.xlsx'
        with st.spinner("正在处理文件..."):
            result = generate_header_from_survey("temp_survey.xlsx", output_file)
            if result and os.path.exists(result):
                with open(result, "rb") as f:
                    st.download_button(
                        label="下载 header-K EU.xlsx",
                        data=f,
                        file_name="header-K EU.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("生成成功！请下载文件。")
            else:
                st.error("生成失败，请检查文件格式。")

# script-K EU.py 的函数
def generate_header_from_survey(survey_file='temp_survey.xlsx', output_file='header-K EU.xlsx', sheet_name=0):
    try:
        # 读取 Excel 文件
        df_survey = pd.read_excel(survey_file, sheet_name=sheet_name)
        st.write(f"成功读取文件，数据形状：{df_survey.shape}")
        st.write(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        st.error(f"错误：未找到文件 {survey_file}。")
        return None
    except Exception as e:
        st.error(f"读取文件时出错：{e}")
        return None
    
    # 提取独特活动名称
    unique_campaigns = [name for name in df_survey['广告活动名称'].dropna() if str(name).strip()]
    st.write(f"独特活动名称数量: {len(unique_campaigns)}: {unique_campaigns}")
    
    # 创建活动到 CPC/SKU/广告组默认竞价/预算 的映射
    non_empty_campaigns = df_survey[
        df_survey['广告活动名称'].notna() & 
        (df_survey['广告活动名称'] != '')
    ]
    required_cols = ['CPC', 'SKU', '广告组默认竞价', '预算']
    if all(col in non_empty_campaigns.columns for col in required_cols):
        campaign_to_values = non_empty_campaigns.drop_duplicates(
            subset='广告活动名称', keep='first'
        ).set_index('广告活动名称')[required_cols].to_dict('index')
    else:
        campaign_to_values = {}
        st.warning(f"警告：缺少列 {set(required_cols) - set(non_empty_campaigns.columns)}，使用默认值")
    
    st.write(f"生成的字典（有 {len(campaign_to_values)} 个活动）: {campaign_to_values}")
    
    # 关键词列：第 H 列（索引 7）到第 Q 列（索引 16）
    keyword_columns = df_survey.columns[7:17]
    st.write(f"关键词列: {list(keyword_columns)}")
    
    # 检查关键词重复
    duplicates_found = False
    st.write("### 检查关键词重复")
    for col in keyword_columns:
        col_index = list(df_survey.columns).index(col) + 1
        col_letter = chr(64 + col_index) if col_index <= 26 else f"{chr(64 + (col_index-1)//26)}{chr(64 + (col_index-1)%26 + 1)}"
        kw_list = [kw for kw in df_survey[col].dropna() if str(kw).strip()]
        
        if len(kw_list) > len(set(kw_list)):
            duplicates_mask = df_survey[col].duplicated(keep=False)
            duplicates_df = df_survey[duplicates_mask][[col]].dropna()
            st.warning(f"警告：{col_letter} 列 ({col}) 有重复关键词")
            for _, row in duplicates_df.iterrows():
                kw = str(row[col]).strip()
                count = (df_survey[col] == kw).sum()
                if count > 1:
                    st.write(f"  重复词: '{kw}' (出现 {count} 次)")
            duplicates_found = True
    
    if duplicates_found:
        st.error("提示：由于检测到关键词重复，本次不生成表格。请清理重复后重试。")
        return None
    
    st.write("关键词无重复，继续生成...")
    
    # 列定义
    columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告组合编号', '广告编号', '关键词编号', '商品投放 ID',
        '广告活动名称', '广告组名称', '开始日期', '结束日期', '投放类型', '状态', '每日预算', 'SKU', '广告组默认竞价',
        '竞价', '关键词文本', '匹配类型', '竞价方案', '广告位', '百分比', '拓展商品投放编号'
    ]
    
    # 默认值
    product = '商品推广'
    operation = 'Create'
    status = '已启用'
    targeting_type = '手动'
    bidding_strategy = '动态竞价 - 仅降低'
    default_daily_budget = 12
    default_group_bid = 0.6
    
    # 从列名中提取所有可能的关键词类别
    keyword_categories = set()
    for col in df_survey.columns:
        col_lower = str(col).lower()
        if 'asin' in col_lower and '否定' not in col_lower:
            # 提取 ASIN 列前缀作为类别
            prefix = col_lower.replace('asin', '').strip()
            parts = re.split(r'[/\-_\s\.]', prefix)
            for part in parts:
                if part and len(part) > 1:
                    keyword_categories.add(part)
        elif any(x in col_lower for x in ['精准词', '广泛词']):
            # 提取关键词列前缀
            for suffix in ['精准词', '广泛词']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    parts = re.split(r'[/\-_\s\.]', prefix)
                    for part in parts:
                        if part and len(part) > 1:
                            keyword_categories.add(part)
                    break
    
    st.write(f"识别到的关键词类别: {keyword_categories}")
    
    # 生成数据行
    rows = []
    
    # 函数：查找匹配的关键词列
    def find_matching_keyword_columns(campaign_name, df_survey, keyword_categories, keyword_columns, match_type):
        campaign_name_normalized = str(campaign_name).lower()
        
        # 确定关键词类别
        matched_categories = []
        for category in keyword_categories:
            if category and category in campaign_name_normalized:
                matched_categories.append(category)
        
        st.write(f"  匹配的关键词类别