import streamlit as st
import pandas as pd
from collections import defaultdict
import re
import io
import uuid

# 函数：从调研 Excel 生成表头 Excel
def generate_header_from_survey(df_survey, output_file='header-K EU.xlsx'):
    print(f"成功读取数据，数据形状：{df_survey.shape}")
    print(f"列名列表: {list(df_survey.columns)}")
    
    # 提取独特活动名称
    unique_campaigns = [name for name in df_survey['广告活动名称'].dropna() if str(name).strip()]
    print(f"独特活动名称数量: {len(unique_campaigns)}: {unique_campaigns}")
    
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
        print(f"警告：缺少列 {set(required_cols) - set(non_empty_campaigns.columns)}，使用默认值")
    
    print(f"生成的字典（有 {len(campaign_to_values)} 个活动）: {campaign_to_values}")
    
    # 关键词列：第 H 列（索引 7）到第 Q 列（索引 16）
    keyword_columns = df_survey.columns[7:17]
    print(f"关键词列: {list(keyword_columns)}")
    
    # 检查关键词重复
    duplicates_found = False
    print("\n=== 检查关键词重复 ===")
    for col in keyword_columns:
        col_index = list(df_survey.columns).index(col) + 1
        col_letter = chr(64 + col_index) if col_index <= 26 else f"{chr(64 + (col_index-1)//26)}{chr(64 + (col_index-1)%26 + 1)}"
        kw_list = [kw for kw in df_survey[col].dropna() if str(kw).strip()]
        
        if len(kw_list) > len(set(kw_list)):
            duplicates_mask = df_survey[col].duplicated(keep=False)
            duplicates_df = df_survey[duplicates_mask][[col]].dropna()
            print(f"警告：{col_letter} 列 ({col}) 有重复关键词")
            for _, row in duplicates_df.iterrows():
                kw = str(row[col]).strip()
                count = (df_survey[col] == kw).sum()
                if count > 1:
                    print(f"  重复词: '{kw}' (出现 {count} 次)")
            duplicates_found = True
    
    if duplicates_found:
        st.error("检测到关键词重复，请清理重复关键词后重试。")
        return None
    
    print("关键词无重复，继续生成...")
    
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
            # 从列名中提取关键词类别
            for suffix in ['精准词', '广泛词', '精准', '广泛', 'asin']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    if prefix:
                        if '/' in prefix:
                            parts = [p.strip() for p in prefix.split('/') if p.strip()]
                            keyword_categories.update(parts)
                        else:
                            keyword_categories.add(prefix)
                    break
    # 添加已知的关键词类别（添加 'host' 以匹配活动名）
    keyword_categories.update(['suzhu', 'host', '宿主', 'case', '包', 'tape'])
    # 过滤空字符串
    keyword_categories = {cat for cat in keyword_categories if cat.strip()}
    
    print(f"识别到的关键词类别: {keyword_categories}")
    
    # 生成数据行
    rows = []
    
    for campaign_name in unique_campaigns:
        # 获取 CPC、SKU、广告组默认竞价、预算
        if campaign_name in campaign_to_values:
            cpc = campaign_to_values[campaign_name]['CPC']
            sku = campaign_to_values[campaign_name]['SKU']
            group_bid = campaign_to_values[campaign_name]['广告组默认竞价']
            budget = campaign_to_values[campaign_name]['预算']
        else:
            cpc = 0.5
            sku = 'SKU-1'
            group_bid = default_group_bid
            budget = default_daily_budget
        
        print(f"处理活动: {campaign_name}")
        
        campaign_name_normalized = str(campaign_name).lower()
        
        # 排序类别，按长度升序（优先短类别如 'host'）
        sorted_categories = sorted(keyword_categories, key=len)
        
        # 动态提取关键词类别（现在过滤空串，添加 host）
        matched_category = None
        for category in sorted_categories:
            if category in campaign_name_normalized:
                matched_category = category
                break
        
        print(f"  匹配的关键词类别: {matched_category}")
        
        # 确定匹配类型
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        is_asin = 'asin' in campaign_name_normalized
        match_type = '精准' if is_exact else '广泛' if is_broad else 'ASIN' if is_asin else None
        print(f"  is_exact: {is_exact}, is_broad: {is_broad}, is_asin: {is_asin}, match_type: {match_type}")
        
        # 提取关键词（用于正向关键词，精准/广泛匹配）
        keywords = []
        matched_columns = []
        if matched_category and (is_exact or is_broad):
            for col in keyword_columns:
                col_lower = str(col).lower()
                if is_exact and matched_category in col_lower and any(x in col_lower for x in ['精准', 'exact']):
                    matched_columns.append(col)
                    keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
                elif is_broad and matched_category in col_lower and any(x in col_lower for x in ['广泛', 'broad']):
                    matched_columns.append(col)
                    keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
            keywords = list(dict.fromkeys(keywords))
            print(f"  匹配的列: {matched_columns}")
            print(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
        else:
            print("  无匹配的关键词列，关键词为空")
        
        # 提取 ASIN（用于商品定向）
        asin_targets = []
        if is_asin:
            # 精确匹配：列名必须与广告活动名称完全一致
            if campaign_name in df_survey.columns:
                asin_targets.extend([asin for asin in df_survey[campaign_name].dropna() if str(asin).strip()])
                print(f"  找到与活动名称完全匹配的列: {campaign_name}")
            else:
                print(f"  未找到与活动名称完全匹配的列: {campaign_name}")
            asin_targets = list(dict.fromkeys(asin_targets))
            print(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
        
        # 新增：竞价调整行（每条活动一条）
        placement_value = "广告位：商品页面" if is_asin else "广告位：搜索结果首页首位"
        bid_adjustment_row = [
            product, '竞价调整', operation, campaign_name, '', '', '', '', '',
            campaign_name, campaign_name, '', '', targeting_type, '', '', '', '',
            '', '', '', bidding_strategy, placement_value, '900', ''
        ]
        rows.append(bid_adjustment_row)
        print(f"  添加竞价调整行: 广告位={placement_value}")
        
        # 广告活动行
        rows.append([
            product, '广告活动', operation, campaign_name, '', '', '', '', '',
            campaign_name, '', '', '', targeting_type, status, budget, '', '',
            '', '', '', bidding_strategy, '', '', ''
        ])
        
        # 广告组行
        rows.append([
            product, '广告组', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', '', group_bid,
            '', '', '', '', '', '', ''
        ])
        
        # 商品广告行
        rows.append([
            product, '商品广告', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', sku, '',
            '', '', '', '', '', '', ''
        ])
        
        # 关键词行（仅精准/广泛匹配）
        if is_exact or is_broad:
            for kw in keywords:
                rows.append([
                    product, '关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, kw, match_type, '', '', '', ''
                ])
        
        # 修改后的否定关键词规则
        neg_exact = []
        neg_phrase = []
        
        if is_broad:
            # 广泛组：S列（否定精准），T列（否定词组）
            s_col = df_survey.iloc[:, 18]  # S列 (index 18)
            t_col = df_survey.iloc[:, 19]  # T列 (index 19)
            neg_exact = [kw for kw in s_col.dropna() if str(kw).strip()]
            neg_phrase = [kw for kw in t_col.dropna() if str(kw).strip()]
            neg_exact = list(dict.fromkeys(neg_exact))
            neg_phrase = list(dict.fromkeys(neg_phrase))
            print(f"  广泛组否定：精准 {len(neg_exact)} 个，词组 {len(neg_phrase)} 个")
        elif is_exact and matched_category:
            # 精准组：根据类别
            if matched_category in ['suzhu', 'host', '宿主']:
                # 宿主精准组：U列（否定精准），V列（否定词组）
                u_col = df_survey.iloc[:, 20]  # U列 (index 20)
                v_col = df_survey.iloc[:, 21]  # V列 (index 21)
                neg_exact = [kw for kw in u_col.dropna() if str(kw).strip()]
                neg_phrase = [kw for kw in v_col.dropna() if str(kw).strip()]
            elif matched_category == 'case':
                # case精准组：W列（否定精准），X列（否定词组）
                w_col = df_survey.iloc[:, 22]  # W列 (index 22)
                x_col = df_survey.iloc[:, 23]  # X列 (index 23)
                neg_exact = [kw for kw in w_col.dropna() if str(kw).strip()]
                neg_phrase = [kw for kw in x_col.dropna() if str(kw).strip()]
            neg_exact = list(dict.fromkeys(neg_exact))
            neg_phrase = list(dict.fromkeys(neg_phrase))
            print(f"  {matched_category}精准组否定：精准 {len(neg_exact)} 个，词组 {len(neg_phrase)} 个")
        
        # 否定关键词行（仅精准/广泛匹配）
        if is_exact or is_broad:
            for kw in neg_exact:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
            for kw in neg_phrase:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定词组', '', '', '', ''
                ])
        
        # 商品定向和否定商品定向（仅 ASIN 组）
        neg_asin = [kw for kw in df_survey.get('否定ASIN', pd.Series()).dropna() if str(kw).strip()]
        if is_asin:
            for asin in asin_targets:
                rows.append([
                    product, '商品定向', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, '', '', '', '', '', f'asin="{asin}"'
                ])
            for asin in neg_asin:
                rows.append([
                    product, '否定商品定向', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    '', '', '', '', '', '', f'asin="{asin}"'
                ])
    
    # 创建 DataFrame
    df_header = pd.DataFrame(rows, columns=columns)
    
    # 使用 BytesIO 保存到内存
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df_header.to_excel(writer, index=False, sheet_name='Sheet1')
    output_buffer.seek(0)
    
    print(f"生成完成！总行数：{len(rows)}")
    
    # 调试输出
    keyword_rows = [row for row in rows if row[1] == '关键词']
    print(f"关键词行数量: {len(keyword_rows)}")
    if keyword_rows:
        print(f"示例关键词行: 实体层级={keyword_rows[0][1]}, 关键词文本={keyword_rows[0][19]}, 匹配类型={keyword_rows[0][20]}")
    
    product_targeting_rows = [row for row in rows if row[1] == '商品定向']
    print(f"商品定向行数量: {len(product_targeting_rows)}")
    if product_targeting_rows:
        print(f"示例商品定向行: 实体层级={product_targeting_rows[0][1]}, 竞价={product_targeting_rows[0][18]}, 拓展商品投放编号={product_targeting_rows[0][24]}")
    
    bid_adjustment_rows = [row for row in rows if row[1] == '竞价调整']
    print(f"竞价调整行数量: {len(bid_adjustment_rows)}")
    if bid_adjustment_rows:
        print(f"示例竞价调整行: 实体层级={bid_adjustment_rows[0][1]}, 广告位={bid_adjustment_rows[0][22]}, 百分比={bid_adjustment_rows[0][23]}")
    
    levels = set(row[1] for row in rows)
    print(f"所有实体层级: {levels}")
    
    return output_buffer

# Streamlit 应用
st.set_page_config(page_title="K EU 小赖版-SP广告批量模版工具", page_icon="🚀", layout="wide")

st.title("🚀 K EU 小赖版-SP广告批量模版工具")

st.markdown("""
### 匹配规则说明
本工具基于调研Excel文件生成SP广告批量模板（header-K EU.xlsx）。以下是核心匹配规则：

1. **活动名称匹配**：
   - 从'广告活动名称'列提取独特活动。
   - 匹配类型：活动名含'精准'或'exact' → 精准匹配；含'广泛'或'broad' → 广泛匹配；含'asin' → ASIN（商品定向）。
   - 关键词类别：从活动名匹配已知类别（如'suzhu'、'host'、'宿主'、'case'、'包'、'tape'），或从列名前缀动态提取（H-Q列）。

2. **关键词提取**：
   - 精准/广泛：从H-Q列匹配类别+匹配类型（如'host精准'列）提取关键词（去重）。
   - ASIN：活动名作为列名精确匹配，提取ASIN列表。

3. **否定关键词**：
   - 广泛组：S列（否定精准）、T列（否定词组）。
   - 精准组：根据类别（如'host' → U/V列；'case' → W/X列）。
   - 否定ASIN：从'否定ASIN'列提取。

4. **默认值与结构**：
   - CPC/SKU/预算/竞价：从调研列取值，否则默认（CPC=0.5, 预算=12, 组竞价=0.6）。
   - 每活动生成：竞价调整（首页首位/商品页面，+900%）、活动、组、商品广告、关键词/商品定向、否定项。
   - 检查重复：H-Q列关键词重复将报错。

5. **列要求**：
   - 必须：'广告活动名称'、CPC/SKU/广告组默认竞价/预算（可选，默认值）。
   - 关键词：H-Q列（索引7-16）。
   - 否定：S-X列（索引18-23）。

上传调研Excel（默认第一个Sheet），点击生成下载模板！
""")

uploaded_file = st.file_uploader("上传调研Excel文件", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        df_survey = pd.read_excel(uploaded_file, sheet_name=0)
        st.success(f"文件上传成功！形状：{df_survey.shape}")
        st.dataframe(df_survey.head(), use_container_width=True)
    except Exception as e:
        st.error(f"读取文件出错：{e}")
        st.stop()

if st.button("生成表头", type="primary"):
    if 'df_survey' not in locals():
        st.warning("请先上传文件！")
    else:
        with st.spinner("正在生成表头..."):
            output_buffer = generate_header_from_survey(df_survey)
            if output_buffer is not None:
                st.success("生成完成！")
                st.download_button(
                    label="下载 header-K EU.xlsx",
                    data=output_buffer.getvalue(),
                    file_name="header-K EU.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("生成失败，请检查日志。")