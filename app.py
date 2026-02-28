import streamlit as st
import pandas as pd
import io

# === 页面配置 ===
st.set_page_config(page_title="金融案件数据分析助手", page_icon="⚖️", layout="wide")

st.title("⚖️ 金融案件数据分析助手 Pro (智能版)")
st.markdown("### 专为律师设计的证券操纵案数据清洗与分析工具")

# === 侧边栏：上传与设置 ===
with st.sidebar:
    st.header("1. 上传案卷数据")
    uploaded_file = st.file_uploader("请上传 Excel 交易流水", type=["xlsx", "xls"])
    
    st.header("2. 输入分析目标")
    target_stock_code = st.text_input("目标股票代码", value="", placeholder="例如: 002776", help="请输入要分析的股票代码")
    
    run_button = st.button("🚀 开始分析", type="primary")

# === 智能列名清洗函数 (核心修复) ===
def smart_rename_columns(df):
    """
    自动识别常见列名变体，统一修改为标准名称
    """
    # 1. 去除列名中的空格和特殊字符
    df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]
    
    # 2. 定义同义词词典
    column_mapping = {
        # 标准名 : [可能的变体]
        '证券代码': ['证券代码', '代码', '证券ID', '股票代码', '证券代号'],
        '成交数量': ['成交数量', '成交量', '数量', '发生数量', '股数', '成交股数'],
        '成交金额': ['成交金额', '金额', '发生金额', '清算金额'],
        '交易日期': ['交易日期', '成交日期', '日期', '发生日期', '业务日期']
    }
    
    # 3. 遍历并重命名
    new_columns = {}
    for col in df.columns:
        for standard_name, variants in column_mapping.items():
            if col in variants:
                new_columns[col] = standard_name
                break
    
    if new_columns:
        df.rename(columns=new_columns, inplace=True)
    
    return df

# === 核心逻辑函数 ===
def clean_and_process(file, target_code):
    if not target_code:
        return None, None, None, "⚠️ 请先在左侧输入目标股票代码！"

    try:
        xls = pd.ExcelFile(file)
    except Exception as e:
        return None, None, None, f"文件读取失败。请确认文件未加密且格式正确。错误: {str(e)}"

    all_data = []
    debug_info = [] # 用于记录每张表的读取情况，方便排错
    
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, sheet_name in enumerate(xls.sheet_names):
        status_text.text(f"正在智能扫描表格: {sheet_name}...")
        
        # 尝试跳过前几行寻找真正的表头（Header Hunter）
        found_valid_header = False
        for skip_rows in range(5): # 尝试跳过 0 到 4 行
            try:
                # 读取数据
                df = pd.read_excel(xls, sheet_name=sheet_name, header=skip_rows)
                # 智能重命名
                df = smart_rename_columns(df)
                
                # 检查是否包含关键列
                if '证券代码' in df.columns and '成交数量' in df.columns:
                    # 再次清洗空行
                    df.dropna(subset=['证券代码'], inplace=True)
                    
                    # 尝试统一日期格式
                    if '交易日期' in df.columns:
                        df['交易日期'] = pd.to_datetime(df['交易日期'], errors='coerce').dt.date
                    
                    all_data.append(df)
                    found_valid_header = True
                    debug_info.append(f"✅ Sheet '{sheet_name}': 成功识别 (跳过 {skip_rows} 行)")
                    break # 找到了就停止尝试skip_rows
            except Exception:
                continue
        
        if not found_valid_header:
            # 如果没找到，记录一下这张表里到底有什么列，方便用户看
            try:
                temp_df = pd.read_excel(xls, sheet_name=sheet_name, nrows=1)
                cols_found = list(temp_df.columns)
            except:
                cols_found = "无法读取"
            debug_info.append(f"❌ Sheet '{sheet_name}': 未找到关键列。程序看到的列名是: {cols_found}")

        progress_bar.progress((i + 1) / len(xls.sheet_names))

    if not all_data:
        # 如果彻底失败，展示详细的诊断信息
        error_msg = "❌ 未找到有效数据表格。\n\n**程序诊断报告：**\n" + "\n".join(debug_info)
        return None, None, None, error_msg

    merged_df = pd.concat(all_data, ignore_index=True)
    
    # 最终数据转换
    merged_df['证券代码'] = merged_df['证券代码'].astype(str).str.zfill(6)
    merged_df['成交数量'] = pd.to_numeric(merged_df['成交数量'], errors='coerce').fillna(0)
    
    target_df = merged_df[merged_df['证券代码'] == target_code].copy()
    
    return merged_df, target_df, xls.sheet_names, "✅ 分析完成"

# === 同日交易分析逻辑 ===
def analyze_same_day(full_df, target_code, target_date_list):
    if '交易日期' not in full_df.columns:
        return pd.DataFrame(columns=["错误: 缺少【交易日期】列，无法分析同日交易"])

    daily_mask = full_df['交易日期'].isin(target_date_list)
    daily_df = full_df[daily_mask].copy()
    
    daily_stats = daily_df.groupby('交易日期')['成交数量'].sum().reset_index()
    daily_stats.rename(columns={'成交数量': '当日全账户总成交量'}, inplace=True)
    
    target_only = daily_df[daily_df['证券代码'] == target_code]
    target_stats = target_only.groupby('交易日期')['成交数量'].sum().reset_index()
    target_stats.rename(columns={'成交数量': '当日目标股票成交量'}, inplace=True)
    
    result = pd.merge(daily_stats, target_stats, on='交易日期', how='left').fillna(0)
    result['目标占比(%)'] = (result['当日目标股票成交量'] / result['当日全账户总成交量'] * 100).round(2)
    return result

# === 主界面逻辑 ===
if run_button and uploaded_file is not None:
    merged_df, target_df, sheet_list, message = clean_and_process(uploaded_file, target_stock_code)
    
    if merged_df is not None:
        st.success(message)
        
        # 基础指标
        total_vol = merged_df['成交数量'].sum()
        target_vol = target_df['成交数量'].sum()
        ratio_vol = (target_vol / total_vol * 100) if total_vol > 0 else 0
        
        # 同日交易分析
        mixed_days = 0
        single_days = 0
        same_day_table = pd.DataFrame()

        if '交易日期' in merged_df.columns:
            target_dates = target_df['交易日期'].dropna().unique()
            days_trade_target = len(target_dates)
            
            for date in target_dates:
                day_data = merged_df[merged_df['交易日期'] == date]
                day_codes = day_data['证券代码'].unique()
                if len(day_codes) > 1:
                    mixed_days += 1
                else:
                    single_days += 1
            
            same_day_table = analyze_same_day(merged_df, target_stock_code, target_dates)
        else:
            st.warning("⚠️ 警告：未找到【交易日期】相关列，跳过同日交易分析。请检查Excel列名。")

        # === 页面展示区 ===
        st.subheader("📊 核心持仓占比")
        c1, c2, c3 = st.columns(3)
        c1.metric("总成交量占比", f"{ratio_vol:.2f}%")
        c2.metric("混合交易天数", f"{mixed_days} 天")
        c3.metric("单一交易天数", f"{single_days} 天")

        st.divider()
        st.subheader("📅 同日交易深度分析")
        st.dataframe(same_day_table, use_container_width=True)

        with st.expander("点击查看目标股票所有交易明细"):
            st.dataframe(target_df)

        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            target_df.to_excel(writer, sheet_name='目标股票明细', index=False)
            same_day_table.to_excel(writer, sheet_name='同日交易分析', index=False)
            
        st.download_button(
            label="📥 下载分析报告",
            data=output.getvalue(),
            file_name=f"案件分析报告_{target_stock_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    else:
        # 如果失败，这里会显示详细的诊断信息
        st.error(message)

elif run_button and uploaded_file is None:
    st.warning("请先在左侧上传 Excel 文件！")