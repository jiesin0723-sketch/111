import streamlit as st
import pandas as pd
import io
import re

# === 页面配置 ===
st.set_page_config(page_title="金融案件数据分析助手", page_icon="⚖️", layout="wide")

st.title("⚖️ 金融案件数据分析助手 Pro (智能版)")
st.markdown("### 专为律师设计的证券操纵案数据清洗与分析工具")

# === 侧边栏：上传与设置 ===
with st.sidebar:
    st.header("1. 上传案卷数据")
    uploaded_file = st.file_uploader("请上传 Excel 交易流水", type=["xlsx", "xls"])

    st.header("2. 输入分析目标")
    target_stock_code = st.text_input(
        "目标股票代码",
        value="",
        placeholder="例如: 002776",
        help="请输入要分析的股票代码",
    )

    run_button = st.button("🚀 开始分析", type="primary")


def normalize_stock_code(value: str) -> str:
    """统一股票代码格式，处理 600519.0 / 空格 等问题。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    if text.isdigit():
        text = text.zfill(6)
    return text


# === 智能列名清洗函数 (核心修复) ===
def smart_rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    自动识别常见列名变体，统一修改为标准名称
    """
    # 1. 去除列名中的空格和换行
    df.columns = [str(c).strip().replace("\n", "").replace(" ", "") for c in df.columns]

    # 2. 定义同义词词典
    column_mapping = {
        "证券代码": ["证券代码", "代码", "证券ID", "股票代码", "证券代号"],
        "成交数量": ["成交数量", "成交量", "数量", "发生数量", "股数", "成交股数"],
        "成交金额": ["成交金额", "金额", "发生金额", "清算金额"],
        "成交价格": ["成交价格", "价格", "成交均价", "成交单价"],
        "交易日期": ["交易日期", "成交日期", "日期", "发生日期", "业务日期"],
        "买卖方向": ["买卖方向", "交易方向", "委托方向", "方向", "买卖标志"],
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
    target_code = normalize_stock_code(target_code)
    if not target_code:
        return None, None, None, "⚠️ 请先在左侧输入目标股票代码！"

    try:
        xls = pd.ExcelFile(file)
    except Exception as e:
        return None, None, None, f"文件读取失败。请确认文件未加密且格式正确。错误: {str(e)}"

    all_data = []
    debug_info = []  # 记录每张表读取情况，方便排错

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, sheet_name in enumerate(xls.sheet_names):
        status_text.text(f"正在智能扫描表格: {sheet_name}...")

        # 尝试跳过前几行寻找真正的表头
        found_valid_header = False
        for skip_rows in range(5):  # 尝试跳过 0 到 4 行
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=skip_rows, dtype=str)
                df = smart_rename_columns(df)

                # 保持原有能力：多Sheet合并 + 空行清洗 + 关键列检查
                if "证券代码" in df.columns and "成交数量" in df.columns:
                    # 空行清洗：证券代码为空的记录剔除
                    df.dropna(subset=["证券代码"], inplace=True)

                    # 交易日期标准化
                    if "交易日期" in df.columns:
                        df["交易日期"] = pd.to_datetime(df["交易日期"], errors="coerce").dt.date

                    all_data.append(df)
                    found_valid_header = True
                    debug_info.append(f"✅ Sheet '{sheet_name}': 成功识别 (跳过 {skip_rows} 行)")
                    break
            except Exception:
                continue

        if not found_valid_header:
            try:
                temp_df = pd.read_excel(xls, sheet_name=sheet_name, nrows=1)
                cols_found = list(temp_df.columns)
            except Exception:
                cols_found = "无法读取"
            debug_info.append(f"❌ Sheet '{sheet_name}': 未找到关键列。程序看到的列名是: {cols_found}")

        progress_bar.progress((i + 1) / len(xls.sheet_names))

    status_text.empty()
    progress_bar.empty()

    if not all_data:
        error_msg = "❌ 未找到有效数据表格。\n\n**程序诊断报告：**\n" + "\n".join(debug_info)
        return None, None, None, error_msg

    merged_df = pd.concat(all_data, ignore_index=True)

    # 最终数据转换
    merged_df["证券代码"] = merged_df["证券代码"].map(normalize_stock_code)
    merged_df["成交数量"] = pd.to_numeric(merged_df["成交数量"], errors="coerce").fillna(0)

    if "成交金额" in merged_df.columns:
        merged_df["成交金额"] = pd.to_numeric(merged_df["成交金额"], errors="coerce")

    if "成交价格" in merged_df.columns:
        merged_df["成交价格"] = pd.to_numeric(merged_df["成交价格"], errors="coerce")

    target_df = merged_df[merged_df["证券代码"] == target_code].copy()

    return merged_df, target_df, xls.sheet_names, "✅ 分析完成"


# === 同日交易分析逻辑 ===
def analyze_same_day(full_df, target_code, target_date_list):
    if "交易日期" not in full_df.columns:
        return pd.DataFrame(columns=["错误: 缺少【交易日期】列，无法分析同日交易"])

    daily_mask = full_df["交易日期"].isin(target_date_list)
    daily_df = full_df[daily_mask].copy()

    # 使用绝对值，避免卖出为负数影响占比
    daily_df["成交数量_abs"] = daily_df["成交数量"].abs()

    daily_stats = daily_df.groupby("交易日期")["成交数量_abs"].sum().reset_index()
    daily_stats.rename(columns={"成交数量_abs": "当日全账户总成交量"}, inplace=True)

    target_only = daily_df[daily_df["证券代码"] == target_code]
    target_stats = target_only.groupby("交易日期")["成交数量_abs"].sum().reset_index()
    target_stats.rename(columns={"成交数量_abs": "当日目标股票成交量"}, inplace=True)

    result = pd.merge(daily_stats, target_stats, on="交易日期", how="left").fillna(0)
    result["目标占比(%)"] = (
        result["当日目标股票成交量"] / result["当日全账户总成交量"] * 100
    ).round(2)
    result = result.sort_values("交易日期")
    return result


# === 新增：成交均价折线图数据 ===
def build_price_trend_df(target_df: pd.DataFrame):
    """
    生成“日期-成交均价”趋势数据：
    - 优先使用“买入”记录（若存在买卖方向列且能识别买入）
    - 均价优先按 成交金额/成交数量 计算；否则按成交价格简单均值
    """
    if "交易日期" not in target_df.columns:
        return pd.DataFrame(columns=["交易日期", "成交均价"]), "缺少【交易日期】列，无法绘图。"

    work_df = target_df.copy()
    work_df = work_df.dropna(subset=["交易日期"])

    if work_df.empty:
        return pd.DataFrame(columns=["交易日期", "成交均价"]), "目标股票无有效日期数据。"

    source_desc = "全部交易"

    # 若有买卖方向，优先取买入记录
    if "买卖方向" in work_df.columns:
        buy_mask = work_df["买卖方向"].astype(str).str.contains("买", na=False)
        if buy_mask.any():
            work_df = work_df[buy_mask].copy()
            source_desc = "买入交易"

    if work_df.empty:
        return pd.DataFrame(columns=["交易日期", "成交均价"]), "未识别到可用于计算均价的记录。"

    # 优先按成交金额/数量计算加权均价
    if "成交金额" in work_df.columns and work_df["成交金额"].notna().any():
        work_df["成交金额"] = pd.to_numeric(work_df["成交金额"], errors="coerce")
        work_df["成交数量_abs"] = pd.to_numeric(work_df["成交数量"], errors="coerce").abs()
        temp = work_df.dropna(subset=["成交金额", "成交数量_abs"]).copy()
        temp = temp[temp["成交数量_abs"] > 0]

        if not temp.empty:
            trend_df = (
                temp.groupby("交易日期", as_index=False)
                .agg(成交金额合计=("成交金额", "sum"), 成交数量合计=("成交数量_abs", "sum"))
            )
            trend_df["成交均价"] = trend_df["成交金额合计"] / trend_df["成交数量合计"]
            trend_df = trend_df[["交易日期", "成交均价"]].sort_values("交易日期")
            return trend_df, f"折线图基于【{source_desc}】，均价按 成交金额/成交数量 计算。"

    # 退化到成交价格简单均值
    if "成交价格" not in work_df.columns:
        return pd.DataFrame(columns=["交易日期", "成交均价"]), "缺少【成交价格/成交金额】列，无法计算均价。"

    work_df["成交价格"] = pd.to_numeric(work_df["成交价格"], errors="coerce")
    temp = work_df.dropna(subset=["成交价格"]).copy()

    if temp.empty:
        return pd.DataFrame(columns=["交易日期", "成交均价"]), "成交价格均为空，无法绘图。"

    trend_df = temp.groupby("交易日期", as_index=False)["成交价格"].mean()
    trend_df.rename(columns={"成交价格": "成交均价"}, inplace=True)
    trend_df = trend_df.sort_values("交易日期")
    return trend_df, f"折线图基于【{source_desc}】，均价按成交价格简单平均计算。"


# === 主界面逻辑 ===
if run_button and uploaded_file is not None:
    merged_df, target_df, sheet_list, message = clean_and_process(uploaded_file, target_stock_code)
    target_code_norm = normalize_stock_code(target_stock_code)

    if merged_df is not None:
        st.success(message)

        if target_df.empty:
            st.warning(f"未检索到目标股票【{target_code_norm}】的交易记录，请检查代码是否正确。")
            st.stop()

        # 基础指标
        total_vol = merged_df["成交数量"].abs().sum()
        target_vol = target_df["成交数量"].abs().sum()
        ratio_vol = (target_vol / total_vol * 100) if total_vol > 0 else 0

        # 同日交易分析
        mixed_days = 0
        single_days = 0
        same_day_table = pd.DataFrame()

        if "交易日期" in merged_df.columns:
            target_dates = target_df["交易日期"].dropna().unique()
            days_trade_target = len(target_dates)

            for date in target_dates:
                day_data = merged_df[merged_df["交易日期"] == date]
                day_codes = day_data["证券代码"].dropna().unique()
                if len(day_codes) > 1:
                    mixed_days += 1
                else:
                    single_days += 1

            same_day_table = analyze_same_day(merged_df, target_code_norm, target_dates)
        else:
            days_trade_target = 0
            st.warning("⚠️ 警告：未找到【交易日期】相关列，跳过同日交易分析。请检查Excel列名。")

        mixed_single_ratio = (mixed_days / single_days * 100) if single_days > 0 else 0

        # 新增：成交均价趋势
        price_trend_df, trend_note = build_price_trend_df(target_df)

        # === 页面展示区 ===
        st.subheader("📊 核心持仓占比")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("总成交量占比", f"{ratio_vol:.2f}%")
        c2.metric("混合交易天数", f"{mixed_days} 天")
        c3.metric("单一交易天数", f"{single_days} 天")
        c4.metric("混合/单一天数比", f"{mixed_single_ratio:.2f}%")

        st.divider()
        st.subheader("📅 同日交易深度分析")
        st.dataframe(same_day_table, use_container_width=True)

        st.divider()
        st.subheader("📈 成交均价趋势折线图")
        st.caption(trend_note)
        if not price_trend_df.empty:
            chart_df = price_trend_df.copy()
            chart_df["交易日期"] = pd.to_datetime(chart_df["交易日期"], errors="coerce")
            chart_df = chart_df.dropna(subset=["交易日期"]).sort_values("交易日期")
            st.line_chart(chart_df.set_index("交易日期")["成交均价"], height=320)
            st.dataframe(price_trend_df, use_container_width=True)
        else:
            st.info("暂无可用于绘图的数据。")

        with st.expander("点击查看目标股票所有交易明细"):
            st.dataframe(target_df, use_container_width=True)

        # 导出 Excel：测算结果 + 筛选数据都写入
        summary_df = pd.DataFrame(
            {
                "指标": [
                    "目标股票代码",
                    "合并Sheet数量",
                    "全账户交易记录数",
                    "目标股票交易记录数",
                    "全账户总成交量(绝对值)",
                    "目标股票总成交量(绝对值)",
                    "目标成交量占比(%)",
                    "目标股票涉及交易日期数",
                    "同日交易目标+其他股票天数",
                    "仅交易目标股票天数",
                    "混合/单一天数比(%)",
                ],
                "数值": [
                    target_code_norm,
                    len(sheet_list) if sheet_list is not None else 0,
                    len(merged_df),
                    len(target_df),
                    round(float(total_vol), 2),
                    round(float(target_vol), 2),
                    round(float(ratio_vol), 2),
                    days_trade_target,
                    mixed_days,
                    single_days,
                    round(float(mixed_single_ratio), 2),
                ],
            }
        )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="分析汇总", index=False)
            target_df.to_excel(writer, sheet_name="目标股票明细", index=False)
            same_day_table.to_excel(writer, sheet_name="同日交易分析", index=False)
            price_trend_df.to_excel(writer, sheet_name="成交均价趋势", index=False)

        st.download_button(
            label="📥 下载分析报告",
            data=output.getvalue(),
            file_name=f"案件分析报告_{target_code_norm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    else:
        # 如果失败，这里会显示详细诊断信息
        st.error(message)

elif run_button and uploaded_file is None:
    st.warning("请先在左侧上传 Excel 文件！")
```