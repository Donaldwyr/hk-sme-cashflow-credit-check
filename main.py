# -*- coding: utf-8 -*-
"""
hk-sme-cashflow-credit-check
中小企业信贷初筛工具 (HK Commercial Banking SME Credit Screening)

功能:
  1. 读取 sample_sme_fin.xlsx (Sheet: BS 资产负债表 / IS 利润表)
  2. 计算商行信贷核心指标:
       流动比率 / 速动比率 / 资产负债率 / 利息保障倍数 / 经营性现金流简易测算
  3. 每个指标给出银行视角评级: 安全 / 关注 / 预警
  4. 自动绘图(现金流趋势 + 偿债指标柱状图) -> output_chart.png
  5. 输出结构化信贷初筛报告 -> credit_report.md

依赖: pandas / openpyxl / matplotlib (均为纯 CPU 库, 无需 GPU/CUDA)
用法: python main.py
"""

import os
import datetime
import pandas as pd
import matplotlib

matplotlib.use("Agg")  # 无界面后台绘图, 不依赖 GPU/CUDA, 无需显示器
import matplotlib.pyplot as plt

# 中文字体(Windows 自带微软雅黑/黑体), 避免图表中文乱码
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

# ==================== 文件路径 ====================
INPUT_FILE = "sample_sme_fin.xlsx"   # 输入: 财务报表
CHART_FILE = "output_chart.png"      # 输出: 图表
REPORT_FILE = "credit_report.md"     # 输出: 信贷初筛报告

# ==================== 科目名称(需与 Excel 第一列一致) ====================
BS_CURRENT_ASSETS = "流动资产合计"   # 流动资产
BS_INVENTORY = "存货"                # 存货
BS_CURRENT_LIAB = "流动负债合计"     # 流动负债
BS_TOTAL_ASSETS = "资产总计"         # 总资产
BS_TOTAL_LIAB = "负债合计"           # 总负债
BS_AR = "应收账款"                   # 应收账款(现金流测算用)
BS_AP = "应付账款"                   # 应付账款(现金流测算用)

IS_NET_PROFIT = "净利润"             # 净利润
IS_INTEREST = "利息费用"             # 利息费用
IS_DEP_AMORT = "折旧与摊销"          # 折旧与摊销
IS_PBT = "税前利润"                  # 税前利润

# ==================== 评级阈值(汇丰香港 CB SME 风控标准) ====================
# 判定规则: 安全 >= 安全线; 关注 = [预警线, 安全线); 预警 < 预警线
# (资产负债率为"越低越好"型: 安全 <= 安全线; 关注 = (安全线, 预警线]; 预警 > 预警线)
CURRENT_RATIO_SAFE = 1.8            # 流动比率: >=1.8 安全; [1.2, 1.8) 关注; <1.2 预警
CURRENT_RATIO_WARN = 1.2
QUICK_RATIO_SAFE = 1.1              # 速动比率: >=1.1 安全; [0.7, 1.1) 关注; <0.7 预警
QUICK_RATIO_WARN = 0.7
DEBT_RATIO_SAFE = 55                # 资产负债率(%): <=55% 安全; (55%, 70%] 关注; >70% 预警
DEBT_RATIO_WARN = 70
INTEREST_COVERAGE_SAFE = 8          # 利息保障倍数: >=8 安全; [3, 8) 关注; <3 预警
INTEREST_COVERAGE_WARN = 3


# ==================== 通用工具函数 ====================
def rate_higher_better(value, safe_min, warn_min):
    """越高越好型指标评级: 安全 / 关注 / 预警"""
    if value is None or pd.isna(value):
        return "N/A"
    if value >= safe_min:
        return "安全"
    if value < warn_min:
        return "预警"
    return "关注"


def rate_lower_better(value, safe_max, warn_max):
    """越低越好型指标评级: 安全 / 关注 / 预警"""
    if value is None or pd.isna(value):
        return "N/A"
    if value <= safe_max:
        return "安全"
    if value > warn_max:
        return "预警"
    return "关注"


def safe_div(numerator, denominator):
    """安全除法: 分母为 0 或缺失时返回 NaN, 避免除零报错"""
    if denominator is None or pd.isna(denominator) or denominator == 0:
        return float("nan")
    return numerator / denominator


# ==================== 数据读取 ====================
def read_sheet(path, sheet_name):
    """
    读取 Excel 指定 sheet。
    约定: 第 1 列 = 科目名称(项目), 第 2 列起 = 各会计期间数值。
    返回: (科目名列表, 期间列表, 数值 DataFrame)
    """
    df = pd.read_excel(path, sheet_name=sheet_name)
    # 去除科目名称两端空格, 便于精确匹配
    df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.strip()
    items = df.iloc[:, 0].tolist()
    periods = [str(c).strip() for c in df.columns[1:]]
    values = df.iloc[:, 1:].apply(pd.to_numeric, errors="coerce")
    return items, periods, values


def get_series(items, values, *item_names):
    """
    按科目名称取数, 返回按期间对齐的数值列表。
    支持传入多个候选名称(容错, 如 "资产总计" / "总资产")。
    """
    for name in item_names:
        if name in items:
            idx = items.index(name)
            return values.iloc[idx].tolist()
    raise KeyError(f"未找到科目: {item_names}")


# ==================== 指标计算 ====================
def compute_metrics(bs, is_):
    """计算全部指标, 返回按期间对齐的指标 DataFrame"""
    bs_items, bs_periods, bs_vals = bs
    is_items, is_periods, is_vals = is_

    # 期间对齐校验: BS 与 IS 的期间列应一致
    if bs_periods != is_periods:
        raise ValueError(f"BS 与 IS 期间列不一致: {bs_periods} vs {is_periods}")

    periods = bs_periods
    n = len(periods)

    # ---- 资产负债表取数 ----
    ca = get_series(bs_items, bs_vals, BS_CURRENT_ASSETS, "流动资产")       # 流动资产
    cl = get_series(bs_items, bs_vals, BS_CURRENT_LIAB, "流动负债")         # 流动负债
    inventory = get_series(bs_items, bs_vals, BS_INVENTORY, "存货净额")     # 存货
    ta = get_series(bs_items, bs_vals, BS_TOTAL_ASSETS, "总资产")           # 总资产
    tl = get_series(bs_items, bs_vals, BS_TOTAL_LIAB, "总负债")             # 总负债
    ar = get_series(bs_items, bs_vals, BS_AR)                               # 应收账款
    ap = get_series(bs_items, bs_vals, BS_AP)                               # 应付账款

    # ---- 利润表取数 ----
    net_profit = get_series(is_items, is_vals, IS_NET_PROFIT)               # 净利润
    interest = get_series(is_items, is_vals, IS_INTEREST)                   # 利息费用
    dep_amort = get_series(is_items, is_vals, IS_DEP_AMORT)                 # 折旧与摊销
    pbt = get_series(is_items, is_vals, IS_PBT)                             # 税前利润

    # ---- 核心指标计算 ----
    # 流动比率 = 流动资产 / 流动负债 (衡量短期偿债能力)
    current_ratio = [safe_div(ca[i], cl[i]) for i in range(n)]

    # 速动比率 = (流动资产 - 存货) / 流动负债 (剔除变现较慢的存货)
    quick_ratio = [safe_div(ca[i] - inventory[i], cl[i]) for i in range(n)]

    # 资产负债率 = 总负债 / 总资产 (衡量整体杠杆水平)
    debt_ratio = [safe_div(tl[i], ta[i]) for i in range(n)]

    # EBIT(息税前利润) = 税前利润 + 利息费用
    # 说明: 利息费用已在税前利润中扣除, 加回后得到未扣利息、未扣税的利润。
    #       在无重大营业外收支时, 该值等于"营业利润"。
    ebit = [pbt[i] + interest[i] for i in range(n)]

    # 利息保障倍数 = EBIT / 利息费用 (衡量利润对利息的覆盖能力)
    icr = [safe_div(ebit[i], interest[i]) for i in range(n)]

    # ---- 经营性现金流简易测算(间接法) ----
    # 简易现金收益 = 净利润 + 折旧与摊销 + 利息费用
    #   加回非现金支出(折旧摊销)与利息, 得到"付息前经营现金口径"。
    cash_earnings = [net_profit[i] + dep_amort[i] + interest[i] for i in range(n)]

    # 调整后经营性现金流 = 简易现金收益 - Δ应收账款 - Δ存货 + Δ应付账款
    #   应收/存货增加会占用现金(减项), 应付增加相当于占用上游资金(加项)。
    #   首期无上年数据, 无法计算营运资金变动, 记为 NaN。
    ocf = [float("nan")] * n
    for i in range(1, n):
        delta_ar = ar[i] - ar[i - 1]
        delta_inv = inventory[i] - inventory[i - 1]
        delta_ap = ap[i] - ap[i - 1]
        ocf[i] = cash_earnings[i] - delta_ar - delta_inv + delta_ap

    metrics = pd.DataFrame(
        {
            "流动比率": current_ratio,
            "速动比率": quick_ratio,
            "资产负债率": debt_ratio,
            "利息保障倍数": icr,
            "EBIT": ebit,
            "现金收益(简易)": cash_earnings,
            "经营性现金流(调整)": ocf,
        },
        index=periods,
    )
    return metrics


# ==================== 评级 ====================
def rate_metrics(metrics):
    """为最新一期指标给出银行视角评级 (汇丰香港 CB SME 风控标准)"""
    latest = metrics.iloc[-1]
    ratings = {
        "流动比率": rate_higher_better(latest["流动比率"], CURRENT_RATIO_SAFE, CURRENT_RATIO_WARN),
        "速动比率": rate_higher_better(latest["速动比率"], QUICK_RATIO_SAFE, QUICK_RATIO_WARN),
        # 资产负债率公式仍为小数(总负债/总资产), 阈值单位为百分数, 故此处 x100
        "资产负债率": rate_lower_better(latest["资产负债率"] * 100, DEBT_RATIO_SAFE, DEBT_RATIO_WARN),
        "利息保障倍数": rate_higher_better(latest["利息保障倍数"], INTEREST_COVERAGE_SAFE, INTEREST_COVERAGE_WARN),
    }
    # 经营性现金流: 正值=安全(正向经营造血), 负值=预警
    ocf_val = latest["经营性现金流(调整)"]
    if pd.isna(ocf_val):
        ratings["经营性现金流"] = "N/A"
    else:
        ratings["经营性现金流"] = "安全" if ocf_val > 0 else "预警"
    return latest, ratings


def overall_rating(ratings):
    """综合评级 (汇丰香港 CB SME 风控标准):
    - 任意 2 项及以上单项预警 -> 预警
    - 任意 1 项单项预警 -> 关注 (有预警即不得判安全, 兜底最低为关注)
    - 无预警但存在单项关注(>=1 项) -> 关注
    - 仅当 warn_count == 0 且 watch_count == 0 时 -> 安全
    """
    warn_count = sum(1 for v in ratings.values() if v == "预警")
    watch_count = sum(1 for v in ratings.values() if v == "关注")
    if warn_count >= 2:
        return "预警"
    if warn_count >= 1:      # 哪怕仅 1 项预警, 最低判关注, 不得判安全
        return "关注"
    if watch_count >= 1:     # 无预警但存在关注项 -> 关注
        return "关注"
    return "安全"            # 仅 warn_count == 0 且 watch_count == 0


# ==================== 绘图 ====================
def plot_charts(metrics, path):
    """绘制现金流趋势与偿债指标柱状图, 保存为 PNG"""
    periods = metrics.index.tolist()
    x = range(len(periods))
    width = 0.35

    fig, axes = plt.subplots(3, 1, figsize=(10, 12))

    # 图1: 现金流趋势
    ax = axes[0]
    ax.bar(
        [i - width / 2 for i in x],
        metrics["现金收益(简易)"],
        width,
        label="现金收益 (净利润+折旧摊销+利息)",
        color="#4472C4",
    )
    ax.bar(
        [i + width / 2 for i in x],
        metrics["经营性现金流(调整)"],
        width,
        label="调整后经营性现金流",
        color="#ED7D31",
    )
    ax.set_xticks(list(x))
    ax.set_xticklabels(periods)
    ax.set_title("现金流趋势 (Cash Flow Trend)")
    ax.set_ylabel("千港元 HKD'000")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    # 图2: 流动性与速动比率柱状图
    ax = axes[1]
    ax.bar(
        [i - width / 2 for i in x],
        metrics["流动比率"],
        width,
        label="流动比率",
        color="#70AD47",
    )
    ax.bar(
        [i + width / 2 for i in x],
        metrics["速动比率"],
        width,
        label="速动比率",
        color="#FFC000",
    )
    ax.axhline(1.0, color="red", linestyle="--", linewidth=1, label="警戒线 1.0x")
    ax.set_xticks(list(x))
    ax.set_xticklabels(periods)
    ax.set_title("偿债指标 - 流动性与速动比率")
    ax.set_ylabel("倍数 (x)")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    # 图3: 资产负债率(柱) + 利息保障倍数(线, 双轴)
    ax = axes[2]
    ax.bar(
        list(x),
        [v * 100 for v in metrics["资产负债率"]],
        width * 2,
        label="资产负债率 (%)",
        color="#A5A5A5",
    )
    ax.set_ylabel("资产负债率 (%)")
    ax2 = ax.twinx()
    ax2.plot(
        list(x),
        metrics["利息保障倍数"],
        marker="o",
        color="#5B9BD5",
        label="利息保障倍数 (x)",
    )
    ax2.set_ylabel("利息保障倍数 (x)")
    ax.set_xticks(list(x))
    ax.set_xticklabels(periods)
    ax.set_title("偿债指标 - 资产负债率与利息保障倍数")
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper left")
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


# ==================== 报告生成 ====================
def fmt(v, digits=2, pct=False):
    """格式化数值, 处理 NaN"""
    if v is None or pd.isna(v):
        return "N/A"
    if pct:
        return f"{v * 100:.{digits}f}%"
    return f"{v:.{digits}f}"


def generate_report(metrics, latest, ratings, path):
    """生成结构化信贷初筛报告 (Markdown)"""
    periods = metrics.index.tolist()
    last = periods[-1]

    lines = []
    lines.append("# 中小企业信贷初筛报告 (SME Credit Screening Report)")
    lines.append("")
    lines.append(f"- 数据来源: `{INPUT_FILE}`")
    lines.append(f"- 数据期间: {', '.join(periods)}")
    lines.append(f"- 报告基准期间: **{last}** (最新一期)")
    lines.append(f"- 生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")

    # ---- 1. 核心指标与评级 ----
    lines.append("## 1. 核心财务指标与银行评级")
    lines.append("")
    lines.append("| 指标 | 计算公式 | 数值 | 银行评级 |")
    lines.append("|------|----------|------|----------|")
    lines.append(
        f"| 流动比率 | 流动资产 / 流动负债 | {fmt(latest['流动比率'])}x | {ratings['流动比率']} |"
    )
    lines.append(
        f"| 速动比率 | (流动资产 - 存货) / 流动负债 | {fmt(latest['速动比率'])}x | {ratings['速动比率']} |"
    )
    lines.append(
        f"| 资产负债率 | 总负债 / 总资产 | {fmt(latest['资产负债率'], pct=True)} | {ratings['资产负债率']} |"
    )
    lines.append(
        f"| 利息保障倍数 | EBIT / 利息费用 | {fmt(latest['利息保障倍数'])}x | {ratings['利息保障倍数']} |"
    )
    lines.append(
        f"| 简易现金收益 | 净利润 + 折旧摊销 + 利息费用 | {fmt(latest['现金收益(简易)'], digits=0)} 千港元 | - |"
    )
    lines.append(
        f"| 经营性现金流(调整) | 简易现金收益 - Δ应收 - Δ存货 + Δ应付 | {fmt(latest['经营性现金流(调整)'], digits=0)} 千港元 | {ratings['经营性现金流']} |"
    )
    lines.append("")
    lines.append(f"**综合评级: {overall_rating(ratings)}**")
    lines.append("")

    # ---- 2. 评级阈值说明 ----
    lines.append("## 2. 评级阈值说明 (汇丰香港 CB SME 风控标准)")
    lines.append("")
    lines.append("| 指标 | 安全 | 关注 | 预警 |")
    lines.append("|------|------|------|------|")
    lines.append("| 流动比率 | >= 1.8x | 1.2x ~ 1.8x | < 1.2x |")
    lines.append("| 速动比率 | >= 1.1x | 0.7x ~ 1.1x | < 0.7x |")
    lines.append("| 资产负债率 | <= 55% | 55% ~ 70% | > 70% |")
    lines.append("| 利息保障倍数 | >= 8x | 3x ~ 8x | < 3x |")
    lines.append("")

    # ---- 3. 趋势分析 ----
    lines.append("## 3. 各期趋势")
    lines.append("")
    header = "| 指标 | " + " | ".join(periods) + " |"
    sep = "|------|" + "|".join(["------"] * len(periods)) + "|"
    lines.append(header)
    lines.append(sep)
    lines.append(
        "| 流动比率 | " + " | ".join(fmt(v) + "x" for v in metrics["流动比率"]) + " |"
    )
    lines.append(
        "| 速动比率 | " + " | ".join(fmt(v) + "x" for v in metrics["速动比率"]) + " |"
    )
    lines.append(
        "| 资产负债率 | " + " | ".join(fmt(v, pct=True) for v in metrics["资产负债率"]) + " |"
    )
    lines.append(
        "| 利息保障倍数 | " + " | ".join(fmt(v) + "x" for v in metrics["利息保障倍数"]) + " |"
    )
    lines.append(
        "| 现金收益(简易) | " + " | ".join(fmt(v, digits=0) for v in metrics["现金收益(简易)"]) + " |"
    )
    lines.append(
        "| 经营性现金流(调整) | " + " | ".join(fmt(v, digits=0) for v in metrics["经营性现金流(调整)"]) + " |"
    )
    lines.append("")

    # ---- 4. 风险提示与结论 ----
    lines.append("## 4. 风险提示与综合结论")
    lines.append("")
    lines.append("### 关注点 (Watch)")
    lines.append("")
    for key in ["流动比率", "速动比率", "资产负债率", "利息保障倍数"]:
        r = ratings[key]
        if r == "关注":
            lines.append(f"- **{key}** 处于关注区间, 建议结合行业均值与同业对标进一步评估。")
        elif r == "预警":
            lines.append(f"- **{key}** 触发预警, 需重点核查现金流与再融资能力。")
    if ratings.get("经营性现金流") == "预警":
        lines.append("- **经营性现金流** 为负, 经营造血能力不足, 需核查应收账款回收与存货周转。")
    lines.append("")
    lines.append("### 正面因素 (Comfort)")
    lines.append("")
    for key in ["流动比率", "速动比率", "资产负债率", "利息保障倍数"]:
        if ratings[key] == "安全":
            lines.append(f"- **{key}** 处于安全区间。")
    if ratings.get("经营性现金流") == "安全":
        lines.append("- **经营性现金流** 为正, 具备一定经营造血能力。")
    lines.append("")
    lines.append("### 银行视角建议")
    lines.append("")
    lines.append("- 结合客户行业地位、订单储备与抵押/担保情况综合评估授信额度。")
    lines.append("- 对处于关注/预警区间的指标, 可要求补充银行流水、应收账龄与存货周转明细。")
    lines.append("- 持续监控现金流趋势与杠杆变化, 作为贷后管理早期预警信号。")
    lines.append("")

    # ---- 5. 附图说明 ----
    lines.append("## 5. 附图")
    lines.append("")
    lines.append(f"![指标图表]({CHART_FILE})")
    lines.append("")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# ==================== 主流程 ====================
def main():
    if not os.path.exists(INPUT_FILE):
        raise SystemExit(
            f"未找到 {INPUT_FILE}。请先运行 `python generate_sample_data.py` 生成样例数据, "
            f"或将真实财务报表按同样结构放入该文件。"
        )

    bs = read_sheet(INPUT_FILE, "BS")
    is_ = read_sheet(INPUT_FILE, "IS")

    metrics = compute_metrics(bs, is_)
    latest, ratings = rate_metrics(metrics)

    # 打印控制台摘要
    print("=" * 50)
    print("中小企业信贷初筛 - 指标摘要 (最新期)")
    print("=" * 50)
    for k, v in latest.items():
        print(f"{k:20s} {fmt(v, digits=2, pct=(k == '资产负债率'))}")
    print("-" * 50)
    print("银行评级:")
    for k, v in ratings.items():
        print(f"  {k:16s} {v}")
    print(f"综合评级: {overall_rating(ratings)}")
    print("=" * 50)

    plot_charts(metrics, CHART_FILE)
    generate_report(metrics, latest, ratings, REPORT_FILE)

    print(f"已生成图表: {CHART_FILE}")
    print(f"已生成报告: {REPORT_FILE}")


if __name__ == "__main__":
    main()



