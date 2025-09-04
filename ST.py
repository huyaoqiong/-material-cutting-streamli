import math
import pandas as pd
import streamlit as st
from io import BytesIO

# 设置页面配置
st.set_page_config(
    page_title="材料裁切优化系统",
    page_icon="📏",
    layout="wide"
)


def read_demands_from_excel(file_content):
    """从Excel文件读取裁切需求"""
    try:
        df = pd.read_excel(BytesIO(file_content))

        required_columns = ['长度(mm)', '数量(根)']
        actual_columns = [col.strip() for col in df.columns]
        if not all(req_col in actual_columns for req_col in required_columns):
            return None, "Excel文件必须包含'长度(mm)'和'数量(根)'两列"

        demands = {}
        for _, row in df.iterrows():
            try:
                length = int(row['长度(mm)'])
                quantity = int(row['数量(根)'])
            except (ValueError, TypeError):
                return None, f"第{_ + 2}行：长度/数量必须是整数"

            if length <= 0 or quantity <= 0:
                return None, f"第{_ + 2}行：长度和数量必须为正数"

            if length in demands:
                demands[length] += quantity
            else:
                demands[length] = quantity

        if not demands:
            return None, "没有有效数据（至少需要1行）"

        return demands, None

    except Exception as e:
        return None, f"读取失败：{str(e)}"


def greedy_cutting_optimization(stock_length, demands):
    """贪心算法实现裁切优化"""
    for length in demands:
        if length > stock_length:
            return None, f"需求长度 {length}mm 超过原始材料长度 {stock_length}mm"

    demand_list = [{'length': length, 'quantity': qty}
                   for length, qty in demands.items()]
    demand_list.sort(key=lambda x: -x['length'])

    cutting_plans = []
    total_stock_used = 0
    total_used_length = 0

    while any(d['quantity'] > 0 for d in demand_list):
        remaining_length = stock_length
        current_pattern = {}

        for demand in demand_list:
            if demand['quantity'] <= 0 or demand['length'] > remaining_length:
                continue

            max_possible = math.floor(remaining_length / demand['length'])
            actual_count = min(max_possible, demand['quantity'])

            if actual_count > 0:
                current_pattern[demand['length']] = actual_count
                remaining_length -= actual_count * demand['length']
                demand['quantity'] -= actual_count

        if not current_pattern:
            return None, "无法满足所有需求（可能尺寸不合理）"

        pattern_used_length = stock_length - remaining_length
        utilization = (pattern_used_length / stock_length) * 100

        existing_idx = next(
            (i for i, plan in enumerate(cutting_plans) if plan['pattern'] == current_pattern),
            None
        )
        if existing_idx is not None:
            cutting_plans[existing_idx]['count'] += 1
        else:
            cutting_plans.append({
                'pattern': current_pattern,
                'count': 1,
                'utilization': utilization,
                'waste': remaining_length
            })

        total_stock_used += 1
        total_used_length += pattern_used_length

    total_utilization = (total_used_length / (total_stock_used * stock_length)) * 100
    total_waste = total_stock_used * stock_length - total_used_length

    formatted_plans = []
    for plan in cutting_plans:
        pattern_desc = " + ".join([f"{count}×{length}mm"
                                   for length, count in plan['pattern'].items()])
        formatted_plans.append({
            'pattern_desc': pattern_desc,
            'count': plan['count'],
            'utilization': f"{plan['utilization']:.2f}%",
            'waste': plan['waste']
        })

    return {
        'cutting_plans': formatted_plans,
        'total_stock_used': total_stock_used,
        'total_utilization': f"{total_utilization:.2f}%",
        'total_waste': total_waste
    }, None


def find_optimal_stock_length(demands, min_length, max_length, step_length):
    """寻找最佳原始材料长度"""
    results = []

    for length in range(min_length, max_length + 1, step_length):
        demand_copy = demands.copy()
        result, _ = greedy_cutting_optimization(length, demand_copy)
        if result:
            results.append({'stock_length': length, **result})

    if not results:
        return None, "指定范围内无可行方案"

    results.sort(key=lambda x: (-float(x['total_utilization'].strip('%')), x['total_stock_used']))
    return results, None


def main():
    """Streamlit 主界面"""
    st.title("📏 材料裁切优化系统")
    st.write("基于贪心算法的高效裁切方案计算工具（支持手动输入/Excel导入）")
    st.divider()

    # 1. 原始材料长度设置
    stock_length = st.number_input(
        "1. 原始材料长度（mm）",
        min_value=100,
        value=3800,
        help="请输入原始材料的长度（最小值100mm）"
    )

    # 2. 裁切需求输入
    st.subheader("2. 裁切需求设置")
    input_method = st.radio(
        "选择需求输入方式",
        ["手动输入（少量需求）", "Excel导入（大量需求）"],
        horizontal=True
    )

    demands = {}
    if input_method == "Excel导入（大量需求）":
        st.caption("请上传包含'长度(mm)'和'数量(根)'列的Excel文件（数据为正整数）")
        file = st.file_uploader("选择Excel文件", type=[".xlsx", ".xls"])

        if file:
            demands, error = read_demands_from_excel(file.getvalue())
            if error:
                st.error(f"Excel处理失败：{error}")
            else:
                st.success(f"✅ 导入成功！共 {len(demands)} 种裁切需求")
                # 显示需求预览
                st.dataframe(
                    pd.DataFrame([[k, v] for k, v in demands.items()],
                                 columns=['长度(mm)', '数量(根)']),
                    use_container_width=True
                )
    else:
        st.caption("请添加裁切需求（至少1项，长度和数量均为正整数）")
        demand_count = st.number_input(
            "需求数量",
            min_value=1,
            value=1,
            step=1,
            help="需要添加的裁切需求总数量"
        )

        # 动态生成输入框
        for i in range(int(demand_count)):
            col1, col2 = st.columns(2)
            with col1:
                length = st.number_input(
                    f"需求 {i + 1} - 长度（mm）",
                    min_value=1,
                    value=355 if i == 0 else 200,
                    key=f"len_{i}"
                )
            with col2:
                quantity = st.number_input(
                    f"需求 {i + 1} - 数量（根）",
                    min_value=1,
                    value=10 if i == 0 else 5,
                    key=f"qty_{i}"
                )

            # 合并相同长度的需求
            if length in demands:
                demands[length] += quantity
            else:
                demands[length] = quantity

    # 3. 计算模式选择
    st.divider()
    st.subheader("3. 计算选项")
    mode = st.radio(
        "选择计算模式",
        ["计算指定长度的裁切方案", "寻找最佳原始材料长度（推荐）"],
        horizontal=True
    )

    # 4. 执行计算
    if st.button("开始计算", type="primary"):
        if not demands:
            st.error("请先添加裁切需求！")
            return

        if mode == "寻找最佳原始材料长度（推荐）":
            col1, col2, col3 = st.columns(3)
            with col1:
                min_length = st.number_input("最小原始长度（mm）", min_value=100, value=1000)
            with col2:
                max_length = st.number_input("最大原始长度（mm）", min_value=100, value=5000)
            with col3:
                step_length = st.number_input("步长（mm）", min_value=10, value=100)

            if min_length >= max_length:
                st.error("最小长度必须小于最大长度！")
                return
            if step_length >= (max_length - min_length):
                st.error("步长过大，无法覆盖有效范围！")
                return

            with st.spinner(f"正在计算最佳长度（{min_length}~{max_length}mm）..."):
                results, error = find_optimal_stock_length(demands, min_length, max_length, step_length)

            if error:
                st.error(f"计算失败：{error}")
            else:
                st.divider()
                st.subheader("## 计算结果：最佳原始长度推荐")
                best_result = results[0]

                # 关键指标
                st.success(f"### 🌟 推荐最佳长度：{best_result['stock_length']}mm")
                st.dataframe(
                    pd.DataFrame([
                        ["总利用率", best_result['total_utilization']],
                        ["总材料用量", f"{best_result['total_stock_used']} 根"],
                        ["总废料长度", f"{best_result['total_waste']} mm"],
                        ["平均单根利用率", f"{float(best_result['total_utilization'].strip('%')):.2f}%"]
                    ], columns=["指标", "数值"]),
                    use_container_width=True,
                    hide_index=True
                )

                # 候选方案对比
                st.subheader("📊 候选长度性能对比（前5名）")
                top_results = results[:5]
                candidate_data = [
                    [i + 1, res['stock_length'], res['total_utilization'],
                     res['total_stock_used'], res['total_waste']]
                    for i, res in enumerate(top_results)
                ]
                st.dataframe(
                    pd.DataFrame(candidate_data,
                                 columns=["排名", "原始长度(mm)", "总利用率", "材料用量(根)", "总废料(mm)"]),
                    use_container_width=True,
                    hide_index=True
                )

                # 裁切方案详情
                st.subheader("📋 最佳长度的裁切方案详情")
                st.dataframe(
                    pd.DataFrame(best_result['cutting_plans']),
                    use_container_width=True,
                    hide_index=True
                )

        else:
            with st.spinner(f"正在计算指定长度（{stock_length}mm）的方案..."):
                result, error = greedy_cutting_optimization(stock_length, demands)

            if error:
                st.error(f"计算失败：{error}")
            else:
                st.divider()
                st.subheader("## 计算结果：指定长度的裁切方案")

                # 关键指标
                st.dataframe(
                    pd.DataFrame([
                        ["原始材料长度", f"{stock_length} mm"],
                        ["总材料用量", f"{result['total_stock_used']} 根"],
                        ["总利用率", result['total_utilization']],
                        ["总废料长度", f"{result['total_waste']} mm"],
                        ["方案数量", f"{len(result['cutting_plans'])} 种"]
                    ], columns=["指标", "数值"]),
                    use_container_width=True,
                    hide_index=True
                )

                # 裁切方案详情
                st.subheader("📋 裁切方案详情")
                st.dataframe(
                    pd.DataFrame(result['cutting_plans']),
                    use_container_width=True,
                    hide_index=True
                )

    # 重新计算按钮（可选）
    if st.button("重新开始"):
        st.experimental_rerun()


if __name__ == "__main__":
    main()
