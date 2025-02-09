import pandas as pd
import streamlit as st
import altair as alt
import os

# 设置页面标题
st.title("题目分析")
st.markdown("<p style='text-align: left; color: red;'>读取数据会很慢，请耐心等待。右上角有个小人在动，就表示正在运行。点击右上角的三点图标，还可以进行一些设置，比如设置为宽屏。</p>", unsafe_allow_html=True)

# 自动读取 output 目录下的所有 xlsx 文件
file_list = [f for f in os.listdir('output') if f.endswith('.xlsx')]

if file_list:
    # 选择要分析的文件
    selected_file = st.selectbox("请选择要统计的作业/考试:", file_list)

    # 读取文件
    df = pd.read_excel(os.path.join('output', selected_file), engine='openpyxl')

    # 处理可能的列名变动
    df.columns = df.columns.str.replace('试题 ', '试题', regex=False)

    # 获取教师和班级信息
    teachers = df['教师'].dropna().unique().tolist()
    classes = df['班级'].dropna().unique().tolist()

    # 选择教师
    selected_teacher = st.selectbox("选择教师:", ["全部"] + teachers)

    # 选择班级（根据教师筛选）
    if selected_teacher != "全部":
        filtered_classes = df[df['教师'] == selected_teacher]['班级'].dropna().unique().tolist()
    else:
        filtered_classes = classes

    selected_class = st.selectbox("选择班级:", ["全部"] + filtered_classes)

    # 根据选择的教师和班级进行筛选
    if selected_teacher != "全部":
        df = df[df['教师'] == selected_teacher]
    if selected_class != "全部":
        df = df[df['班级'] == selected_class]

    results = []
    i = 1

    while True:
        answer_col = f'回答{i}'
        if answer_col not in df.columns:
            break

        answers = df[answer_col].dropna()
        valid_answers = answers[~answers.isin(["-", "- -"])]
        result = valid_answers.value_counts().reset_index()
        result.columns = ['答案', '出现次数']

        # 处理“学生”列，防止 TypeError
        if '姓氏' in df.columns and '名' in df.columns:
            df['学生'] = df['姓氏'].fillna('') + df['名'].fillna('')
        else:
            df['学生'] = "未知"

        # 计算每个答案对应的学生
        def get_students(x):
            students = df.loc[df[answer_col] == x, '学生'].dropna().tolist()
            return ', '.join(students) if students else "无"

        result['学生'] = result['答案'].apply(get_students)

        standard_answer_col = f'标准答案{i}'
        standard_answer = df[standard_answer_col].iloc[0] if standard_answer_col in df.columns else "未知"

        correct_count = (df[answer_col] == standard_answer).sum()
        total_count = df[answer_col].notna().sum() - df[answer_col].isin(["-", "- -"]).sum()
        accuracy = (correct_count / total_count * 100) if total_count > 0 else 0

        question_content = df[f'试题{i}'].iloc[0] if f'试题{i}' in df.columns else "未知"

        results.append({
            '题号': i,
            '试题': question_content,
            '标准答案': standard_answer,
            '答题人数': total_count,
            '正确率': accuracy,
            '答案统计': result[['答案', '出现次数', '学生']],
            '错误答案统计': result[result['答案'] != standard_answer].sort_values(by='出现次数', ascending=False)
        })

        i += 1

    # 排序选项
    sort_option = st.selectbox("选择排序方式:", ["按照题目原本顺序", "按照正确率升序", "按照正确率降序"])

    if sort_option == "按照正确率升序":
        sorted_results = sorted(results, key=lambda x: x['正确率'])
    elif sort_option == "按照正确率降序":
        sorted_results = sorted(results, key=lambda x: x['正确率'], reverse=True)
    else:
        sorted_results = results

    # 侧边栏导航
    st.sidebar.title("题目导航")
    for res in sorted_results:
        st.sidebar.markdown(f"[第{res['题号']}题 (正确率: {res['正确率']:.2f}%)](#{res['题号']})")

    # 显示统计结果
    for res in sorted_results:
        st.markdown(f"<a id='{res['题号']}'></a>", unsafe_allow_html=True)
        st.subheader(f"第{res['题号']}题")
        st.write(f"题目: {res['试题']}")
        st.write(f"标准答案: {res['标准答案']}")
        st.write(f"答题人数: {res['答题人数']}")
        st.write(f"正确率: {res['正确率']:.2f}%")

        if not res['错误答案统计'].empty:
            st.write("#### 错误答案统计")
            error_stats = res['错误答案统计']

            bar_chart = alt.Chart(error_stats).mark_bar(color='red').encode(
                y=alt.Y('答案', sort='-x'),
                x='出现次数',
                tooltip=['答案', '出现次数', '学生']
            )

            st.altair_chart(bar_chart, use_container_width=True)

            for _, row in error_stats.iterrows():
                color = 'green' if row['答案'] == res['标准答案'] else 'red'
                st.markdown(f"<div style='color:black;'>答案: <span style='color:{color};'>{row['答案']}</span></div>", unsafe_allow_html=True)
                st.write(f"出现次数: {row['出现次数']}")
                st.write(f"学生: {row['学生']}")
                st.write("")

    st.success("统计完成！")

else:
    st.error("未找到任何 xlsx 文件，请检查 output 目录。")
