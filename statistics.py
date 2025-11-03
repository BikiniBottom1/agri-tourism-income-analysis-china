import pandas as pd
import numpy as np
from scipy import stats


def comprehensive_descriptive_stats(input_file, output_file='comprehensive_descriptive_stats.xlsx'):
    """
    生成完整的描述性统计分析（包括分类变量和连续变量）

    参数:
        input_file: 处理后的结构化数据文件路径
        output_file: 输出的统计结果文件路径
    """

    # 读取数据
    print("正在读取数据...")
    df = pd.read_excel(input_file)

    total_n = len(df)
    print(f"样本总数: {total_n}")

    # 创建Excel写入器
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        # ============= 1. 样本规模 =============
        print("\n生成样本规模统计...")
        sample_size = pd.DataFrame({
            '统计项': ['样本总数'],
            '数量(n)': [total_n]
        })
        sample_size.to_excel(writer, sheet_name='1_样本规模', index=False)

        # ============= 2. 个体特征 =============
        print("生成个体特征统计...")

        # 2.1 性别
        gender_stats = df['gender'].value_counts().sort_index()
        gender_table = pd.DataFrame({
            '类别': ['男', '女', '合计'],
            '频数': [
                gender_stats.get(0, 0),
                gender_stats.get(1, 0),
                gender_stats.get(0, 0) + gender_stats.get(1, 0)
            ],
            '百分比(%)': [
                f"{(gender_stats.get(0, 0) / total_n * 100):.2f}",
                f"{(gender_stats.get(1, 0) / total_n * 100):.2f}",
                "100.00"
            ]
        })

        # 2.2 年龄分层
        age_labels = ['35岁及以下', '36-45岁', '46-55岁', '56-65岁', '66岁及以上']
        age_stats = df['age_cat'].value_counts().sort_index()
        age_data = []
        cumsum = 0
        for i, label in enumerate(age_labels, 1):
            count = age_stats.get(i, 0)
            pct = (count / total_n * 100) if total_n > 0 else 0
            cumsum += pct
            age_data.append({
                '类别': label,
                '频数': count,
                '百分比(%)': f"{pct:.2f}",
                '累计百分比(%)': f"{cumsum:.2f}"
            })
        age_data.append({'类别': '合计', '频数': total_n, '百分比(%)': '100.00', '累计百分比(%)': '100.00'})
        age_table = pd.DataFrame(age_data)

        # 2.3 教育程度
        edu_labels = ['小学及以下', '初中/中专', '高中', '大专', '本科']
        edu_stats = df['edu'].value_counts().sort_index()
        edu_data = []
        cumsum = 0
        for i, label in enumerate(edu_labels, 1):
            count = edu_stats.get(i, 0)
            pct = (count / total_n * 100) if total_n > 0 else 0
            cumsum += pct
            edu_data.append({
                '类别': label,
                '编码': i,
                '频数': count,
                '百分比(%)': f"{pct:.2f}",
                '累计百分比(%)': f"{cumsum:.2f}"
            })
        edu_data.append({'类别': '合计', '编码': '', '频数': total_n, '百分比(%)': '100.00', '累计百分比(%)': '100.00'})
        edu_table = pd.DataFrame(edu_data)

        # 教育程度均值
        edu_mean = df['edu'].mean()
        edu_std = df['edu'].std()
        edu_summary = pd.DataFrame({
            '统计量': ['均值', '标准差'],
            '数值': [f"{edu_mean:.2f}", f"{edu_std:.2f}"]
        })

        # 保存个体特征
        gender_table.to_excel(writer, sheet_name='2_个体特征_性别', index=False)
        age_table.to_excel(writer, sheet_name='2_个体特征_年龄', index=False)
        edu_table.to_excel(writer, sheet_name='2_个体特征_教育', index=False)

        # 在同一sheet添加教育程度均值
        workbook = writer.book
        worksheet = writer.sheets['2_个体特征_教育']
        startrow = len(edu_table) + 3
        edu_summary.to_excel(writer, sheet_name='2_个体特征_教育',
                             startrow=startrow, index=False)

        # ============= 3. 家庭结构特征 =============
        print("生成家庭结构统计...")

        household_vars = {
            'f_size': '家庭总人口',
            'up15_size': '15周岁以上人口数',
            'l_size': '劳动力人口数',
            'migrant': '外出务工人数'
        }

        household_data = []
        for var, name in household_vars.items():
            if var in df.columns:
                household_data.append({
                    '变量': name,
                    '均值': f"{df[var].mean():.2f}",
                    '标准差': f"{df[var].std():.2f}",
                    '最小值': f"{df[var].min():.0f}",
                    '最大值': f"{df[var].max():.0f}",
                    '有效样本数': df[var].notna().sum()
                })

        household_table = pd.DataFrame(household_data)
        household_table.to_excel(writer, sheet_name='3_家庭结构特征', index=False)

        # ============= 4. 经济特征 =============
        print("生成经济特征统计...")

        # 4.1 家庭年收入基本统计
        income_basic = pd.DataFrame({
            '统计量': ['均值', '标准差', '最小值', '最大值', '中位数', '有效样本数'],
            '家庭年收入(万元)': [
                f"{df['income'].mean():.2f}",
                f"{df['income'].std():.2f}",
                f"{df['income'].min():.2f}",
                f"{df['income'].max():.2f}",
                f"{df['income'].median():.2f}",
                df['income'].notna().sum()
            ]
        })

        # 4.2 收入对数的统计
        ln_income_basic = pd.DataFrame({
            '统计量': ['均值', '标准差', '最小值', '最大值', '有效样本数'],
            'ln(收入)': [
                f"{df['ln_income'].mean():.4f}",
                f"{df['ln_income'].std():.4f}",
                f"{df['ln_income'].min():.4f}",
                f"{df['ln_income'].max():.4f}",
                df['ln_income'].notna().sum()
            ]
        })

        # 4.3 按参与状态分组的收入对比
        participate_income = df.groupby('participate')['ln_income'].agg(['mean', 'std', 'count'])
        participate_income_table = pd.DataFrame({
            '组别': ['未参与农文旅(0)', '参与农文旅(1)'],
            '均值': [f"{participate_income.loc[0, 'mean']:.4f}",
                     f"{participate_income.loc[1, 'mean']:.4f}"],
            '标准差': [f"{participate_income.loc[0, 'std']:.4f}",
                       f"{participate_income.loc[1, 'std']:.4f}"],
            '样本数': [int(participate_income.loc[0, 'count']),
                       int(participate_income.loc[1, 'count'])]
        })

        # 4.4 t检验
        group0 = df[df['participate'] == 0]['ln_income'].dropna()
        group1 = df[df['participate'] == 1]['ln_income'].dropna()
        t_stat, p_value = stats.ttest_ind(group0, group1)

        ttest_result = pd.DataFrame({
            '检验项': ['t统计量', 'p值', '显著性'],
            '数值': [f"{t_stat:.4f}", f"{p_value:.4f}",
                     '***' if p_value < 0.01 else '**' if p_value < 0.05 else '*' if p_value < 0.1 else '不显著']
        })

        # 保存经济特征
        income_basic.to_excel(writer, sheet_name='4_经济特征_收入', index=False)

        worksheet = writer.sheets['4_经济特征_收入']
        ln_income_basic.to_excel(writer, sheet_name='4_经济特征_收入',
                                 startrow=len(income_basic) + 2, index=False)
        participate_income_table.to_excel(writer, sheet_name='4_经济特征_收入',
                                          startrow=len(income_basic) + len(ln_income_basic) + 5, index=False)
        ttest_result.to_excel(writer, sheet_name='4_经济特征_收入',
                              startrow=len(income_basic) + len(ln_income_basic) + len(participate_income_table) + 8,
                              index=False)

        # ============= 5. 产业参与特征 =============
        print("生成产业参与统计...")

        participate_stats = df['participate'].value_counts().sort_index()
        participate_table = pd.DataFrame({
            '类别': ['未参与', '参与', '合计'],
            '频数': [
                participate_stats.get(0, 0),
                participate_stats.get(1, 0),
                total_n
            ],
            '百分比(%)': [
                f"{(participate_stats.get(0, 0) / total_n * 100):.2f}",
                f"{(participate_stats.get(1, 0) / total_n * 100):.2f}",
                "100.00"
            ]
        })
        participate_table.to_excel(writer, sheet_name='5_产业参与特征', index=False)

        # ============= 6. 主观感知变量 =============
        print("生成主观感知统计...")

        perception_vars = {
            'transport': '交通通畅程度',
            'info': '信息化建设程度',
            'attraction': '旅游吸引力',
            'env': '环境卫生条件'
        }

        perception_data = []
        for var, name in perception_vars.items():
            if var in df.columns:
                perception_data.append({
                    '变量': name,
                    '均值': f"{df[var].mean():.2f}",
                    '标准差': f"{df[var].std():.2f}",
                    '最小值': int(df[var].min()),
                    '最大值': int(df[var].max()),
                    '中位数': f"{df[var].median():.2f}",
                    '有效样本数': df[var].notna().sum()
                })

        perception_table = pd.DataFrame(perception_data)
        perception_table.to_excel(writer, sheet_name='6_主观感知', index=False)

        # ============= 7. 政策支持 =============
        print("生成政策支持统计...")

        if 'policy' in df.columns:
            policy_table = pd.DataFrame({
                '变量': ['政策扶持力度'],
                '均值': [f"{df['policy'].mean():.2f}"],
                '标准差': [f"{df['policy'].std():.2f}"],
                '最小值': [int(df['policy'].min())],
                '最大值': [int(df['policy'].max())],
                '中位数': [f"{df['policy'].median():.2f}"],
                '有效样本数': [df['policy'].notna().sum()]
            })
            policy_table.to_excel(writer, sheet_name='7_政策支持', index=False)

        # ============= 8. 培训情况 =============
        print("生成培训统计...")

        if 'training_yes' in df.columns:
            training_stats = df['training_yes'].value_counts().sort_index()
            training_table = pd.DataFrame({
                '类别': ['未参加培训', '参加培训', '合计'],
                '频数': [
                    training_stats.get(0, 0),
                    training_stats.get(1, 0),
                    total_n
                ],
                '百分比(%)': [
                    f"{(training_stats.get(0, 0) / total_n * 100):.2f}",
                    f"{(training_stats.get(1, 0) / total_n * 100):.2f}",
                    "100.00"
                ]
            })
            training_table.to_excel(writer, sheet_name='8_培训情况', index=False)

        # ============= 9. 综合汇总表 =============
        print("生成综合汇总表...")

        summary_data = []

        # 样本规模
        summary_data.append({
            '类别': '样本规模',
            '变量': 'n',
            '统计结果': str(total_n),
            '说明': '样本总数'
        })

        # 个体特征
        summary_data.append({
            '类别': '个体特征',
            '变量': 'gender',
            '统计结果': f"男: {gender_stats.get(0, 0)} ({(gender_stats.get(0, 0) / total_n * 100):.2f}%); 女: {gender_stats.get(1, 0)} ({(gender_stats.get(1, 0) / total_n * 100):.2f}%)",
            '说明': '性别分布'
        })

        summary_data.append({
            '类别': '个体特征',
            '变量': 'age_cat',
            '统计结果': f"详见年龄分层表",
            '说明': '年龄分层'
        })

        summary_data.append({
            '类别': '个体特征',
            '变量': 'edu',
            '统计结果': f"{edu_mean:.2f} ± {edu_std:.2f}",
            '说明': '教育程度(均值±标准差)'
        })

        # 家庭结构
        for var, name in household_vars.items():
            if var in df.columns:
                summary_data.append({
                    '类别': '家庭结构',
                    '变量': var,
                    '统计结果': f"{df[var].mean():.2f} ± {df[var].std():.2f}",
                    '说明': name
                })

        # 经济特征
        summary_data.append({
            '类别': '经济特征',
            '变量': 'income',
            '统计结果': f"{df['income'].mean():.2f} ± {df['income'].std():.2f}",
            '说明': '家庭年收入(万元)'
        })

        summary_data.append({
            '类别': '经济特征',
            '变量': 'ln_income',
            '统计结果': f"未参与: {participate_income.loc[0, 'mean']:.4f}; 参与: {participate_income.loc[1, 'mean']:.4f}; t={t_stat:.4f}, p={p_value:.4f}",
            '说明': 'ln(收入)按参与状态对比'
        })

        # 产业参与
        summary_data.append({
            '类别': '产业参与',
            '变量': 'participate',
            '统计结果': f"未参与: {participate_stats.get(0, 0)} ({(participate_stats.get(0, 0) / total_n * 100):.2f}%); 参与: {participate_stats.get(1, 0)} ({(participate_stats.get(1, 0) / total_n * 100):.2f}%)",
            '说明': '是否参与农文旅'
        })

        # 主观感知
        for var, name in perception_vars.items():
            if var in df.columns:
                summary_data.append({
                    '类别': '主观感知',
                    '变量': var,
                    '统计结果': f"{df[var].mean():.2f} ± {df[var].std():.2f}",
                    '说明': name
                })

        # 政策支持
        if 'policy' in df.columns:
            summary_data.append({
                '类别': '政策支持',
                '变量': 'policy',
                '统计结果': f"{df['policy'].mean():.2f} ± {df['policy'].std():.2f}",
                '说明': '政策扶持力度'
            })

        # 培训
        if 'training_yes' in df.columns:
            summary_data.append({
                '类别': '培训',
                '变量': 'training_yes',
                '统计结果': f"未培训: {training_stats.get(0, 0)} ({(training_stats.get(0, 0) / total_n * 100):.2f}%); 培训: {training_stats.get(1, 0)} ({(training_stats.get(1, 0) / total_n * 100):.2f}%)",
                '说明': '是否接受培训'
            })

        summary_table = pd.DataFrame(summary_data)
        summary_table.to_excel(writer, sheet_name='0_综合汇总', index=False)

    print(f"\n完成！所有统计结果已保存到: {output_file}")
    print("\n生成的工作表:")
    print("  - 0_综合汇总: 所有变量的简要统计")
    print("  - 1_样本规模")
    print("  - 2_个体特征_性别/年龄/教育")
    print("  - 3_家庭结构特征")
    print("  - 4_经济特征_收入(含t检验)")
    print("  - 5_产业参与特征")
    print("  - 6_主观感知")
    print("  - 7_政策支持")
    print("  - 8_培训情况")

    return

if __name__ == "__main__":
    input_file = "structured_data.xlsx"  

    print("=" * 70)
    print("生成完整描述性统计分析")
    print("=" * 70)

    comprehensive_descriptive_stats(input_file, 'comprehensive_descriptive_stats.xlsx')

    print("\n" + "=" * 70)
    print("统计分析完成！")
    print("=" * 70)