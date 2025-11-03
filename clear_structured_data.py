import pandas as pd
import numpy as np
import os
import re

def process_survey_data(input_file, output_file='structured_data.xlsx'):
    """
    处理问卷数据，将其转换为结构化的数据表

    参数:
        input_file: 输入的Excel文件路径
        output_file: 输出的Excel文件路径
    """

    # 读取原始数据
    print("正在读取数据...")
    df = pd.read_excel(input_file)

    # 创建新的数据框
    processed_data = pd.DataFrame()

    # 1. 序号 (ID)
    processed_data['ID'] = range(1, len(df) + 1)

    # 2. 性别 (gender): 0=男, 1=女
    # 假设原始数据中"男"和"女"在某一列
    gender_col = df.iloc[:, 1]  # 根据图片，性别在第2列
    processed_data['gender'] = gender_col.map({'男': 0, '女': 1})

    # 3. 年龄分层 (age_cat)
    age_col = df.iloc[:, 2]  # 年龄在第3列
    age_mapping = {
        '35岁及以下': 1,
        # '小学及以下': 1,  # 可能的变体
        '36-45岁': 2,
        '46-55岁': 3,
        '56-65岁': 4,
        '66岁及以上': 5
    }
    processed_data['age_cat'] = age_col.map(age_mapping)

    # 4. 受教育程度 (edu)
    edu_col = df.iloc[:, 3]  # 教育程度在第4列
    edu_mapping = {
        '小学及以下': 1,
        '初中/中专': 2,
        '高中': 3,
        '大专': 4,
        '本科': 5
    }
    processed_data['edu'] = edu_col.map(edu_mapping)

    # 5. 家庭人口相关变量 (问题4)
    processed_data['f_size'] = pd.to_numeric(df.iloc[:, 4], errors='coerce')  # 家庭总人口
    processed_data['up15_size'] = pd.to_numeric(df.iloc[:, 5], errors='coerce')  # 15周岁以上人口
    processed_data['l_size'] = pd.to_numeric(df.iloc[:, 6], errors='coerce')  # 劳动人口
    processed_data['migrant'] = pd.to_numeric(df.iloc[:, 7], errors='coerce')  # 外出务工人数

    # 6. 家庭年总收入 (问题5)
    processed_data['income'] = pd.to_numeric(df.iloc[:, 8], errors='coerce')

    # 生成收入的自然对数
    # processed_data['ln_income'] = np.log(processed_data['income'].replace(0, np.nan))

    # 7. 是否参与农文旅 (问题6) - 决策变量
    participate_col = df.iloc[:, 9]
    processed_data['participate'] = participate_col.map({'是': 1, '否': 0})

    # 8. 处理问题7 - 从事农文旅的方式 (多选题，需要拆分)
    # 这部分需要根据实际列数和内容调整

    # 9. 农文旅收入 (问题8)
    processed_data['agri_income'] = pd.to_numeric(df.iloc[:, 11], errors='coerce').replace('(跳过)', '').fillna(0)

    # 10. 分红收入 (问题9)
    processed_data['dividend'] = pd.to_numeric(df.iloc[:, 12], errors='coerce').replace('(跳过)', '').fillna(0)

    # 11. 处理问题10 - 未从事农文旅的原因 (多选题)
    # 根据实际情况添加虚拟变量

    # 12. 技能培训 (问题11)
    training_col = df.iloc[:, 14] if len(df.columns) > 14 else None
    if training_col is not None:
        training_mapping = {
            '是，政府组织': 1,
            '是，企业培训': 2,
            '是，在学校学习过': 3,
            '否': 4
        }
        processed_data['training'] = training_col.map(training_mapping)
        # 生成二分变量
        processed_data['training_yes'] = processed_data['training'].apply(
            lambda x: 1 if x in [1, 2, 3] else 0 if x == 4 else np.nan
        )

    # 13. 处理问题12 - 淡季工作 (多选题)
    # 根据实际情况添加虚拟变量

    # 14. 耕地面积分层 (问题13)
    land_col = df.iloc[:, 16] if len(df.columns) > 16 else None
    if land_col is not None:
        land_mapping = {
            '无': 0,
            '1-5亩': 1,
            '6-10亩': 2,
            '11-15亩': 3,
            '16-20亩': 4,
            '21亩及以上': 5
        }
        processed_data['land_cat'] = land_col.map(land_mapping)

    # 15. 交通通畅程度 (问题14)
    transport_col = df.iloc[:, 17] if len(df.columns) > 17 else None
    if transport_col is not None:
        likert_mapping = {
            '极差': 1,
            '较差': 2,
            '一般': 3,
            '较高': 4,
            '非常完善': 5,
            '极弱': 1,
            '较弱': 2,
            '较强': 4,
            '极强': 5
        }
        processed_data['transport'] = transport_col.map(likert_mapping)

    # 16. 政策扶持力度 (问题15) - 工具变量
    policy_col = df.iloc[:, 18] if len(df.columns) > 18 else None
    if policy_col is not None:
        processed_data['policy'] = policy_col.map(likert_mapping)

    # 17. 信息化建设程度 (问题16)
    info_col = df.iloc[:, 19] if len(df.columns) > 19 else None
    if info_col is not None:
        processed_data['info'] = info_col.map(likert_mapping)

    # 18. 旅游吸引力 (问题17)
    attraction_col = df.iloc[:, 20] if len(df.columns) > 20 else None
    if attraction_col is not None:
        processed_data['attraction'] = attraction_col.map(likert_mapping)

    # 19. 环境卫生条件 (问题18)
    env_col = df.iloc[:, 21] if len(df.columns) > 21 else None
    if env_col is not None:
        env_mapping = {
            '完全不适合': 1,
            '适合但需要改进': 2,
            '适合需要加大投入建设': 3,
            '适合': 4,
            '非常适合': 5
        }
        processed_data['env'] = env_col.map(env_mapping)

    # 20. 处理问题19 - 主要问题 (多选题)
    # 根据实际列数添加虚拟变量
    col_data = df.iloc[:, 22].replace('(跳过)', '').fillna('')
    # ---- Step 1: 提取所有不同选项 ----
    all_items = []
    for x in col_data:
        items = [i.strip() for i in str(x).split('┋') if i.strip()]
        all_items.extend(items)
    unique_items = pd.Series(all_items).value_counts().index.tolist()
    # ---- Step 2: 识别“其他（请注明）” ----
    other_pattern = re.compile(r'其他（请注明）〖(.*?)〗')
    other_items = set()
    for x in col_data:
        matches = other_pattern.findall(str(x))
        other_items.update(matches)
    # 把“其他”具体说明加入 unique_items（加上前缀）
    for o in other_items:
        unique_items.append(f'其他_{o}')
    # ---- Step 3: 创建虚拟变量列 ----
    def has_keyword(x, kw):
        if kw.startswith('其他_'):
            real_kw = kw.replace('其他_', '')
            return 1 if f'〖{real_kw}〗' in str(x) else 0
        else:
            return 1 if kw in str(x) else 0
    for kw in unique_items:
        safe_col = re.sub(r'\W+', '_', kw)  # 列名安全化
        processed_data[safe_col] = col_data.apply(lambda x: has_keyword(x, kw))

    # 保存处理后的数据
    print(f"正在保存数据到 {output_file}...")
    processed_data.to_excel(output_file, index=False)

    # 生成数据字典
    data_dict = {
        '变量名': ['ID', 'gender', 'age_cat', 'edu', 'f_size', 'up15_size', 'l_size',
                   'migrant', 'income', 'ln_income', 'participate', 'agri_income',
                   'dividend', 'training', 'training_yes', 'land_cat', 'transport',
                   'policy', 'info', 'attraction', 'env'],
        '变量含义': ['样本唯一编号', '受访者性别', '年龄分层', '受教育程度', '家庭总人口',
                     '15周岁以上人口数', '家庭劳动人口数', '常年外出务工人数',
                     '家庭年总收入(万元)', '收入的自然对数', '决策变量(处理组)',
                     '农文旅收入(万元)', '分红收入(万元)', '技能培训', '是否培训(二分)',
                     '耕地面积分层', '交通通畅程度', '政策扶持力度', '信息化建设程度',
                     '旅游吸引力', '环境卫生条件'],
        '编码说明': ['1,2,3,...', '0=男; 1=女', '1=35岁及以下; 2=36-45岁; 3=46-55岁; 4=56-65岁; 5=66岁及以上',
                     '1=小学及以下; 2=初中/中专; 3=高中; 4=大专; 5=本科', '数值',
                     '数值', '数值', '数值', '数值', 'ln(income)', '1=是; 0=否',
                     '数值', '数值', '1=政府组织; 2=企业培训; 3=否', '1=是; 0=否',
                     '0=无; 1=1-5亩; 2=6-10亩; 3=11-15亩; 4=16-20亩; 5=21亩及以上',
                     '1=极差; 2=较差; 3=一般; 4=较高; 5=非常完善',
                     '1=极弱; 2=较弱; 3=一般; 4=较强; 5=极强',
                     '1=极差; 2=较差; 3=一般; 4=较高; 5=非常完善',
                     '1=极弱; 2=较弱; 3=一般; 4=较强; 5=极强',
                     '1=完全不适合; 2=适合但需要改进; 3=适合需要加大投入建设; 4=适合; 5=非常适合']
    }

    dict_df = pd.DataFrame(data_dict)
    dict_output = output_file.replace('.xlsx', '_数据字典.xlsx')
    dict_df.to_excel(dict_output, index=False)

    # 输出描述性统计
    print("\n数据处理完成！")
    print(f"处理后的数据已保存到: {output_file}")
    print(f"数据字典已保存到: {dict_output}")
    print(f"\n总样本量: {len(processed_data)}")
    print(f"总变量数: {len(processed_data.columns)}")

    # 显示基本统计信息
    print("\n基本描述性统计:")
    print(processed_data.describe())

    return processed_data

if __name__ == "__main__":
    input_file = "农文旅融合对农户增收的影响研究问卷(1).xls" 

    if os.path.exists(input_file):
        processed_df = process_survey_data(input_file)
    else:
        print(f"错误: 找不到文件 {input_file}")
        print("请将代码中的 'your_survey_data.xlsx' 替换为实际文件路径")