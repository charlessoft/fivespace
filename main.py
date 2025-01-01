import sys

import pandas as pd
import requests
from PIL import Image
from io import BytesIO
from docxtpl import DocxTemplate


def generate_randa(user_name, data):
    # 定义请求的 URL
    url = 'https://fivespace.vercel.app/api/chart'  # 替换为你的服务器地址

    # 定义请求的 JSON 数据
    # data = {
    #     "width": 1024,  # 指定图表宽度
    #     "height": 768,  # 指定图表高度
    #     "option": {
    #         "title": {
    #             "text": "五维评价"
    #         },
    #         "legend": {
    #             "data": ["张三", "平均分"]
    #         },
    #         "radar": {
    #             "indicator": [
    #                 {"name": "爱国情怀", "max": 6500},
    #                 {"name": "正直诚信", "max": 16000},
    #                 {"name": "遵规守纪", "max": 30000},
    #                 {"name": "仁爱友善", "max": 38000},
    #                 {"name": "包容理解", "max": 52000}
    #             ]
    #         },
    #         "series": [
    #             {
    #                 "name": "Budget vs spending",
    #                 "type": "radar",
    #                 "data": [
    #                     {
    #                         "value": [4200, 3000, 20000, 35000, 50000],
    #                         "name": "张三"
    #                     },
    #                     {
    #                         "value": [5000, 14000, 28000, 26000, 42000],
    #                         "name": "平均分"
    #                     }
    #                 ]
    #             }
    #         ]
    #     }
    # }

    # 发送 POST 请求
    response = requests.post(url, json=data)

    # 检查请求是否成功
    if response.status_code == 200:
        # 将响应的二进制数据转换为图像
        image = Image.open(BytesIO(response.content))
        # 显示图像
        # image.show()
        # 或者保存图像到本地文件
        image.save(f'./pic/{user_name}.png')

    else:
        print(f"请求失败，状态码: {response.status_code}")


def generate_randa1(file_path):
    # 读取整个工作表
    df = pd.read_excel(file_path, sheet_name='1')
    df_cleaned = df.dropna(subset=['姓名'])  # 去掉 '姓名' 列中为空的行

    # 定义每个维度的列名
    columns_map = {
        "正": ["爱国情怀", "正直诚信", "遵规守纪", "仁爱友善", "包容理解"],
        "善": ["学习态度", "学业进步", "学科成绩", "阅读习惯", "科学探索"],
        "健": ["体质体能", "运动技能", "体锻习惯", "健康生活", "课堂表现"],
        "美": ["文明礼仪", "文化修养", "审美情趣", "音乐成绩", "美术成绩"],
        "德": ["劳动意识", "劳动技能", "劳动实践", "课程表现", "劳动创新"]
    }

    # 计算每个维度的平均分
    average_scores_map = {}
    for dimension, columns in columns_map.items():
        average_scores_map[dimension] = df_cleaned[columns].mean().tolist()

    # 遍历每个学生，生成对应的雷达图数据
    for index, row in df_cleaned.iterrows():
        student_name = row['姓名']
        student_scores_map = {dimension: row[columns].tolist() for dimension, columns in columns_map.items()}

        for dimension, student_scores in student_scores_map.items():
            average_scores = average_scores_map[dimension]
            print(f"{student_name} - {dimension}维度分数: {student_scores}")
            print(f"{dimension}维度平均分: {average_scores}")

    #
    #
    # # 计算每个维度的平均分
    # average_scores_map = {}
    # for dimension, columns in columns_map.items():
    #     average_scores_map[dimension] = df_cleaned[columns].mean().tolist()
    #
    # # 创建一个新的 DataFrame 来存储平均分
    # average_scores = pd.DataFrame()
    # for line in columns_map:
    #     for key, columns in line.items():
    #         for column in columns:
    #             average_scores[column + '平均分'] = [df[column].mean()]
    #
    # # 计算每列的平均分
    # # for column in columns_to_calculate:
    # #     average_scores[column + '平均分'] = [df[column].mean()]
    #
    # # 计算每个维度的平均分（或其他需要的计算）
    # # average_scores = columns_to_calculate.mean(axis=1)
    #
    # for index, row in df.iterrows():
    #     student_name = row['姓名']
    #     student_scores = row.iloc[1:].tolist()
    #     print(student_name, student_scores)
    #     continue
    #
    #     # 准备雷达图数据结构
    #     data = {
    #         "width": 300,  # 指定图表宽度
    #         "height": 300,  # 指定图表高度
    #         "option": {
    #             "title": {
    #                 "text": f"{student_name}的五维评价"
    #             },
    #             "legend": {
    #                 "data": [student_name, "平均分"]
    #             },
    #             "radar": {
    #                 "indicator": [
    #                     {"name": "爱国情怀", "max": 10},
    #                     {"name": "正直诚信", "max": 10},
    #                     {"name": "遵规守纪", "max": 10},
    #                     {"name": "仁爱友善", "max": 10},
    #                     {"name": "包容理解", "max": 10}
    #                 ]
    #             },
    #             "series": [
    #                 {
    #                     "name": "个人 vs 平均",
    #                     "type": "radar",
    #                     "data": [
    #                         {
    #                             "value": student_scores,
    #                             "name": student_name
    #                         },
    #                         {
    #                             "value": average_scores,
    #                             "name": "平均分"
    #                         }
    #                     ]
    #                 }
    #             ]
    #         }
    #     }
    #     # print(data)
    #     generate_randa(student_name, data)


def parse_excel(file_path):
    # 读取指定工作表和列
    df = pd.read_excel(file_path, sheet_name='1', usecols='B,W:AA')
    # 去掉 B 列中为空或者 NaN 的行
    df_cleaned = df.dropna(subset=['姓名'])  # 使用 'B' 列名来检查空值

    # print(df_cleaned)
    # 计算每个维度的平均分
    average_scores = df_cleaned.iloc[:, 1:].mean().tolist()

    # 遍历每个学生，生成对应的雷达图数据
    for index, row in df_cleaned.iterrows():
        student_name = row['姓名']
        student_scores = row.iloc[1:].tolist()
        print(student_name, student_scores)

        # 准备雷达图数据结构
        data = {
            "width": 300,  # 指定图表宽度
            "height": 300,  # 指定图表高度
            "option": {
                "title": {
                    "text": f"{student_name}的五维评价"
                },
                "legend": {
                    "data": [student_name, "平均分"]
                },
                "radar": {
                    "indicator": [
                        {"name": "爱国情怀", "max": 10},
                        {"name": "正直诚信", "max": 10},
                        {"name": "遵规守纪", "max": 10},
                        {"name": "仁爱友善", "max": 10},
                        {"name": "包容理解", "max": 10}
                    ]
                },
                "series": [
                    {
                        "name": "个人 vs 平均",
                        "type": "radar",
                        "data": [
                            {
                                "value": student_scores,
                                "name": student_name
                            },
                            {
                                "value": average_scores,
                                "name": "平均分"
                            }
                        ]
                    }
                ]
            }
        }
        # print(data)
        generate_randa(student_name, data)
    # sys.exit(0)


from docx import Document
import re


def _replace_paragraphs(doc, replacements, df, sheet_row):
    for paragraph in doc.paragraphs:
        for replace_item in replacements:
            key = f"«{replace_item}»"
            value = sheet_row[replace_item]
            if key in paragraph.text:
                if replace_item == '姓名':
                    print("ss")
                value = sheet_row[replace_item]
                # 如果值是 NaN，则替换为 ''
                if pd.isna(value):
                    value = ''
                else:
                    if pd.api.types.is_numeric_dtype(df[replace_item]):
                        value = str(int(value))  # 用 0 替换数值型 NaN
                    else:
                        value = str(value)  # 用空字符串替换非数值型 NaN
                paragraph.text = paragraph.text.replace(key, value)


def _replace_table(doc, replacements, df, sheet_row):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # print("=====",cell.text)
                # if '学生签字' in cell.text:
                #     print(cell.text)
                for replace_item in replacements:
                    key = f"«{replace_item}»"
                    value = sheet_row[replace_item]
                    # 如果值是 NaN，则替换为 ''
                    if pd.isna(value):
                        value = ''
                    else:
                        if pd.api.types.is_numeric_dtype(df[replace_item]):
                            value = str(int(value))  # 用 0 替换数值型 NaN
                        else:
                            value = str(value)  # 用空字符串替换非数值型 NaN
                    if key in cell.text:
                        if '学生签字' in cell.text:
                            print(cell.text)
                        cell.text = cell.text.replace(key, value)
                # for replace_item in replacements:
                #     key = f"«{replace_item}»"
                #     value = sheet_row[replace_item]


def replace_text_in_docx(excel_path, doc_path):
    # tpl = DocxTemplate(doc_path)
    # print(tpl.get_document_xml())

    df = pd.read_excel(excel_path, sheet_name='1')
    replacements = ["50米×8往返跑", "50米×8往返跑等级", "50米×8往返跑评分", "50米跑", "50米跑等级", "50米跑评分",
                    "A体重", "A身高", "B视力右", "B视力左", "X收缩压", "X脉搏", "X舒张压", "一分钟仰卧起坐",
                    "一分钟仰卧起坐等级", "一分钟仰卧起坐评分", "一分钟跳绳", "一分钟跳绳等级", "一分钟跳绳评分",
                    "书法", "事假", "仁爱友善", "体育", "体质体能", "体锻习惯", "作文", "信息", "健康生活", "劳动",
                    "劳动创新", "劳动实践", "劳动意识", "劳动技能", "包容理解", "听力右", "听力左", "坐位体前屈",
                    "坐位体前屈等级", "坐位体前屈评分", "姓名", "学业进步", "学习态度", "学科成绩", "审美情趣", "总分",
                    "总分等级", "总评", "操行等级", "数学", "文化修养", "文明礼仪", "旷课", "正直诚信", "爱国情怀",
                    "班级职务", "病假", "科学", "科学探索", "综合实践", "美术", "美术成绩", "肺活量", "英语", "评语",
                    "课堂表现", "课程表现", "运动技能", "迟到", "道德与法治", "遵规守纪", "阅读", "阅读习惯", "附加分",
                    "音乐", "音乐成绩", "龋齿"]
    # student_name  = df["姓名"]

    # 遍历 DataFrame 的每一行
    for index, sheet_row in df.iterrows():
        # 为每一行创建一个新的文档对象
        new_doc = Document(doc_path)  # 从原始模板创建
        student_name = sheet_row['姓名']
        for field in replacements:
            print(f"{field}: {sheet_row[field]}")

        _replace_paragraphs(new_doc, replacements, df, sheet_row)

        _replace_table(new_doc, replacements, df, sheet_row)

        # 遍历每一个段落
        # for paragraph in new_doc.paragraphs:
        #     for replace_item in replacements:
        #         key = f"«{replace_item}»"
        #         value = shell_row[replace_item]
        #         # 如果值是 NaN，则替换为 ''
        #         if pd.isna(value):
        #             value = ''
        #         else:
        #             if pd.api.types.is_numeric_dtype(df[replace_item]):
        #                 value = str(int(value))  # 用 0 替换数值型 NaN
        #             else:
        #                 value = ''  # 用空字符串替换非数值型 NaN
        #         if key in paragraph.text:
        #             # 替换文本
        #             paragraph.text = paragraph.text.replace(key, value)

        # for table in new_doc.tables:
        #     for row in table.rows:
        #         for cell in row.cells:
        #             for replace_item in replacements:
        #                 key = f"«{replace_item}»"
        #                 value = shell_row[replace_item]
        #                 # 如果值是 NaN，则替换为 ''
        #                 if pd.isna(value):
        #                     value = ''
        #                 else:
        #                     if pd.api.types.is_numeric_dtype(df[replace_item]):
        #                         value = str(int(value))  # 用 0 替换数值型 NaN
        #                     else:
        #                         value = ''  # 用空字符串替换非数值型 NaN
        #                     # value = str(value)
        #                 # print(cell.text)
        #                 if re.search(key, cell.text):
        #                     cell.text = cell.text.replace(key, value)

        # 保存每一条记录到新的 Word 文件
        # new_output_path = output_path.replace('.docx', f'_{index}.docx')
        new_output_path = f"./output/{student_name}.docx"
        new_doc.save(new_output_path)
        print(f"Saved: {new_output_path}")
        sys.exit(0)


if __name__ == '__main__':
    generate_randa1("./data/全科成绩及五育评价数据录入表1.xls")
    # replace_text_in_docx()
    # 定义需要替换的域和值
    # replacements = {
    #     r'\{field1\}': 'Value1',
    #     r'\{field2\}': 'Value2',
    #     # 添加更多的替换项
    # }

    # 调用函数进行替换
    # replace_text_in_docx("./data/全科成绩及五育评价数据录入表.xls", './data/素质评价单（两栏装订）.docx')
    # # parse_excel("../../../data/观澜/全科成绩及五育评价数据录入表1.xls")
