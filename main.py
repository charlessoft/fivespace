import sys

import pandas as pd
import requests
from PIL import Image
from io import BytesIO



def generate_randa(user_name,data):
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

def parse_excel(file_path):
    # 读取指定工作表和列
    df = pd.read_excel(file_path, sheet_name='1',  usecols='B,W:AA')
    # 去掉 B 列中为空或者 NaN 的行
    df_cleaned = df.dropna(subset=['姓   名'])  # 使用 'B' 列名来检查空值

    # print(df_cleaned)
    # 计算每个维度的平均分
    average_scores = df_cleaned.iloc[:, 1:].mean().tolist()

    # 遍历每个学生，生成对应的雷达图数据
    for index, row in df_cleaned.iterrows():
        student_name = row['姓   名']
        student_scores = row.iloc[1:].tolist()
        print(student_name,student_scores)

        # 准备雷达图数据结构
        data = {
            "width": 1024,  # 指定图表宽度
            "height": 768,  # 指定图表高度
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
        generate_randa(student_name,data)
    # sys.exit(0)



if __name__ == '__main__':
    parse_excel("../../../data/观澜/全科成绩及五育评价数据录入表1.xls")