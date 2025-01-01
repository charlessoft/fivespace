import pandas as pd
from mailmerge import MailMerge

def mail_merge(excel_path,doc_path):
    lst = []
    df = pd.read_excel(excel_path, sheet_name='1').fillna('')
    for index, sheet_row in df.iterrows():
        template = doc_path
        document = MailMerge(template)
        # 获取模板中的合并字段
        merge_fields = document.get_merge_fields()

        # new_doc = Document(doc_path)  # 从原始模板创建
        student_name = sheet_row['姓名']
        uid = sheet_row['uid']

        sheet_row = sheet_row.astype(str)
        data = {field: sheet_row[field] for field in merge_fields if field in sheet_row}
        # print(data)
        # 执行合并
        document.merge(**data)

        # 保存合并后的文档
        output_filename = f"./output/{uid}.docx"
        document.write(output_filename)
        lst.append(output_filename)
        print(index)
    return lst
if __name__ == '__main__':
    excel_path = "./data/全科成绩及五育评价数据录入表.xls"
    template = "./data/素质评价单（两栏装订）.docx"
    mail_merge(excel_path,template)
    # df = pd.read_excel(excel_path, sheet_name='1').fillna('')
    # for index, sheet_row in df.iterrows():
    #
    #     document = MailMerge(template)
    #     # 获取模板中的合并字段
    #     merge_fields = document.get_merge_fields()
    #
    #     # new_doc = Document(doc_path)  # 从原始模板创建
    #     student_name = sheet_row['姓名']
    #     # if student_name == '庄欣菲':
    #     #     continue
    #     sheet_row = sheet_row.astype(str)
    #     data = {field: sheet_row[field] for field in merge_fields if field in sheet_row}
    #     # print(data)
    #     # 执行合并
    #     document.merge(**data)
    #
    #     # 保存合并后的文档
    #     output_filename = f"./output/merged_document_{index}.docx"
    #     document.write(output_filename)
    #     print(index)
    #
