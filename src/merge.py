import os
import pandas as pd
import datetime


def auto_merge():
    try:
        print(' ')
        print('温馨提示：请把需要合并的表格复制到input文件夹下，再进行以下操作！')
        print(' ')
        start = input('请输入表头所在行号(默认为1):')
        start = 1 if start == '' else int(start)
        end_exclude = input('请输入末尾需要剔除的行数(默认为0)：')
        end_exclude = 0 if end_exclude == '' else int(end_exclude)
        input_path = './input/'
        output_path = './output/'
        data_list = []
        for file in os.listdir(input_path):
            if ".xlsx" in file or ".xls" in file:
                df = pd.read_excel(os.path.join(input_path, file), header=start - 1)
                # 剔除末尾行数
                if end_exclude > 0:
                    df = df[:-end_exclude]
                # 删除所有空值行
                df = df.dropna(axis=0, how='all')
                try:
                    # 去除表头空格
                    df = df.rename(columns=lambda x: x.strip())
                    # 剔除示例数据
                    df = df[~ df[u'学员姓名'].str.contains('张三')]
                    df = df[~ df[u'学员姓名'].str.contains('李四')]
                except Exception as e:
                    print('warning:', e)
                data_list.append(df)
        res = pd.concat(data_list)
        file_name = datetime.datetime.today().strftime('%Y%m%d') + '合并结果.xlsx'
        if not os.path.isdir(output_path):
            os.makedirs(output_path)
        # 输出结果
        res.to_excel(os.path.join(output_path, file_name), index=False)
        # print(res)
        print('合并成功%s个表( success )!' % len(data_list))
    except Exception as e:
        print('合并错误(error)', e)
    print('')
    print('作者：许建阳')
    print('')
    os.system('pause')


auto_merge()
