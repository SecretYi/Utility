# -*- coding:utf-8 -*-
from pyexcel_xls import get_data

format_key = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14']
dict_country_file = dict()
dict_file_save_start = dict()


class DefineException(Exception):
    pass


def create_file(index, name):
    dict_country_file[index] = name


def save_file(key_index, value):
    global dict_file_save_start
    try:
        name = dict_country_file[column_index] + ".txt"
        # 如果已经进行了第一次写入,就进行追加
        if dict_file_save_start.get(column_index, False):
            with open(name, 'a', encoding='utf-8') as f:
                f.write('<string name={}>"{}"</string>'.format(format_key[key_index], value))
                f.write('\n')
        # 开始写入
        else:
            dict_file_save_start[column_index] = True
            with open(name, 'w', encoding='utf-8') as f:
                f.write('<string name={}>"{}"</string>'.format(format_key[key_index], value))
                f.write('\n')
    except IndexError:
        print(key_index)
        pass
    except KeyError:
        pass


if __name__ == '__main__':
    file_name = r'Filto安卓1.5 文案 (3).xlsx'
    data = get_data(file_name)
    start_record = False
    try:
        # pyexcel_xls加载xls文件是一个字典,sheet是key,内容是value
        for sheet, content in data.items():
            for row in content:
                for column_index in range(len(row)):
                    # 通过第一行进行多国文件名获取
                    if row[column_index] == '位置描述':
                        start_record = True
                    elif start_record:
                        # 获取的多国文件名进行缓存
                        create_file(column_index, row[column_index])
                if start_record:
                    raise DefineException()
    except DefineException:
        pass
    for sheet, content in data.items():
        for row_index in range(len(content)):
            row = content[row_index]
            if len(row) > 0 and str(row[0]).isdigit():
                for column_index in range(len(content[row_index])):
                    if column_index in dict_country_file.keys():
                        # 将内容以安卓格式写入文件
                        save_file(row[0] - 1, row[column_index])
