# -*- coding=UTF-8 -*-
import os
import re
import xlsxwriter
from datetime import datetime
import codecs
import chardet


def parser_file(list_file):
    for txt in list_file:
        parser_txt_content(txt)


def parser_txt_content(file_name):
    """
    解析为对应的步骤
    :param file_name: 
    :return: 
    """
    with open(file_name, 'rb') as file:
        try:
            pattern = re.compile("^\d{0,4}\.?\d{1,3}[、.].*$", flags=re.M)
            raw=file.read()
            if raw.startswith(codecs.BOM_UTF8):
                encoding='utf-8-sig'
            else:
                result=chardet.detect(raw)
                encoding = result['encoding']
            file.close()
            with open(file_name, 'r',encoding=encoding) as truth_file:
                m = pattern.findall(truth_file.read())
                dict_pattern = {}
                index = 1
                for temp in m:
                    dict_pattern[str(index)] = temp
                    index += 1
                # print(dict_pattern)
                # calc_dict_level(dict_pattern)
                return dict_pattern
        except UnicodeDecodeError:
            print("error file="+file_name)
            return file_name


def pattern_dict_privacy(pattern_dict):
    """
    生成每个步骤所指向的上一个步骤的索引
    :param pattern_dict: 
    :return: 返回一个dict
    """
    pattern = re.compile("(权利要求|权利要|权利)\d+\、?\d?(或者)?或?到?与?和?至?\d{0,2}中?(任一项)?(任意一项)?的?所述", flags=re.M)
    pattern_index = re.compile("\d")
    calc_dict = {}
    for key in pattern_dict:
        m = pattern.search(pattern_dict[key])
        if m is not None:
            index_list = pattern_index.findall(m.group())
            if len(index_list) > 1:
                calc_dict[key] = index_list[1]
            else:
                calc_dict[key] = index_list[0]
        else:
            calc_dict[key] = None
    # print(calc_dict)
    # calc_level(calc_dict)
    return calc_dict


def list_files(path):
    list_dir = os.listdir(path)
    list_file = []
    for file in list_dir:
        if file.endswith(".TXT") or file.endswith(".txt"):
            list_file.append(os.path.join(path, file))
    print('dir calc file num is %d' % len(list_file))
    # print(list_file)
    return list_file


def calc_self_dict(dict_self):
    """
    计算层级数
    :param dict_self: 
    :return:返回层级数
    """
    index_num = 1
    for dict_index in dict_self:
        if type(dict_self[dict_index]) == dict:
            index_num += calc_self_dict(dict_self[dict_index])
            # else:
            #     print("no dict,content: %s type is: %s" % (dict_index, str(type(dict_index))))
    return index_num


def generate_self_dict(dict_total, judge_name):
    index = 0
    for value in dict_total:
        if dict_total[value] == judge_name:
            index = 1
            index += generate_self_dict(dict_total, value)
            # else:
            #     index = 0
    return index


def calc_level(dict_to_generate):
    """
    计算主权层级
    :param dict_to_generate: 
    :return: 返回主权层级
    """
    list_key = []
    print(dict_to_generate)
    for key in dict_to_generate:
        if dict_to_generate[key] is None:
            list_key.append(key)
    container_sum = []
    print(list_key)
    for i in list_key:
        container_sum.append(int(generate_self_dict(dict_to_generate, i) + 1))
    return max(container_sum), len(list_key)


def main():
    file_path = input("输入解析文件的目录路径: ")
    list_file = list_files(file_path)
    start_time = datetime.now()
    work_book = xlsxwriter.Workbook('主权层级.xlsx')
    work_sheet = work_book.add_worksheet()
    work_sheet.set_column('A:A', 40)
    work_sheet.set_column('B:A', 30)

    style_bold = work_book.add_format({'bold': True, 'font_size': 16})
    style_normal = work_book.add_format({'font_size': 14})
    style_normal.set_align('center')
    style_normal.set_align('vcenter')
    style_bold.set_align('center')
    style_bold.set_align('vcenter')
    work_sheet.write(0, 0, "文件名字", style_bold)
    work_sheet.write(0, 1, "层数等级", style_bold)
    index = 1
    for txt in list_file:
        dict_step = parser_txt_content(txt)
        dict_index = pattern_dict_privacy(dict_step)
        level = calc_level(dict_index)
        print(level)
        work_sheet.write(index, 0, os.path.basename(txt), style_normal)
        work_sheet.write(index, 1, level[0], style_normal)
        index += 1

    end_time = datetime.now()

    work_sheet.write(index + 1, 3, "处理耗时：%s" % str(end_time - start_time), style_normal)
    work_book.close()


if __name__ == '__main__':
    main()
