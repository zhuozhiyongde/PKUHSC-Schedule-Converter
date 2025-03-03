#!/usr/bin/env python
# -*- encoding: utf-8 -*-
# @Author  :   Arthals
# @File    :   converter_fall.py
# @Time    :   2024/08/12 20:59:38
# @Contact :   zhuozhiyongde@126.com
# @Software:   Visual Studio Code


from numpy import place
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import re
import os


def week_to_number(week_str):
    # print(week_str)
    if not re.match(r"周", week_str):
        week_str = "周" + week_str
    week_dict = {
        "周一": 1,
        "周二": 2,
        "周三": 3,
        "周四": 4,
        "周五": 5,
        "周六": 6,
        "周日": 7,
    }
    return week_dict.get(week_str, "输入错误")


def calculate_week(date):
    # 定义起始日期
    start_date = datetime.strptime("2024-08-12", "%Y-%m-%d")
    delta = date - start_date
    return delta.days // 7 + 1


def exact_day(date_str):
    """
    Input:
    8.20/5-7 周二
    -
    8.12（周一）第1-2节
    -
    8.22 四	1-2
    -
    9.20 五 上午
    -
    9.6（周五上午）
    -
    8.14（周三）第3节

    Return:
    星期,开始节数,结束节数,周数

    周数从 2024.8.12 开始计算
    """
    date_str = re.sub("\s+", " ", date_str)

    # 正则表达式模式
    pattern1 = re.compile(r"(\d+\.\d+)/(\d+-\d+)\s*(周[一二三四五六日])")
    pattern2 = re.compile(r"(\d+\.\d+)\s*（?(周[一二三四五六日])）?\s*第(\d+)-(\d+)节")
    pattern3 = re.compile(r"(\d+\.\d+)\s*([一二三四五六日])\s*(\d+)-(\d+)")
    pattern4 = re.compile(r"(\d+\.\d+)\s*([一二三四五六日])\s*(上午|下午)")
    pattern5 = re.compile(r"(\d+\.\d+)（(周[一二三四五六日])(上午|下午)）")
    pattern6 = re.compile(r"(\d+\.\d+)（(周[一二三四五六日])）第(\d+)节")

    # 尝试匹配 pattern1
    match = pattern1.match(date_str)
    if match:
        date, sections, weekday = match.groups()
        month, day = map(int, date.split("."))
        start_section, end_section = map(int, sections.split("-"))
        date_obj = datetime(2024, month, day)
        week_number = calculate_week(date_obj)
        return week_to_number(weekday), start_section, end_section, week_number

    # 尝试匹配 pattern2
    match = pattern2.match(date_str)
    if match:
        date, weekday, start_section, end_section = match.groups()
        month, day = map(int, date.split("."))
        start_section, end_section = int(start_section), int(end_section)
        date_obj = datetime(2024, month, day)
        week_number = calculate_week(date_obj)
        return week_to_number(weekday), start_section, end_section, week_number

    # 尝试匹配 pattern3
    match = pattern3.match(date_str)
    if match:
        date, weekday, start_section, end_section = match.groups()
        month, day = map(int, date.split("."))
        weekday = "周" + weekday
        start_section, end_section = int(start_section), int(end_section)
        date_obj = datetime(2024, month, day)
        week_number = calculate_week(date_obj)
        return week_to_number(weekday), start_section, end_section, week_number

    # 尝试匹配 pattern4
    match = pattern4.match(date_str)
    if match:
        date, weekday, section = match.groups()
        month, day = map(int, date.split("."))
        weekday = "周" + weekday
        start_section = 1 if section == "上午" else 5
        end_section = 4 if section == "上午" else 8
        date_obj = datetime(2024, month, day)
        week_number = calculate_week(date_obj)
        return week_to_number(weekday), start_section, end_section, week_number

    # 尝试匹配 pattern5
    match = pattern5.match(date_str)
    if match:
        date, weekday, section = match.groups()
        month, day = map(int, date.split("."))
        start_section = 1 if section == "上午" else 5
        end_section = 4 if section == "上午" else 8
        date_obj = datetime(2024, month, day)
        week_number = calculate_week(date_obj)
        return week_to_number(weekday), start_section, end_section, week_number

    match = pattern6.match(date_str)
    if match:
        date, weekday, section = match.groups()
        month, day = map(int, date.split("."))
        start_section = int(section)
        end_section = int(section)
        date_obj = datetime(2024, month, day)
        week_number = calculate_week(date_obj)
        return week_to_number(weekday), start_section, end_section, week_number

    print(date_str)
    raise ValueError("无法处理的日期格式：", date_str)


def convert_course(sheet):
    # 读取 data.xlsx 文件，加载其中的 “体检理论” 工作表，不含标题行
    df = pd.read_excel(data_source, sheet_name=sheet, header=None)
    print("正在处理", sheet, "工作表", "...")
    # 打印行数，列数
    # print(df.shape)

    # 逐行检查，如果一行内只有一个单元格横跨所有列，那么这一行是一个标题行，打印
    for i in range(df.shape[0]):
        if i == 0:
            subject = df.iloc[i, 0]
            print(subject)
        if df.iloc[i].count() == 1:
            # print(df.iloc[i].values[0])
            continue
        else:
            # 打印非标题行的内容
            # print(df.iloc[i].values)
            """
            Input:
            1、2组日期/节次 3、4组日期/节次 5、6组日期/节次 7、8组日期/节次 授课科室 授课形式 授课内容 学时 一线教师 职称 二线教师 职称 授课地点
            ---
            日期 星期 节次 学时 授课形式 授课内容 授课科室 一线教师 职称 二线教师 职称 授课地点
            """
            cur_index = i + 1
            # 获取 i 行所有单元格内容
            title = df.iloc[i].tolist()
            raw_title = "".join(title)
            raw_title = re.sub("\s+", "", raw_title)

            if re.match("1、2组日期/节次", raw_title):
                title_mode = "sep"
            else:
                title_mode = "mix"
            cur_index = i + 1
            break

    while df.iloc[cur_index].count() != 1:
        if title_mode == "sep":
            # 获取前四列单元格
            date_cells = df.iloc[cur_index].tolist()[:4]

            is_sep = True
            # 查看是否有 nan
            if any(pd.isna(date_cells)):
                is_sep = False

            data_cells = [str(i) for i in date_cells]

            teacher_index = title.index("一线教师")
            place_index = title.index("授课地点")

            if is_sep:
                for i, sep in enumerate(sep_target):
                    with open(os.path.join(output_dir, f"{sep}.csv"), "a") as f:
                        # 课程名称, 星期, 开始节数, 结束节数, 老师, 地点, 周数
                        data_str = df.iloc[cur_index, i]
                        day, start, end, week = exact_day(data_str)
                        teacher = df.iloc[cur_index, teacher_index]
                        if not pd.isna(teacher):
                            teacher = re.sub("\s+", "", teacher)
                        place = df.iloc[cur_index, place_index]
                        place = re.sub("\s+", "", place)

                        if pd.isna(teacher):
                            teacher = "希波克拉底"
                            all_cells = [
                                j for j in df.iloc[cur_index].tolist() if not pd.isna(j)
                            ]
                            exam = all_cells[1]
                            f.write(
                                f"{exam},{day},{start},{end},{teacher},{place},{week}\n"
                            )
                        else:
                            f.write(
                                f"{sheet},{day},{start},{end},{teacher},{place},{week}\n"
                            )
            else:
                for sep in sep_target:
                    with open(os.path.join(output_dir, f"{sep}.csv"), "a") as f:
                        # 课程名称, 星期, 开始节数, 结束节数, 老师, 地点, 周数
                        data_str = df.iloc[cur_index, 0]
                        assert not pd.isna(data_str), cur_index
                        day, start, end, week = exact_day(data_str)
                        teacher = df.iloc[cur_index, teacher_index]
                        if not pd.isna(teacher):
                            teacher = re.sub("\s+", "", teacher)
                        place = df.iloc[cur_index, place_index]
                        place = re.sub("\s+", "", place)

                        if pd.isna(teacher):
                            teacher = "希波克拉底"
                            exam = sheet + "考试"
                            f.write(
                                f"{exam},{day},{start},{end},{teacher},{place},{week}\n"
                            )
                        else:
                            f.write(
                                f"{sheet},{day},{start},{end},{teacher},{place},{week}\n"
                            )

        else:
            data_cells = df.iloc[cur_index].tolist()[:3]
            data_cells = [str(i) for i in data_cells]

            teacher_index = title.index("一线教师")
            place_index = title.index("授课地点")

            for sep in sep_target:
                with open(os.path.join(output_dir, f"{sep}.csv"), "a") as f:
                    # 课程名称, 星期, 开始节数, 结束节数, 老师, 地点, 周数
                    data_str = "".join(data_cells)
                    day, start, end, week = exact_day(data_str)
                    teacher = df.iloc[cur_index, teacher_index]
                    if not pd.isna(teacher):
                        teacher = re.sub("\s+", "", teacher)
                    place = df.iloc[cur_index, place_index]
                    place = re.sub("\s+", "", place)

                    if pd.isna(teacher):
                        teacher = "希波克拉底"
                        exam = sheet + "考试"
                        f.write(
                            f"{exam},{day},{start},{end},{teacher},{place},{week}\n"
                        )
                    else:
                        f.write(
                            f"{sheet},{day},{start},{end},{teacher},{place},{week}\n"
                        )

        cur_index += 1


def init_output():
    os.makedirs(output_dir, exist_ok=True)
    for sep in sep_target:
        with open(os.path.join(output_dir, f"{sep}.csv"), "w") as f:
            # 课程名称,星期,开始节数,结束节数,老师,地点,周数
            f.write("课程名称,星期,开始节数,结束节数,老师,地点,周数\n")


output_dir = "wakeup"
sep_target = ["第1、2组", "第3、4组", "第5、6组", "第7、8组"]
data_source = "data.xlsx"


def main():
    init_output()
    # 获取所有工作表
    sheets = pd.ExcelFile(data_source).sheet_names
    sheets = sheets[1:]
    for sheet in sheets:
        convert_course(sheet)


if __name__ == "__main__":
    main()
