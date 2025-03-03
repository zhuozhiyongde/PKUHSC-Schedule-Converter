from docx2python import docx2python
import pandas as pd
import re
import os

doc = docx2python("data.docx")
doc_body = doc.body

weeks = 1


def clean_cell(cell):
    # return dict，key 为组号，value 为课程，键包括 4 个组，以及一个 all
    # 清理数据，删 -，对于换行符进行 split
    # '毒理学基础414' -> ['毒理学基础', '414']
    # "['1组', '-', '2组', '营养卫生学（讨论）', '3组', '-', '4组', '-']"
    # "['1组', '流病实习', '417', '', '2组', '环卫实习', '3组', '毒理实习', '4组', '职卫实习']"
    # 从 cell 中移除 ""
    cell = [item for item in cell if item != ""]
    if len(cell) == 0:
        return None
    """
    课程名称,星期,开始节数,结束节数,老师,地点,周数
    技能培训,3,5,7,张新颜,北楼B1层101教室,2
    这里只需要返回课程名称，地点
    """
    # print(cell)
    raw = "".join(cell)
    raw = raw.replace(" ", "")
    # replace 中文-中文 中的 -
    raw = re.sub(r"(?<=[\u4e00-\u9fa5])-(?=[\u4e00-\u9fa5])", "", raw)
    if "放假" in raw or "调休" in raw:
        return None
    if "组" not in raw:
        # 中文 regex
        course_pattern = r"[\u4e00-\u9fa5]+"
        # 找到所有中文
        course = re.findall(course_pattern, raw)
        # print(course)
        # 找到地点
        location_pattern = r"(?:\d{3}、)?\d{3}"
        location = re.findall(location_pattern, raw)
        # print(location)
        assert len(course) == 1 and len(location) == 1, [cell] + [raw]
        output = {
            "all": {
                "course": course[0],
                "location": location[0],
            }
        }
        return output
    else:
        # '1、2组职卫实习参观' -> {1: '职卫实习参观', 2: '职卫实习参观'}
        # 找到所有仅单个的数字，不允许多数字
        if "、" not in raw:
            # print(len(cell), cell)
            key = 1
            output = {}
            for item in cell:
                if "组" in item:
                    key = int(re.search(r"\d+", item).group())
                    output[key] = {}
                else:
                    if item == "-":
                        output[key] = None
                    elif re.match(r"\d{3}", item):
                        output[key]["location"] = item
                    else:
                        output[key]["course"] = item
            return output
        else:
            groups = re.findall(r"\d+", raw)
            course = re.findall(r"[\u4e00-\u9fa5]+", raw)
            course = course[0]
            # 移除“组”
            course = re.sub(r"组", "", course)
            output = {}
            for group in groups:
                output[int(group)] = {
                    "course": course,
                }
            return output


def init_output():
    # 先删除 1,2,3,4.csv
    for i in range(1, 5):
        if os.path.exists(f"{i}.csv"):
            os.remove(f"{i}.csv")
    # 初始化 1,2,3,4.csv
    for i in range(1, 5):
        # 课程名称,星期,开始节数,结束节数,老师,地点,周数
        with open(f"{i}.csv", "w", encoding="utf-8") as f:
            f.write("课程名称,星期,开始节数,结束节数,老师,地点,周数\n")


init_output()

for table in doc_body:
    table = pd.DataFrame(table)
    # print shape
    # print(table.shape)
    if table.shape[0] < 10:
        continue
    # 导出为 csv 文件
    table = table.iloc[1:-4, 2:]
    # print(table)
    # table.to_csv(f"table_{weeks}.csv", index=False)
    # 清理数据，逐单元格调用，传入 info 为 weeks, row, col
    for day, col in enumerate(table.columns):
        for index, value in table[col].items():
            # 处理value
            # (weeks, index, col)
            output = clean_cell(value)
            if output is not None:
                """
                课程名称,星期,开始节数,结束节数,老师,地点,周数
                """
                # print(output, weeks, index, day)
                if "all" in output:
                    for i in range(1, 5):
                        with open(f"{i}.csv", "a", encoding="utf-8") as f:
                            # 检查是否有 location 这个 key
                            if "location" in output["all"]:
                                f.write(
                                    f"{output['all']['course']},{day + 1},{index},{index},,{output['all']['location']},{weeks}\n"
                                )
                            else:
                                f.write(
                                    f"{output['all']['course']},{day + 1},{index},{index},,,{weeks}\n"
                                )
                else:
                    for key, value in output.items():
                        if value is None:
                            continue
                        with open(f"{key}.csv", "a", encoding="utf-8") as f:
                            # 检查是否有 location 这个 key
                            if "location" in value:
                                f.write(
                                    f"{value['course']},{day + 1},{index},{index},,{value['location']},{weeks}\n"
                                )
                            else:
                                f.write(
                                    f"{value['course']},{day + 1},{index},{index},,,{weeks}\n"
                                )
    weeks += 1
    # exit(0)
