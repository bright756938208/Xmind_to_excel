import xmind
import os
import sys
import xlsxwriter
import configparser
import json


class MyConfigparse(configparser.ConfigParser):
    def __init__(self):
        configparser.ConfigParser.__init__(self, defaults=None)

    def optionxform(self, optionstr):
        # 原始方法为 return optionstr.lower()，会全部转换小写，这里修改下防止自动转换大小写
        return optionstr


# 这个可以直接拿到当前执行文件的目录
floder_path = os.path.split(os.path.abspath(sys.argv[0]))[0]
# xmind_path = os.path.join(floder_path, 'test.xmind')

config = MyConfigparse()
xmind_to_excel_config = os.path.join(floder_path, 'xmind_to_excel_config.ini')
# 读取配置文件
config.read(xmind_to_excel_config, encoding="utf-8-sig")
# 以词典来读取配置文件
_sections = config._sections
if not _sections:
    input("配置文件读取内容为空，请检查配置文件是否在同一个目录，按回车键继续...")
    raise Exception("请检查配置文件路径是否正确: %s" % xmind_to_excel_config)
# 解析基本配置
try:
    # excel_config_template_chose 的配置
    excel_config_template_chose_dict = dict(_sections['excel_config_template_chose'])
    excel_template_name_list = eval(excel_config_template_chose_dict["excel_template_name_list"])
    print_xmind_data = int(excel_config_template_chose_dict["print_xmind_data"])
    print_xmind_row_data = int(excel_config_template_chose_dict["print_xmind_row_data"])
    print_excel_data = int(excel_config_template_chose_dict["print_excel_data"])
except:
    input("excel_config_template_chose中的配置项内容解析出现问题，请检查是否填错，是否有如中文标点符号，格式错误，编码格式与系统不符合等情况\n按回车键继续...")
    raise Exception("")

xmind_path = input("输入xmind路径,按回车键即可，如按回车没反应多按几下，可直接拖入xmind:").strip()
xmind_path = xmind_path.strip('"')
xmind_path = xmind_path.strip("'")
# 根据xmind路径拿到xmind名，后面用来生成xlsx的命名和路径
name_index = xmind_path.rfind('\\')
if name_index < 0:
    name_index = xmind_path.rfind('/')
if name_index < 0:
    input("输入的xmind路径不对，请输入xmind绝对路径，当前xmind路径: %s\n按回车键继续..." % xmind_path)
    raise Exception("")
file_prefix_name = xmind_path[name_index + 1:xmind_path.rfind('.')]

try:
    xmind_workbook = xmind.load(xmind_path)
except Exception:
    input("xmind读取出现异常，请检查是否没有读取权限或路径有问题，xmind路径: %s\n按回车键继续..." % xmind_path)
    raise Exception("")
# 获得xmind中的数据,字典列表类型
xmind_datas = xmind_workbook.getData()


class ExcelTemplate:
    def __init__(self, _sections={}, excel_template_name="", file_prefix_name=""):
        # 将模板配置解析出来并存到字典中
        self.excel_template_name = excel_template_name
        try:
            self.excel_config_dict = dict(_sections[excel_template_name])
        except KeyError:
            input("模板:%s 在配置中找不到，请检查模板名是否写错\n按回车键继续..." % self.excel_template_name)
            raise Exception("")
        try:
            # excel模板配置
            self.excel_label_list = eval(self.excel_config_dict["excel_label_list"])
            self.column_width_list = eval(self.excel_config_dict["column_width_list"])
            self.not_repeat_column = json.loads(self.excel_config_dict["not_repeat_column"])
            self.column_default_value = json.loads(self.excel_config_dict["column_default_value"])

            self.nodes_config_dict = json.loads(self.excel_config_dict["nodes_config_dict"])
            self.replace_markers = json.loads(self.excel_config_dict["replace_markers"])

            self.font_size = int(self.excel_config_dict["font_size"])
            self.bold = int(self.excel_config_dict["bold"])
            self.text_wrap = int(self.excel_config_dict["text_wrap"])
        except json.decoder.JSONDecodeError:
            input("模板:%s, not_repeat_column, column_default_value, nodes_config_dict 或 replace_markers 中至少一个解析出错，"
                  "\n请检查json格式是否写错，如缺少引号，包含中文标点符号，包含空值，多了符号等\n按回车键继续..."
                  % self.excel_template_name)
            raise Exception("")
        except ValueError:
            input("模板:%s, font_size, bold 或 text_wrap 中至少一个解析出错，"
                  "\n请检查格式是否写错，包含中文标点符号等\n按回车键继续..."
                  % self.excel_template_name)
            raise Exception("")
        except SyntaxError:
            input("模板:%s, excel_label_list 或 column_width_list 中至少一个解析出错，"
                  "\n请检查格式是否写错或者包含空格，缺少引号，包含中文标点符号等\n按回车键继续..."
                  % self.excel_template_name)
            raise Exception("")
        except:
            input("模板:%s 内容解析出现问题，请检查如中文标点符号、类型错误、配置文件编码格式与系统编码克制不符合等情况\n"
                  "按回车键继续..." % self.excel_template_name)
            raise Exception("")

        xlsx_name = file_prefix_name + "_" + excel_template_name + ".xlsx"
        self.xlsx_path = os.path.join(floder_path, xlsx_name)
        self.workbook = xlsxwriter.Workbook(self.xlsx_path)
        # 定义单元格格式
        self.item_style = self.workbook.add_format({
            'font_size': self.font_size,  # 字体大小
            'bold': self.bold,  # 是否粗体
            'text_wrap': self.text_wrap  # 自动换行，可在文本中加 '\n'来控制换行的位置
        })


def process_xmind_data(xmind_all_level):
    """解析xmind数据，并转换成嵌套列表
    :param xmind_all_level: xmind一个画布中下的某个中心主题列表
    :return: 转换成一行一行的嵌套的列表
    """
    xmind_data_list = [[], []]
    up_level_data = []
    global level
    level = 0
    global row
    row = 1

    def process_data(xmind_all_level):
        global level
        global row
        for level0 in xmind_all_level:
            # 现将子级数据存储起来
            next_level = level0.get("topics")
            if next_level:
                # 把子分支数据删除，避免当数据过多时，存储的上级数据特别大
                level0.pop("topics")
                up_level_data.append(level0)
                # print(up_level_data)
                level += 1
                # 将存储的下级数据再带进去，尽管上面删除的子级数据这里也不会造成数据丢失
                process_data(next_level)
            else:
                xmind_data_list[row] = up_level_data + [level0]
                row += 1
                xmind_data_list.append([])
            # 这里level0是浅拷贝，level0.pop("topics")中删除键值对，level_all[-1]中对应的也会删除，不会出现永远不相等的情况
            if level0 == xmind_all_level[-1]:
                try:
                    up_level_data.pop()
                except:
                    pass
                level -= 1

    process_data(xmind_all_level)
    return xmind_data_list


def get_excel_row_data_for_xmind_row(xmind_row_data, excel_template):
    '''
    将xmind行数据按配置的规则转换成excel行数据
    :param excel_row_data:
    :param row:
    :param worksheet:
    :return: excel行数据列表
    '''
    excel_template_name = excel_template.excel_template_name

    def extract_column_default_value(excel_row_data):
        # 将配置的列默认值写入
        if excel_template.column_default_value:
            for key in excel_template.column_default_value:
                try:
                    default_value_column = excel_template.excel_label_list.index(key)
                    excel_row_data[default_value_column] = excel_template.column_default_value[key]
                except ValueError:
                    print("模板:%s, 默认值配置column_default_value中的列[ %s ]在excel_label_list中找不到，本次将不会写入默认值"
                          % (excel_template_name, key))

    # 获取节点数对应的配置名
    nodes = str(len(xmind_row_data))
    node_config_name = None
    for nodes_key in excel_template.nodes_config_dict:
        if nodes in nodes_key.split("|"):
            node_config_name = excel_template.nodes_config_dict[nodes_key]
    if not node_config_name:
        # 说明没有这个节点数对应点配置，将采用默认方式写一行
        print("模板:%s, 中没有找到节点数为 %s 的配置项, 节点数配置: %s, 本次会将整行数据的title按顺序存入"
              % (excel_template_name, nodes, excel_template.nodes_config_dict))
        excel_row_data = [cell_data["title"] for cell_data in xmind_row_data]
        print("行数据内容：%s" % excel_row_data)
        while len(excel_row_data) < len(excel_template.excel_label_list):
            excel_row_data.append("")
        extract_column_default_value(excel_row_data)
        return excel_row_data

    # 先将列表按节点数写好，后续方便直接通过下标写入值
    excel_row_data = [None for aa in range(len(excel_template.excel_label_list))]
    extract_column_default_value(excel_row_data)
    # 用对应节点数的配置项规则来写数据
    try:
        node_config = json.loads(excel_template.excel_config_dict[node_config_name])
    except KeyError:
        input("模板:%s, 节点配置项[ %s ]找不到，请检查 节点数配置项 或 节点规则配置名 是否写错或者包含空格，缺少引号等\n节点数配置项:%s\n按回车键继续..."
              % (excel_template_name, node_config_name, excel_template.nodes_config_dict))
        raise Exception("")
    except json.decoder.JSONDecodeError:
        input("模板:%s, 节点配置项[ %s ]解析出错，请检查json格式是否写错或缺少引号，包含中文标点符号等\n按回车键继续..."
              % (excel_template_name, node_config_name))
        raise Exception("")
    for key in node_config:
        try:
            column = excel_template.excel_label_list.index(key)
        except ValueError:
            input("模板:%s, 在节点配置[%s]中,key [%s] 不在 excel_label_list中，请确认[%s]是否写错、包含空格等\n按回车键继续..."
                  % (excel_template_name, node_config_name, key, key))
            raise Exception("")
        cell_content_matching_rule = node_config[key]
        if not cell_content_matching_rule:
            # 如果没有配置，说明是空值，不用管了
            continue
        # 将多个规则用 | 分开到list中
        cell_content_rule_list = cell_content_matching_rule.split("|")
        for cell_content_rule in cell_content_rule_list:
            # 用:号分割规则生效的前提条件
            rule_list = cell_content_rule.split(":")
            if len(rule_list) > 1:
                condition_type = rule_list[0]
                condition = rule_list[1]
                cell_content_rule = rule_list[2]
                # 处理规则前提条件
                if condition_type == "if":
                    if not get_cell_data_for_rule(xmind_row_data, condition, excel_template):
                        continue
                elif condition_type == "ifnot":
                    if get_cell_data_for_rule(xmind_row_data, condition, excel_template):
                        continue
                else:
                    input("检测到规则条件配置错误, 规则条件配置错误，目前配置的条件:%s，规则条件只能为 if 或 notif\n"
                          "模板:%s, 规则名:%s，关键字:%s, 规则:%s"
                          % (condition_type, excel_template_name, node_config_name, key, cell_content_matching_rule))
                    raise Exception("按回车键继续...")
            else:
                cell_content_rule = rule_list[0]
                if not cell_content_rule:
                    print("检测到有包含空的规则, 模板:%s 规则名:%s，关键字:%s, 规则:%s "
                          % (excel_template_name, node_config_name, key, cell_content_matching_rule))
                    continue
            cell_data = get_cell_data_for_rule(xmind_row_data, cell_content_rule, excel_template)
            # 当获取到不为空的数据时，就写入并且退出循环不在管|后面的规则匹配了
            if cell_data:
                excel_row_data[column] = cell_data
                break
    return excel_row_data


def get_cell_data_for_rule(xmind_row_data, cell_content_rule, excel_template):
    '''
    根据传入的规则，从xmind行数据中提取对应单元格的数据
    :param xmind_row_data:
    :param cell_content_rule:
    :return:
    '''
    cell_data = ""
    sub_key = "title"
    excel_template_name = excel_template.excel_template_name
    if not cell_content_rule:
        return cell_data
    if cell_content_rule.find("~") >= 0:
        # 含有'~'意味着需要两个以上层级的数据字符 进行拼接
        index = cell_content_rule.find("~")
        head_level = int(cell_content_rule[:index])
        end_level = int(cell_content_rule[index + 1:])
        current_level = head_level
        try:
            cell_data = xmind_row_data[head_level][sub_key]
            # 将不同层级数据进行拼接
            while xmind_row_data[current_level] != xmind_row_data[end_level]:
                current_level += 1
                # 防止为空拼接字符串出现crash
                if cell_data is None:
                    cell_data = ""
                row_data = xmind_row_data[current_level][sub_key]
                if row_data is None:
                    row_data = ""
                cell_data = cell_data + "_" + row_data
        except IndexError:
            input("~配置要求层级超过实际层级，行总层级数据:%s，开始层级:%s，结束层级:%s，"
                  "请检查模板:%s，节点数为 %s 的规则配置是否层级超出范围\n按回车键继续..."
                  % ([e["title"] for e in xmind_row_data], head_level, end_level, excel_template_name,
                     len(xmind_row_data)))
            raise Exception("")
    elif cell_content_rule.find(".") >= 0:
        # 含有'.'意味着需要该层级子类数据
        index = cell_content_rule.find(".")
        current_level = int(cell_content_rule[:index])
        sub_key = cell_content_rule[index + 1:]
        if "markers" in sub_key:
            # 标记提取和替换
            cell_data = get_replaced_markers(markers=xmind_row_data[current_level]["markers"], sub_key=sub_key,
                                             excel_template=excel_template)
        else:
            try:
                cell_data = xmind_row_data[current_level][sub_key]
            except KeyError:
                print("xmind数据中找不到['%s']这个关键字，模板:%s, 规则:%s ,请检查节点数为 %s 的配置项规则是否写错、包含空格、xmind数据是否正常，找不到的将为空值\n"
                      "当前节点xmind数据:%s" % (sub_key, excel_template_name, cell_content_rule, len(xmind_row_data),
                                          xmind_row_data[current_level]))
    else:
        current_level = int(cell_content_rule)
        try:
            cell_data = xmind_row_data[current_level][sub_key]
        except KeyError:
            print("xmind数据中找不到['%s']这个关键字，模板:%s, 节点数:%s, 规则:%s ,请检查节点数为 %s 的配置项规则是否写错、包含空格、xmind数据是否正常，找不到的将为空值\n"
                  "当前节点xmind数据:%s" % (
                  sub_key, excel_template_name, cell_content_rule, len(xmind_row_data), xmind_row_data[current_level]))
    return cell_data


def get_replaced_markers(markers, sub_key, excel_template):
    """
    获取根据配置项转换后的markers
    :param markers:
    :param sub_key:
    :param excel_template:
    :return: 转换后的markers值，如markers为空列表，则返回""
    """
    if not markers:
        return ""
    # 先要将 sub_key 中的.切割出来，默认为"priority"，也就是优先级
    key = "priority"
    if sub_key.find(".") >= 0:
        index = sub_key.find(".")
        key = sub_key[index + 1:]
    # markers 是一个列表，并且每一项值是用 - 区分类型和该类型的值如 "priority-4"
    for marker in markers:
        if marker.find(key) >= 0:
            if marker in excel_template.replace_markers:
                return excel_template.replace_markers[marker]


def process_repeat_for_column(writer_excel_datas, excel_template):
    '''
    处理重复行数据
    :return: 最终的excel数据列表
    '''
    row_number_max = len(writer_excel_datas)
    is_repeat = False
    for column_title in excel_template.not_repeat_column:
        # 检查是否有配置异常，如没异常则到列名对应的列数索引
        if column_title in excel_template.excel_label_list:
            column = excel_template.excel_label_list.index(column_title)
        else:
            input("模板:%s, 不重复列配置 not_repeat_column 中的列名[ %s ]不在excel_label_list中，本次不会处理该重复列\n按回车键继续..."
                  % (excel_template.excel_template_name, column_title))
            continue
        repeat_column_process_mode = excel_template.not_repeat_column[column_title]

        # 先将需要去重列的数据存到列表
        column_list = []
        for row_list in writer_excel_datas:
            column_list.append(row_list[column])
        current_row_number = -1
        for current_row in column_list:
            repeat = 0
            current_row_number += 1
            row_number = current_row_number
            while row_number < row_number_max - 1:
                row_number += 1
                # 再将一个单元格的数据从下一行的数据开始挨个对比
                if current_row != writer_excel_datas[row_number][column]:
                    continue
                else:
                    repeat += 1
                    is_repeat = True
                    if repeat_column_process_mode == 1:
                        # 需要修改当前列数据，修改为添加了自增1的字符串
                        process_row = current_row + "_%s" % repeat
                        writer_excel_datas[row_number][column] = process_row
                        print("[ %s ], 第%s行检查到了重复内容，已自动修改为: %s" % (column_title, row_number + 1, process_row))
                    else:
                        print("[ %s ], 第%s行检查到了重复内容: %s" % (column_title, row_number + 1, current_row))
    if is_repeat and repeat_column_process_mode not in [1, 2]:
        input("模板:%s, 重复内容检查完毕！ 请好好修改，按回车键继续..." % excel_template.excel_template_name)
        raise Exception("")
    return writer_excel_datas


def write_xmind_data_to_excel(xmind_row_data_list_dict, excel_template):
    '''
    :param xmind_row_data_list_dict: xmind行数据列表，列表是一个画布， 字典保存了多个画布
    :param excel_template: 生成excel的模板
    :return:
    '''
    for sheet_name in xmind_row_data_list_dict:
        xmind_row_data_list = xmind_row_data_list_dict[sheet_name]
        write_excel_row_list = []
        # 将配置的表格的第一行先存入excel行数据
        write_excel_row_list.append(excel_template.excel_label_list)
        # 将xmind行数据根据配置的节点规则转成excel 行数据并存起来
        for xmind_row_data in xmind_row_data_list:
            # print([abc["title"] for abc in excel_row_data])
            if xmind_row_data:
                excel_row_data = get_excel_row_data_for_xmind_row(xmind_row_data, excel_template)
                write_excel_row_list.append(excel_row_data)

        # 打印xmind节点配置转换成的excel数据
        if print_excel_data:
            for ex_data in write_excel_row_list:
                print(ex_data)
            input("已经打印xmind节点配置转换成的excel数据(未去重),按回车键继续\n")

        # 处理配置要求的列重复的数据，没有配置时不会处理
        write_excel_row_list = process_repeat_for_column(write_excel_row_list, excel_template)

        # 将编辑excel的一个sheet
        workbook = excel_template.workbook
        worksheet = workbook.add_worksheet(sheet_name)
        # 设置列宽度
        column_index = 0
        for column_width in excel_template.column_width_list:
            worksheet.set_column(column_index, column_index, column_width)
            column_index += 1
        # 将最终的数据一行一行写入excel
        print("开始将xmind画布 [%s] 的数据写入到excel" % sheet_name)
        row = 0
        for row_data_list in write_excel_row_list:
            worksheet.write_row(row, 0, row_data_list, excel_template.item_style)
            row += 1
        print("xmind画布 [%s] 的数据写入完毕。" % sheet_name)
    try:
        workbook.close()
    except PermissionError:
        input("保存excel失败，没有保存文件的权限，请检查是否已经打开了相同命名的excel，\n如果已经打开: %s 请先关闭它\n按回车键继续..."
              % excel_template.xlsx_path)
        raise Exception("")
    print("excel已生成, 保存路径: %s" % excel_template.xlsx_path)


def main():
    if print_xmind_data != 0:
        # 打印xmind中的数据,json类型
        print(xmind_workbook.to_prettify_json())
        input("已打印读取到的xmind数据，按回车键继续...\n")

    xmind_data_row_list_dict = {}
    # 将xmind每个画布的数据处理成一行一行的list
    for canvas in xmind_datas:
        xmind_all_level = [canvas["topic"]]
        # 拿到初次处理的数据xmind数据
        xmind_data_list = process_xmind_data(xmind_all_level)
        canvas_name = canvas["title"]
        xmind_data_row_list_dict[canvas_name] = xmind_data_list
        if print_xmind_row_data:
            print(xmind_data_list)
            input("已打印画布[%s]转换成一行一行的数据，按回车键继续...\n" % canvas_name)

    # 每个模板分别输出excel
    for excel_template_name in excel_template_name_list:
        excel_template = ExcelTemplate(_sections, excel_template_name, file_prefix_name)
        print("开始使用模板 [%s] 来生成excel" % excel_template_name)
        write_xmind_data_to_excel(xmind_data_row_list_dict, excel_template)


if __name__ == '__main__':
    main()
    input("执行完毕，按回车键继续...")
