"""This module convert Excel files to GDScript files."""

#! /usr/bin/env python
# -*- coding: utf-8 -*
# @description: A tool to convert excel files to GDScprit
# @copyright Hidden Moss, https://hiddenmoss.com/
# @author Yuancheng Zhang, https://github.com/endaye
# @see repo: https://github.com/hidden-moss/xls2gd

import os
import os.path
import sys
import datetime
import codecs
import subprocess
import json
import re
import xlrd
import csv

__authors__ = ["Yuancheng Zhang"]
__copyright__ = "Copyright 2024, Hidden Moss"
__credits__ = ["Yuancheng Zhang"]
__license__ = "MIT"
__version__ = "v1.1.0"
__maintainer__ = "Yuancheng Zhang"
__status__ = "Development"

INPUT_FOLDER = "./"
OUTPUT_GD_FOLDER = "./"
OUTPUT_GD_NAME_TEMPLATE = "data_{sheet_name}.gd"
OUTPUT_CSV_FOLDER = "./"
OUTPUT_CSV_NAME_TEMPLATE = "locale_{sheet_name}.csv"


INFO = {"c": "\033[36minfo\033[0m", "b": "info"}
ERROR = {"c": "\033[31merror\033[0m", "b": "error"}
SUCCESS = {"c": "\033[92msucess\033[0m", "b": "success"}
FAILED = {"c": "\033[91mfailed\033[0m", "b": "failed"}

SCRIPT_HEAD = """# @copyright Hidden Moss
# @see Hidden Moss: https://hiddenmoss.com/
# @see repo: https://github.com/hidden-moss/xls2gd
# @author Yuancheng Zhang, https://github.com/endaye
# 
#* source file: %s
#! THIS FILE IS GENERATED BY tool_xls2gd.py.
#! DONT'T CHANGE IT MANULLY.

"""

# data type
INT, FLOAT, STRING, BOOL = r"int", r"float", r"string", r"bool"
INT_ARR, FLOAT_ARR, STRING_ARR, BOOL_ARR = r"int[]", r"float[]", r"string[]", r"bool[]"
VECTOR2, VECTOR3, COLOR = "vector2", "vector3", "color"
GDSCRIPT, COMMENT = "gdscript", "comment"
TRANSLATE = "translate"
DEFAULT_LANG = "zh_CN"

CONFIG_FILE = "tool_xls2gd.config"

KEY_1, KEY_2, KEY_3 = "key1", "key2", "key3"

GUI = None
GD_CNT = 0
MAX_XLS_NAME_LEN = 0
IS_COLOR = False


def make_table(filename):
    """Make tables from excel file."""
    if not os.path.isfile(filename):
        raise NameError(f"{filename} is not a valid filename")
    book_xlrd = xlrd.open_workbook(filename)

    excel = {}
    excel["filename"] = filename
    excel["data"] = {}
    excel["meta"] = {}
    excel["csv"] = {}

    for sheet in book_xlrd.sheets():
        sheet_name = sheet.name.replace(" ", "_")
        if not sheet_name.startswith("o-"):
            continue

        sheet_name_array = sheet_name.split("-")
        sheet_name = sheet_name_array[-1]

        # log(sheet_name +' sheet')
        data = excel["data"][sheet_name] = {}
        meta = excel["meta"][sheet_name] = {}
        t_csv = excel["csv"][sheet_name] = {}
        meta["kv"] = "kv" in sheet_name_array
        meta["has_csv"] = False

        # 必须大于4行
        if sheet.nrows < 4:
            return {}, -1, f"sheet[{sheet_name}] rows must > 4"

        # 解析标题和类型
        col_idx = 1
        type_dict = {}
        for col_idx in range(sheet.ncols):
            title = str(sheet.cell_value(1, col_idx)).replace(" ", "_")
            title_type = sheet.cell_type(1, col_idx)
            type_name = str(sheet.cell_value(2, col_idx)).lower()
            type_type = sheet.cell_type(2, col_idx)
            # 检查标题数据格式
            if title is None:
                return (
                    {},
                    -1,
                    f"sheet[{sheet_name}] title columns[{col_idx + 1}] must be String",
                )
            if title_type != xlrd.XL_CELL_TEXT:
                return (
                    {},
                    -1,
                    f"sheet[{sheet_name}] title columns[{col_idx + 1}] must be String",
                )
            # 检查类型数据格式
            if type_type != xlrd.XL_CELL_TEXT:
                return (
                    {},
                    -1,
                    f"sheet[{sheet_name}] type columns[{col_idx + 1}] must be String",
                )
            if (
                type_name != INT
                and type_name != FLOAT
                and type_name != STRING
                and type_name != BOOL
                and type_name != INT_ARR
                and type_name != FLOAT_ARR
                and type_name != STRING_ARR
                and type_name != BOOL_ARR
                and type_name != VECTOR2
                and type_name != VECTOR3
                and type_name != COLOR
                and type_name != COMMENT
                and type_name != GDSCRIPT
                and type_name != TRANSLATE
            ):
                return (
                    {},
                    -1,
                    f"sheet[{sheet_name}] type column[{col_idx + 1}] type wrong",
                )
            type_dict[title] = type_name
            if type_name == TRANSLATE:
                meta["has_csv"] = True

        meta["type_dict"] = type_dict

        # *读取主键key1，key2，key3，主键类型必须是Int或者String
        row_idx, col_idx = 3, 0
        for col_idx in range(sheet.ncols):
            key = str(sheet.cell_value(row_idx, col_idx)).lower()
            col_name = str(sheet.cell_value(1, col_idx))
            col_type = str(sheet.cell_value(2, col_idx)).lower()
            if key in (KEY_1, KEY_2, KEY_3):
                if col_type not in (INT, FLOAT, STRING):
                    return (
                        {},
                        -1,
                        f"sheet[{sheet_name}] {key} type must be Int, Float, or String",
                    )
                meta[key] = col_name

        # 检查主键
        if (
            (KEY_3 in meta and not (KEY_2 in meta and KEY_1 in meta))
            or (KEY_2 in meta and KEY_1 not in meta)
            or (KEY_1 not in meta)
        ):
            return {}, -1, f"sheet[{sheet_name}] {KEY_1} {KEY_2} {KEY_3} are wrong"

        key1 = meta[KEY_1] if KEY_1 in meta else None
        key2 = meta[KEY_2] if KEY_2 in meta else None
        key3 = meta[KEY_3] if KEY_3 in meta else None

        # 读取数据，从第5行开始
        row_idx = 4
        for row_idx in range(row_idx, sheet.nrows):
            row = {}
            key_v1, key_v3, key_v2 = None, None, None
            lang_kv = {}

            for col_idx in range(sheet.ncols):
                title = sheet.cell_value(1, col_idx)
                value = sheet.cell_value(row_idx, col_idx)
                vtype = sheet.cell_type(row_idx, col_idx)
                # 本行有数据
                v = None
                if type_dict[title] == INT and vtype == xlrd.XL_CELL_NUMBER:
                    v = int(value)
                elif type_dict[title] == FLOAT and vtype == xlrd.XL_CELL_NUMBER:
                    v = float(value)
                elif type_dict[title] == STRING:
                    v = format_str(value)
                elif type_dict[title] == BOOL and vtype == xlrd.XL_CELL_BOOLEAN:
                    v = "true" if value == 1 else "false"
                elif type_dict[title] == INT_ARR and vtype == xlrd.XL_CELL_TEXT:
                    v = str(value)
                elif type_dict[title] == FLOAT_ARR and vtype == xlrd.XL_CELL_TEXT:
                    v = str(value)
                elif type_dict[title] == STRING_ARR:
                    v = format_str(value)
                elif type_dict[title] == BOOL_ARR and vtype == xlrd.XL_CELL_TEXT:
                    v = str(value)
                elif type_dict[title] == VECTOR2 and vtype == xlrd.XL_CELL_TEXT:
                    v = str(value)
                elif type_dict[title] == VECTOR3 and vtype == xlrd.XL_CELL_TEXT:
                    v = str(value)
                elif type_dict[title] == COLOR and vtype == xlrd.XL_CELL_TEXT:
                    v = str(value)
                elif type_dict[title] == GDSCRIPT and vtype in (
                    xlrd.XL_CELL_TEXT,
                    xlrd.XL_CELL_NUMBER,
                ):
                    v = str(value)
                elif type_dict[title] == GDSCRIPT and vtype == xlrd.XL_CELL_BOOLEAN:
                    v = "true" if value == 1 else "false"
                elif type_dict[title] == TRANSLATE and vtype == xlrd.XL_CELL_TEXT:
                    v = str(value)
                    if key_v1 and key_v2 and key_v3:
                        key_csv = f"{sheet_name}_{title}_{key_v1}_{key_v2}_{key_v3}"
                    elif key_v1 and key_v2:
                        key_csv = f"{sheet_name}_{title}_{key_v1}_{key_v2}"
                    elif key_v1:
                        key_csv = f"{sheet_name}_{title}_{key_v1}"
                    key_csv = key_csv.replace(" ", "_").upper()
                    t_csv[key_csv] = str(value)
                    v = key_csv

                elif type_dict[title] == COMMENT:
                    continue

                row[title] = v

                if title == key1:
                    key_v1 = v
                if title == key2:
                    key_v2 = v
                if title == key3:
                    key_v3 = v

                # TODO: 检查key_v1是类型是string的话，不能为数字，需要符合gd命名规范

            # 键值检查
            if not (key1 is None or key2 is None or key3 is None):
                if key_v1 not in data:
                    data[key_v1] = {}
                if key_v2 not in data[key_v1]:
                    data[key_v1][key_v2] = {}
                if key_v3 is None:
                    return (
                        {},
                        -1,
                        f'sheet[{sheet_name}][{row_idx + 1}] {KEY_3} data "{key3}" is empty',
                    )
                elif key_v3 in data[key_v1][key_v2]:
                    return (
                        {},
                        -1,
                        f'sheet[{sheet_name}][{row_idx + 1}] {KEY_3} data "{key3}" is duplicated',
                    )
                else:
                    data[key_v1][key_v2][key_v3] = row
                    lang_suffix = f"{key_v1}_{key_v2}_{key_v3}"
            elif not (key1 is None or key2 is None):
                if key_v1 not in data:
                    data[key_v1] = {}
                if key_v2 is None:
                    return (
                        {},
                        -1,
                        f'sheet[{sheet_name}][{row_idx + 1}] {KEY_2} data "{key2}" is empty',
                    )
                elif key_v2 in data[key_v1]:
                    return (
                        {},
                        -1,
                        f'sheet[{sheet_name}][{row_idx + 1}] {KEY_2} data "{key2}" is duplicated',
                    )
                else:
                    data[key_v1][key_v2] = row
                    lang_suffix = f"{key_v1}_{key_v2}"
            elif key1 is not None:
                if key_v1 is None:
                    return (
                        {},
                        -1,
                        f'sheet[{sheet_name}][{row_idx + 1}] {KEY_1} data "{key1}" is empty',
                    )
                elif key_v1 in data:
                    return (
                        {},
                        -1,
                        f'sheet[{sheet_name}][{row_idx + 1}] {KEY_1} data "{key1}" is duplicated',
                    )
                else:
                    data[key_v1] = row
                    lang_suffix = str(key_v1)
            else:
                return {}, -1, f'sheet[{sheet_name}] missing "Key"s'

            for k, v in lang_kv.items():
                lang_id = v + lang_suffix
                row[k] = lang_id

    return excel, 0, "ok"


def format_str(v):
    """Format strings."""
    if isinstance(v, int) or isinstance(v, float):
        v = str(v)
    s = v
    s = v.replace('"', "'")
    # s = v.replace('\"', '\\\'')
    # s = s.replace('\'', '\\\'')
    return s


def get_int(v):
    """Get interger."""
    if v is None:
        return "null"
    return v


def get_float(v):
    """Get float."""
    if v is None:
        return "null"
    return v


def get_string(v):
    """Get string."""
    if v is None:
        return "null"
    return '"' + v + '"'


def get_bool(v):
    """Get boolean."""
    if v is None:
        return "null"
    return v


def get_int_arr(v):
    """Get interger array."""
    if v is None:
        return "null"
    tmp_vec_str = v.split(",")
    res_str = "["
    i = 0
    for val in tmp_vec_str:
        if val is not None and val != "":
            if i != 0:
                res_str += ", "
            res_str = res_str + val
            i += 1
    res_str += "]"
    return res_str


def get_float_arr(v):
    """Get float array."""
    if v is None:
        return "null"
    tmp_vec_str = v.split(",")
    res_str = "["
    i = 0
    for val in tmp_vec_str:
        if val is not None and val != "":
            if i != 0:
                res_str += ", "
            res_str = res_str + val
            i += 1
    res_str += "]"
    return res_str


def get_string_arr(v):
    """Get string array."""
    if v is None:
        return "null"
    tmp_vec_str = v.split(",")
    res_str = "["
    i = 0
    for val in tmp_vec_str:
        if val is not None and val != "":
            if i != 0:
                res_str += ", "
            res_str = res_str + '"' + val + '"'
            i += 1
    res_str += "]"
    return res_str


def get_bool_arr(v):
    """Get boolean array."""
    if v is None:
        return "null"
    tmp_vec_str = v.split(",")
    res_str = "["
    i = 0
    for val in tmp_vec_str:
        if val is not None and val != "":
            if i != 0:
                res_str += ", "
            res_str = res_str + val.lower()
            i += 1
    res_str += "]"
    return res_str


def get_vector2(v):
    """Get Vector2."""
    if v is None:
        return "null"
    tmp_vec_str = v.split(",")
    if len(tmp_vec_str) != 2:
        # todo: error check
        return "null"
    res_str = "Vector2("
    i = 0
    for val in tmp_vec_str:
        if val is not None and val != "":
            if i != 0:
                res_str += ", "
            res_str = res_str + val.lower()
            i += 1
    res_str += ")"
    return res_str


def get_vector3(v):
    """Get Vector3."""
    if v is None:
        return "null"
    tmp_vec_str = v.split(",")
    if len(tmp_vec_str) != 3:
        # todo: error check
        return "null"
    res_str = "Vector3("
    i = 0
    for val in tmp_vec_str:
        if val is not None and val != "":
            if i != 0:
                res_str += ", "
            res_str = res_str + val.lower()
            i += 1
    res_str += ")"
    return res_str


def get_color(v):
    """Get Color."""
    if v is None:
        return "null"
    tmp_vec_str = v.split(",")
    if len(tmp_vec_str) != 4:
        # todo: error check
        return "null"
    res_str = "Color("
    i = 0
    for val in tmp_vec_str:
        if val is not None and val != "":
            if i != 0:
                res_str += ", "
            res_str = res_str + val.lower()
            i += 1
    res_str += ")"
    return res_str


def get_gd(v):
    """Get GDScript."""
    if not v:
        return "null"
    return v


def get_translate(v):
    """Get translate."""
    # TODO: check translate
    if v is None:
        return "null"
    return '"' + v + '"'


def write_to_gd_script(excel, output_gd_path, output_csv_path, xls_file):
    """Write to GDScript."""
    for sheet_name, sheet in excel["data"].items():
        meta = excel["meta"][sheet_name]
        type_dict = meta["type_dict"]
        key1 = meta[KEY_1] if KEY_1 in meta else None
        key2 = meta[KEY_2] if KEY_2 in meta else None
        key3 = meta[KEY_3] if KEY_3 in meta else None

        gd_file_name = OUTPUT_GD_NAME_TEMPLATE.format(sheet_name=sheet_name)
        suffix = ""
        outfp = codecs.open(output_gd_path + "/" + gd_file_name, "w", "utf-8")
        outfp.write(SCRIPT_HEAD % (excel["filename"].replace(".//", "")))
        outfp.write("var " + sheet_name + suffix + " = {\r\n")

        if key1 and key2 and key3:
            write_to_gd_key(sheet, [key1, key2, key3], type_dict, outfp, 1)
        elif key1 and key2:
            write_to_gd_key(sheet, [key1, key2], type_dict, outfp, 1)
        elif key1 and (not meta["kv"]):
            write_to_gd_key(sheet, [key1], type_dict, outfp, 1)
        elif key1:
            # key-value style sheet
            write_to_gd_kv(sheet, [key1], type_dict, outfp, 1)
        else:
            outfp.close()
            raise RuntimeError("key missing")

        outfp.write("}\r\n")
        outfp.close()
        global GD_CNT
        GD_CNT += 1
        log(SUCCESS, f"[{GD_CNT:02d}] {xls_file:{MAX_XLS_NAME_LEN}} => {gd_file_name}")
        if meta["has_csv"]:
            csv_sheet = excel["csv"][sheet_name]
            if len(csv_sheet) > 0:
                write_to_csv(csv_sheet, sheet_name, output_csv_path, xls_file)


def write_to_gd_key(data, keys, type_dict, outfp, depth):
    """Write to GDScript. Promary key style sheet."""
    cnt = 0
    key_x = keys[depth - 1]
    indent = get_indent(depth)
    prefix = (
        ("{}:\r\n" + indent + "{{\r\n")
        if type_dict[key_x] in (INT, FLOAT)
        else ('"{}":\r\n' + indent + "{{\r\n")
    )
    suffix_comma = "},\r\n"
    suffix_end = "}\r\n"

    prefix = indent + prefix
    suffix_comma = indent + suffix_comma
    suffix_end = indent + suffix_end

    for key, value in data.items():
        outfp.write(prefix.format(key))
        if depth == len(keys):
            write_to_gd_row(value, type_dict, outfp, depth + 1)
        else:
            write_to_gd_key(value, keys, type_dict, outfp, depth + 1)
        cnt += 1
        outfp.write(suffix_end if cnt == len(data) else suffix_comma)


def write_to_gd_row(row, type_dict, outfp, depth):
    """Write to GDScript. Row style sheet."""
    cnt = 0
    indent = get_indent(depth)
    template = '{}"{}": {}'
    for key, value in row.items():
        if type_dict[key] == INT:
            outfp.write(template.format(indent, key, get_int(value)))
        elif type_dict[key] == FLOAT:
            outfp.write(template.format(indent, key, get_float(value)))
        elif type_dict[key] == STRING:
            outfp.write(template.format(indent, key, get_string(value)))
        elif type_dict[key] == BOOL:
            outfp.write(template.format(indent, key, get_bool(value)))
        elif type_dict[key] == INT_ARR:
            outfp.write(template.format(indent, key, get_int_arr(value)))
        elif type_dict[key] == FLOAT_ARR:
            outfp.write(template.format(indent, key, get_float_arr(value)))
        elif type_dict[key] == STRING_ARR:
            outfp.write(template.format(indent, key, get_string_arr(value)))
        elif type_dict[key] == BOOL_ARR:
            outfp.write(template.format(indent, key, get_bool_arr(value)))
        elif type_dict[key] == VECTOR2:
            outfp.write(template.format(indent, key, get_vector2(value)))
        elif type_dict[key] == VECTOR3:
            outfp.write(template.format(indent, key, get_vector3(value)))
        elif type_dict[key] == COLOR:
            outfp.write(template.format(indent, key, get_color(value)))
        elif type_dict[key] == GDSCRIPT:
            outfp.write(template.format(indent, key, get_gd(value)))
        elif type_dict[key] == TRANSLATE:
            outfp.write(template.format(indent, key, get_translate(value)))
        else:
            outfp.close()
            raise RuntimeError(f'key "{key}" type "{type_dict[key]}" is wrong')

        cnt += 1
        if cnt == len(row):
            outfp.write("\r\n")
        else:
            outfp.write(",\r\n")


def write_to_gd_kv(data, keys, type_dict, outfp, depth):
    """Write to GDScript. Key-value style sheet."""
    cnt = 0
    key_x = keys[depth - 1]
    indent = get_indent(depth)
    prefix = "[{}]: " if type_dict[key_x] in (INT, FLOAT) else "{}: "
    suffix_comma = ",\r\n"
    suffix_end = "\r\n"

    prefix = indent + prefix

    for _, kv in data.items():
        key, value = None, None
        for k, v in kv.items():
            if type_dict[k] == INT and k.lower() == "key":
                key = get_int(v)
            elif type_dict[k] == FLOAT and k.lower() == "key":
                key = get_float(v)
            elif type_dict[k] == STRING and k.lower() == "key":
                key = get_string(v)
            elif type_dict[k] == STRING and k.lower() == "value":
                value = get_string(v)
            else:
                raise RuntimeError("kv excel format is wrong")

        if not (key and value):
            outfp.close()
            raise RuntimeError("kv excel format is wrong")

        outfp.write(prefix.format(key))
        outfp.write(value)
        cnt += 1
        outfp.write(suffix_end if cnt == len(data) else suffix_comma)


def write_to_csv(sheet, sheet_name, output_csv_path, xls_file):
    """Export to CSV."""
    csv_file_name = OUTPUT_CSV_NAME_TEMPLATE.format(sheet_name=sheet_name)
    csv_file_fullpath = output_csv_path + "/" + csv_file_name
    filenames = ["id", DEFAULT_LANG]
    data_csv = {}

    if os.path.isfile(csv_file_fullpath):
        # read old csv
        with open(csv_file_fullpath, encoding="utf-8", newline="") as f:
            r = csv.DictReader(f)
            filenames = r.fieldnames
            for row in r:
                data_csv[row["id"]] = row

    # update data_csv
    for key, value in sheet.items():
        if data_csv.get(key) is None:
            data_csv[key] = {}
        data_csv[key]["id"] = key
        data_csv[key][DEFAULT_LANG] = value

    # write new csv
    with open(csv_file_fullpath, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, filenames)
        w.writeheader()
        for row in data_csv.values():
            w.writerow(row)
        log(SUCCESS, f"[{GD_CNT:02d}] {xls_file:{MAX_XLS_NAME_LEN}} => {csv_file_name}")


def get_indent(depth):
    """Get indent."""
    indent = ""
    for _ in range(depth):
        indent += "\t"
    return indent


def check_config():
    """Check config file."""
    if not os.path.isfile(CONFIG_FILE):
        default_config = {
            "input_folder": INPUT_FOLDER,
            "output_gd_folder": OUTPUT_GD_FOLDER,
            "output_gd_name_template": OUTPUT_GD_NAME_TEMPLATE,
            "output_csv_folder": OUTPUT_CSV_FOLDER,
            "output_csv_name_template": OUTPUT_CSV_NAME_TEMPLATE,
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as json_file:
            json_file.write(json.dumps(default_config, indent=True))
            json_file.close()
            subprocess.check_call(["attrib", "+H", CONFIG_FILE])
            log(INFO, f"generate config at {CONFIG_FILE}")


def load_config():
    """Load config file."""
    check_config()
    log(INFO, f"load config from \t{CONFIG_FILE}")

    with open(CONFIG_FILE, encoding="utf-8") as json_file:
        config = json.load(json_file)
        global INPUT_FOLDER, OUTPUT_GD_FOLDER, OUTPUT_GD_NAME_TEMPLATE, OUTPUT_CSV_FOLDER, OUTPUT_CSV_NAME_TEMPLATE
        INPUT_FOLDER = config["input_folder"]
        OUTPUT_GD_FOLDER = config["output_gd_folder"]
        OUTPUT_GD_NAME_TEMPLATE = config["output_gd_name_template"]
        OUTPUT_CSV_FOLDER = config["output_csv_folder"]
        OUTPUT_CSV_NAME_TEMPLATE = config["output_csv_name_template"]
        json_file.close()


def save_config():
    """Save config file."""
    if not os.path.isfile(CONFIG_FILE):
        return

    config = {
        "input_folder": INPUT_FOLDER,
        "output_gd_folder": OUTPUT_GD_FOLDER,
        "output_gd_name_template": OUTPUT_GD_NAME_TEMPLATE,
        "output_csv_folder": OUTPUT_CSV_FOLDER,
        "output_csv_name_template": OUTPUT_CSV_NAME_TEMPLATE,
    }
    with open(CONFIG_FILE, "r+", encoding="utf-8") as json_file:
        json_file.truncate(0)  # need '0' when using r+
        json_file.write(json.dumps(config, indent=True))
        json_file.close()
        log(INFO, f"save config at {CONFIG_FILE}")


def main():
    """Main function."""
    global GD_CNT
    GD_CNT = 0
    input_path = INPUT_FOLDER
    output_gd_path = OUTPUT_GD_FOLDER
    output_csv_path = OUTPUT_CSV_FOLDER
    log(INFO, f"input path: \t{input_path}")
    log(INFO, f"output *.gd path: \t{output_gd_path}")
    log(INFO, f"output *.csv path: \t{output_csv_path}")
    if not os.path.exists(input_path):
        raise RuntimeError("input path does NOT exist.")
    if not os.path.exists(output_gd_path):
        os.mkdir(output_gd_path)
        log(INFO, f"make a new dir: \t{output_gd_path}")
    if not os.path.exists(output_csv_path):
        os.mkdir(output_csv_path)
        log(INFO, f"make a new dir: \t{output_csv_path}")

    xls_files = os.listdir(input_path)
    if len(xls_files) == 0:
        raise RuntimeError("input dir is empty.")

    # find max string len
    global MAX_XLS_NAME_LEN
    MAX_XLS_NAME_LEN = len(max(xls_files, key=len))

    # filer files by .xls
    xls_files = [
        x
        for x in xls_files
        if (x[-4:] in [".xls"] or x[-5:] in [".xlsm", ".xlsx"]) and x[0:2] not in ["~$"]
    ]
    log(INFO, f"total XLS: \t\t{len(xls_files)}")

    for _, xls_file in enumerate(xls_files):
        t, ret, err_str = make_table(f"{INPUT_FOLDER}/{xls_file}")
        if ret != 0:
            GD_CNT += 1
            log(FAILED, f"[{GD_CNT:02d}] {xls_file}")
            raise RuntimeError(err_str)
        # print(json.dumps(t, indent=4))
        write_to_gd_script(t, output_gd_path, output_csv_path, xls_file)


def run():
    """Function entry."""
    # print command line arguments
    for arg in sys.argv[1:]:
        if arg == "-c":
            global IS_COLOR
            IS_COLOR = True

    try:
        log(INFO, f"time: \t\t{datetime.datetime.now()}")
        load_config()
        main()
        log(INFO, f"total GDScript: \t\t{GD_CNT}")
        log(INFO, "done.")
        # log(INFO, 'press Enter to exit...')
        # input()
    except (
        RuntimeError,
        ValueError,
        SyntaxError,
        AssertionError,
        PermissionError,
    ) as err:
        err_type = str(type(err))
        err_type = re.findall(r"<class \'(.+?)\'>", err_type)[0]
        err_str = f"[{err_type + str(err)}] "
        log(ERROR, err_str)
        # log(INFO, 'check error please...')
        # input()


def set_gui(frame):
    """Set GUI."""
    if frame is not None:
        global GUI, INFO, ERROR, SUCCESS, FAILED
        GUI = frame
        INFO = "info"
        ERROR = "error"
        SUCCESS = "sucess"
        FAILED = "failed"
    else:
        log(ERROR, "frame is None.")


def log(prefix, s):
    """Print logs."""
    if GUI is not None:
        GUI.write(prefix, s)
    elif IS_COLOR:
        print(f'[{prefix["c"]}] {s}')
    else:
        print(f'[{prefix["b"]}] {s}')


if __name__ == "__main__":
    run()
