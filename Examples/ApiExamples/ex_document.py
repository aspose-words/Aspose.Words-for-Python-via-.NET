# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
import re
import traceback
import unittest
import io
import os
import glob
import xml.etree.ElementTree as xml_tree
from asyncio.log import logger
from urllib.request import urlopen, Request
from datetime import datetime, timedelta, timezone
import requests

import aspose.words as aw
import aspose.pydrawing as drawing
from aspose.words import ControlChar

from io import BytesIO
import uuid
import asyncio
import websockets

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR, FONTS_DIR, GOLDS_DIR
from document_helper import DocumentHelper

class ExDocument(ApiExampleBase):

    def test_comments(self):
        self.add_comments_sae_drug_name("S101001.rtf", [{"感染性肺炎": {"该事件结局为": "插入批注的内容"}}, {"窦性心动过缓": {"针对该事件的治疗药物包括": "插入批注的内容"}}])

    def add_comments_sae_drug_name(self, file_path, para_text_list):
        doc = aw.Document(MY_DIR + file_path)
        # 查找文档中的特定文本并添加批注
        for item in para_text_list:
            for key, value in item.items():
                for k, v in value.items():
                    is_comment_added = False
                    opt = aw.replacing.FindReplaceOptions()
                    opt.use_substitutions = True
                    doc.range.replace(k, "$0", opt)
                    comment = aw.Comment(doc, '', "", datetime.now())
                    comment.set_text(v)
                    # 创建批注范围
                    comment_start = aw.CommentRangeStart(doc, comment.id)
                    comment_end = aw.CommentRangeEnd(doc, comment.id)
                    # 查找文档中的特定文本并添加批注
                    for top_run in doc.get_child_nodes(aw.NodeType.RUN, True):
                        top_run = top_run.as_run()
                        if top_run.text == key and top_run.font.bold is True:
                            top_para = top_run.parent_paragraph
                            if top_para.get_ancestor(aw.NodeType.HEADER_FOOTER):
                                continue
                            next_para = top_para.next_sibling
                            while next_para is not None:
                                next_para = next_para.as_paragraph()
                                for run in next_para.runs:
                                    run = run.as_run()
                                    if run.text == k:
                                        # 插入批注范围和批注
                                        next_para.insert_before(comment_start, run)
                                        next_para.insert_after(comment_end, run)
                                        next_para.insert_after(comment, run)
                                        is_comment_added = True
                                        break

                                if is_comment_added:
                                    break
                                next_para = next_para.next_sibling

                            if not is_comment_added:
                                next_para = top_para.next_sibling.as_paragraph()
                                i = 0
                                while next_para.runs[i] is not None:
                                    current_run = next_para.runs[i]
                                    i += 1

                                after_run = current_run.clone(True)
                                current_run.parent_node.insert_after(after_run, current_run)
                                after_run = after_run.as_run()
                                after_run.text = current_run.text[-1]
                                current_run.text = current_run.text[0:len(current_run.text) - 1]

                                next_para.insert_before(comment_start, after_run)
                                next_para.insert_after(comment_end, after_run)
                                next_para.insert_after(comment, after_run)

        # 保存文档
        doc.save(ARTIFACTS_DIR + file_path)
        return file_path

    def test_table_style(self):
        doc = aw.Document(MY_DIR + "merged2.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True)
        table = table.as_table()

        table.clear_borders()

        for row in table.rows:
            row = row.as_row()
            if row.is_first_row:
                for cell in row.cells:
                    cell = cell.as_cell()
                    cell.cell_format.borders.top.line_style = aw.LineStyle.SINGLE
                    cell.cell_format.borders.bottom.line_style = aw.LineStyle.SINGLE
                    cell.cell_format.shading.background_pattern_color = drawing.Color.blue

            if row.is_last_row:
                for cell in row.cells:
                    cell = cell.as_cell()
                    cell.cell_format.borders.bottom.line_style = aw.LineStyle.SINGLE

        doc.save(ARTIFACTS_DIR + "output.docx")

    def test_memory(self):
        text_to_find = "呼吸频率"
        doc = aw.Document(MY_DIR + "a.rtf")

        # 查找文档中的特定文本并添加批注
        opt = aw.replacing.FindReplaceOptions()
        opt.use_substitutions = True
        doc.range.replace(text_to_find, "$0", opt)

        # 创建一个批注
        comment = aw.Comment(doc, "Author", "", datetime.now())
        comment.set_text("This is a comment")

        for find_para_run in doc.get_child_nodes(aw.NodeType.RUN, True):
            find_para_run = find_para_run.as_run()
            if find_para_run.text == "窦性心动过缓":
                for run in find_para_run.parent_paragraph.runs:
                    run = run.as_run()
                    if run.text == text_to_find:
                        # 创建批注范围
                        comment_start = aw.CommentRangeStart(doc, comment.id)
                        comment_end = aw.CommentRangeEnd(doc, comment.id)
                        # 插入批注范围和批注
                        paragraph = run.parent_paragraph
                        paragraph.insert_before(comment_start, run)
                        paragraph.insert_after(comment_end, run)
                        paragraph.insert_after(comment, run)
                        break

        # 保存文档
        doc.save(ARTIFACTS_DIR + "output.rtf")

    def add_comments_utils(file_path, text_to_find, para_text, comment_text):

        doc = aw.Document(file_path)
        # 查找文档中的特定文本并添加批注
        opt = aw.replacing.FindReplaceOptions()
        opt.use_substitutions = True
        doc.range.replace(text_to_find, "$0", opt)
        # 创建一个批注
        table = doc.get_child_nodes(aw.NodeType.TABLE, True)[3].as_table()
        for row in table.rows:
            for cell in row.as_row():
                cell = cell.as_cell()
                for paragraph in cell.paragraphs:
                    paragraph = paragraph.as_paragraph()
                    for run in paragraph.runs:
                        run = run.as_run()
                        run.font.name = "Times New Roman"  # 设置西文是新罗马字体
                        run.font.name_far_east = "宋体"
                        run.font.size = 8
        comment = aw.Comment(doc, '', "", date)
        comment.set_text(comment_text)

        # 查找文档中的特定文本并添加批注
        for find_para_run in doc.get_child_nodes(aw.NodeType.RUN, True):
            find_para_run = find_para_run.as_run()
            if find_para_run.text == para_text:
                for run in find_para_run.parent_paragraph.runs:
                    run = run.as_run()
                    if run.text == text_to_find:
                        # 创建批注范围
                        comment_start = aw.CommentRangeStart(doc, comment.id)
                        comment_end = aw.CommentRangeEnd(doc, comment.id)
                        # 插入批注范围和批注
                        paragraph = run.parent_paragraph
                        paragraph.insert_before(comment_start, run)
                        paragraph.insert_after(comment_end, run)
                        paragraph.insert_after(comment, run)
                        break

        # 保存文档
        doc.save(file_path)
        return file_path

    def test_ported_code(self):

        doc = aw.Document(MY_DIR + "Part.docx")
        all_numbers = []

        for para in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            para_text = para.to_string(aw.SaveFormat.TEXT)
            pattern = re.compile(r"(?:\d*\.*\d+)")
            numbers = list(re.findall(pattern, para_text))
            for number in numbers:
                all_numbers.append(number)

        all_numbers = list(dict.fromkeys(all_numbers))

        opt = aw.replacing.FindReplaceOptions()
        opt.match_case = True
        opt.find_whole_words_only = True
        for number in all_numbers:
            doc.range.replace(number, number, opt)

        for run in doc.get_child_nodes(aw.NodeType.RUN, True):
            run = run.as_run()
            if any(run.text in s for s in all_numbers):
                run.font.bold = True

        doc.save(ARTIFACTS_DIR + "output.docx")

    def test_part(self):
        fields_list = []

        doc = aw.Document(MY_DIR + "Part - Copy.docx")
        builder = aw.DocumentBuilder(doc)

        fields = doc.get_child_nodes(aw.NodeType.FIELD_START, True)
        for field in fields:
            field = field.as_field_start()
            if field.field_type == aw.fields.FieldType.FIELD_TOC:
                toc = field.get_field().get_field_code()
                fields_list.append(toc)

        builder.move_to_document_end()
        for field_list in fields_list:
            builder.insert_field(field_list)
            builder.writeln()

        doc.save(ARTIFACTS_DIR + "123.docx")

    def aw_read_table_id(self, table=None, id=None):
        table.convert_to_horizontally_merged_cells()

        table_data = []
        for row in table.rows:
            content = {
                "type": "tableRow",
                "content": []
            }
            row_index = table.index_of(row)
            cell = row.as_row().first_cell
            row_span = 1
            col_span = 1
            current_cell = cell
            cell_index = 0
            cell_text = ""
            while current_cell is not None:
                cell_index = current_cell.parent_row.index_of(current_cell)
                if current_cell.cell_format.vertical_merge == aw.tables.CellMerge.FIRST and current_cell.cell_format.horizontal_merge == aw.tables.CellMerge.FIRST:
                    cell_text = current_cell.get_text()
                    current_cell = current_cell.next_cell
                    for i in range(row_index, table.rows.count):
                        if table.rows[i].cells[cell_index].cell_format.vertical_merge == aw.tables.CellMerge.PREVIOUS:
                            row_span += 1
                    while current_cell is not None and current_cell.cell_format.horizontal_merge == aw.tables.CellMerge.PREVIOUS:
                        col_span = col_span + 1
                        current_cell = current_cell.next_cell
                elif current_cell.cell_format.horizontal_merge == aw.tables.CellMerge.FIRST:
                    cell_text = current_cell.get_text()
                    current_cell = current_cell.next_cell
                    while current_cell is not None and current_cell.cell_format.horizontal_merge == aw.tables.CellMerge.PREVIOUS:
                        col_span = col_span + 1
                        current_cell = current_cell.next_cell
                elif current_cell.cell_format.vertical_merge == aw.tables.CellMerge.FIRST:
                    cell_text = current_cell.get_text()
                    cell_index = current_cell.parent_row.index_of(current_cell)
                    for i in range(row_index, table.rows.count):
                        if table.rows[i].cells[cell_index].cell_format.vertical_merge == aw.tables.CellMerge.PREVIOUS:
                            row_span += 1
                    current_cell = current_cell.next_cell
                else:

                    cell_text = current_cell.get_text()
                    current_cell = current_cell.next_cell

                cell_content = {
                    "type": "tableCell",
                    "attrs": {
                        "colspan": col_span,
                        "rowspan": row_span,
                        "colwidth": None
                    },
                    "content": []
                }

                paragraph = {
                    "type": "paragraph",
                    "content": [
                        {
                            "type": "text",
                            "text": cell_text,
                        }
                    ]
                }
                cell_content["content"].append(paragraph)
                content["content"].append(cell_content)

                col_span = 1
                row_span = 1

            table_data.append(content)

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()  # 开始一个新表格
        for i in range(len(table_data)):
            cell_data_index = 0
            for row_data in table_data[i].get('content', []):
                # 遍历表格单元格
                for cell_data in row_data.get('content', []):
                    new_cell = builder.insert_cell()  # 插入新的表格行
                    new_cell.cell_format.borders.line_style = aw.LineStyle.DASH_SMALL_GAP
                    builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE

                    row_index = table.index_of(new_cell)
                    cell_index = new_cell.parent_row.index_of(new_cell)

                    previous_row = table.rows[row_index - 1]
                    if previous_row is not None:
                        previous_cell = previous_row.cells[cell_index]
                        if previous_cell is not None and \
                                (previous_cell.cell_format.vertical_merge == aw.tables.CellMerge.FIRST or
                                 previous_cell.cell_format.vertical_merge == aw.tables.CellMerge.PREVIOUS):
                            for y in range(i, -1, -1):
                                curr_row_span = table_data[y].get('content', [])[cell_data_index].get('attrs', {}).get(
                                    'rowspan', 1)
                                if curr_row_span > 1:
                                    if curr_row_span >= table.rows.count - y:
                                        builder.cell_format.vertical_merge = aw.tables.CellMerge.PREVIOUS
                                    else:
                                        builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE

                    builder.cell_format.horizontal_merge = aw.tables.CellMerge.NONE
                    cell_content = ''  # 单元格内容初始化为空字符串
                    if isinstance(cell_data, dict):  # 检查单元格内容是否为字典类型
                        for paragraph_content in cell_data.get('content', []):
                            text = paragraph_content["text"]  # 获取段落文本内容
                            cell_content += text  # 将段落内容添加到单元格内容中
                    else:
                        cell_content = cell_data  # 如果不是字典类型，则直接使用内容

                    new_cell.cell_format.wrap_text = True  # 设置单元格内容自动换行
                    new_cell.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER  # 设置单元格内容垂直居中

                    builder.write(cell_content)  # 写入单元格内容

                    colspan = row_data.get('attrs', {}).get('colspan', 1)  # 获取列合并数
                    rowspan = row_data.get('attrs', {}).get('rowspan', 1)  # 获取行合并数

                    if colspan > 1 and rowspan > 1:
                        builder.cell_format.vertical_merge = aw.tables.CellMerge.FIRST
                        builder.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST  # 水平合并第一个单元格
                        for i in range(1, colspan):
                            builder.insert_cell()
                            builder.cell_format.horizontal_merge = aw.tables.CellMerge.PREVIOUS
                    elif colspan > 1:
                        builder.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST  # 水平合并第一个单元格
                        for i in range(1, colspan):
                            builder.insert_cell()
                            builder.cell_format.horizontal_merge = aw.tables.CellMerge.PREVIOUS
                    elif rowspan > 1:
                        builder.cell_format.vertical_merge = aw.tables.CellMerge.FIRST  # 垂直合并第一个单元格

                    cell_data_index += 1

            builder.end_row()
        builder.end_table()

        builder.document.save(ARTIFACTS_DIR + "output.docx")

        return_data = {
            "type": "table",
            "attrs": {
                "id": id
            },
            "content": table_data
        }

        return return_data

    def test_another_1(self):
        doc = aw.Document(MY_DIR + "ano.docx")
        # doc.unlink_fields()
        # doc.save(ARTIFACTS_DIR + "unlink.docx")
        current_level = 0
        data = []
        stack = []
        for s in doc.sections:
            sect = s.as_section()
            for node in sect.body.get_child_nodes(aw.NodeType.ANY, True):
                block_id = uuid.uuid4()
                block_id1 = uuid.uuid4()
                if node.node_type == aw.NodeType.PARAGRAPH:
                    node = node.as_paragraph()
                    if not node.get_ancestor(aw.NodeType.TABLE):
                        if node.paragraph_format.outline_level in [0, 1, 2, 3, 4, 5]:
                            level = int(node.paragraph_format.outline_level) + 1
                            if level > current_level:

                                # 如果级别更深，将当前标题添加到堆栈
                                stack.append((current_level, data))
                                data = []
                                current_level = level
                            elif level < current_level:
                                # 如果级别更浅，将堆栈中的项添加回数据
                                while stack and stack[-1][0] >= level:
                                    old_level, old_data = stack.pop()
                                    data = old_data + data
                                    current_level = old_level

                            label = ''
                            if node.list_format.is_list_item:
                                label = node.list_label.label_string

                            builder = []
                            for child in node.get_child_nodes(aw.NodeType.ANY, True):
                                if child.node_type is not aw.NodeType.COMMENT or aw.NodeType.COMMENT_RANGE_START or aw.NodeType.COMMENT_RANGE_END:
                                    builder.append(child.to_string(aw.SaveFormat.TEXT))
                            result = ''.join(builder)
                            text_without_comments = result

                            data.append(
                                {
                                    "Title": label + text_without_comments if label else text_without_comments,
                                    "block_id": str(block_id),
                                    "Content": [],
                                    "Level": level,
                                    "Table": [],
                                    "Tbale_name": [],
                                }
                            )
                        else:
                            if data:
                                if node.node_type is not aw.NodeType.TABLE:
                                    if node.get_text().strip() and "Toc" not in node.get_text()\
                                            and "TOC" not in node.get_text():
                                        data[-1]["Content"].append(
                                            {"type": "text",
                                             "content": node.get_text().strip().replace("  SEQ 表 \* ARABIC ",
                                                                                        '').replace(
                                                 'TOC \h \c "表" HYPERLINK \l "_Toc14741"', '').replace(
                                                 "\u0013 SEQ 图 \\* ARABIC \u00141\u0015 ", ''),
                                             "block_id": data[-1]["block_id"] + '&&' + str(block_id1),
                                             "parent_block_id": data[-1]["block_id"]})
                if data:
                    if node.node_type == aw.NodeType.FIELD_START:
                        field = node.as_field_start()
                        if field.field_type == aw.fields.FieldType.FIELD_TOC:
                            results = field.get_field().display_result.split('\r')
                            for result in results:
                                if result.strip():
                                    data[-1]["Content"].append(
                                        {"type": "toc",
                                         "content": result,
                                         "block_id": data[-1]["block_id"] + '&&' + str(block_id1),
                                         "parent_block_id": data[-1]["block_id"]}
                                    )

                    if node.node_type == aw.NodeType.SHAPE:
                        shape = node.as_shape()
                        if shape.has_image:
                            image_id = str(uuid.uuid1())  # 长度是36
                            try:
                                image_extension = aw.FileFormatUtil.image_type_to_extension(shape.image_data.image_type)
                                image_file_name = f"{image_id}{image_extension}"
                                image_path = os.path.join(IMAGE_DIR, image_file_name)
                                shape.image_data.save(image_path)
                                data[-1]["Content"].append(
                                    {"type": "image",
                                     "content": [{"type": "image", "attrs": {
                                         "src": image_file_name,
                                         "alt": "tips", "title": ''}}],
                                     "block_id": data[-1]["block_id"] + '&&' + str(block_id1),
                                     "parent_block_id": data[-1]["block_id"]}
                                )
                            except Exception as e:
                                # 捕获并处理无法转换图像类型的错误
                                print(f"Error saving image: {e}. Skipping this image.")
                                continue

                    if node.node_type == aw.NodeType.TABLE:
                        parent_node = node.as_table()
                        _able_content = self.aw_read_table_id(parent_node, data[-1]["block_id"] + '&&' + str(block_id1))
                        data[-1]["Content"].append(
                            {"type": "table",
                             "content": _able_content,
                             "block_id": data[-1]["block_id"] + '&&' + str(block_id1),
                             "parent_block_id": data[-1]["block_id"]})
        while stack:
            old_level, old_data = stack.pop()
            data = old_data + data

        sections = [
                    {
                        "type": "text",
                        "content": "問題ない",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&bea3bf57-7148-4e0b-8cd6-42b865a245dc",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "1 标题页\t1",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&3c8d0bc7-25aa-4d87-b623-06e50e0ef5f3",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "2 概要\t5",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0a7ce8ba-b098-49dc-9ce8-4b49ec04dfd1",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "3 个例临床研究报告目录\t22",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&d583e100-02f6-4dfd-be7b-77f803e65904",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "表格目录\t25",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0755dd8e-6268-448c-ab70-754af8639dae",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "图表目录\t26",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&26882db7-0ec2-4a04-9945-34146f4bf30e",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "4 缩略语和术语定义表\t28",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&435a9c72-bf15-4810-ba83-81947019aa30",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "5 伦理学\t29",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&b6a21d17-4200-4b0e-b6fa-f39cd75669d2",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "5.1 独立伦理委员会（IEC）\t29",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&bfca61a1-b799-47be-a638-68adac1601bc",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "5.2 研究的伦理行为\t29",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&6ef71dc8-9f4f-414b-80a3-d13e739adf67",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "5.3 受试者知情与同意\t29",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&ebb73cdc-6fc4-45d4-af14-40b6a0ee58c3",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "6 研究者和研究管理机构\t30",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&2fb21742-011e-4435-9674-e3a528d301d7",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "7 简介\t31",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&679dc3d7-adb1-4e82-aab8-7f04d284ca57",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "8 研究目标\t32",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&33d52b14-6f78-460a-8b14-aedb601e6dc0",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "8.1 主要研究目的\t32",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&87d93c87-3806-4729-9703-4070aa528a76",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "8.2 次要研究目的\t32",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&954b3056-bbce-4e0a-8699-5b383cd78aa7",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "8.3 探索性研究目的\t32",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&db4eee5a-a68a-45ac-a922-6b99a377df5d",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9 研究计划\t32",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&2cd1a61e-650e-4d55-8cd2-21c7bdbb257b",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.1 整体研究计划\t32",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&4b6d8c39-dbee-4b77-8264-f2c636476428",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.1.1 研究示意图\t33",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&a23231a2-9b0f-4f9f-a6a0-266e13fd3d07",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.2 研究设计讨论\t33",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&6d79a26d-3fb8-4fe1-969c-5eb590475de4",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.2.1患者群体的选择\t33",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&b7e17716-1222-4591-bad6-23dbda4ac4d5",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.2.2主要终点采用IRC评估的ORR\t34",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&97a6bf77-77f0-41a7-a08a-be1ec3468fe4",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.2.3 疗效评估标准的选择\t34",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&5db98a67-6a74-4f5f-a680-c81e8c62d2c2",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.3 研究人群的选择\t35",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&dfaec5d9-e7a5-41da-afb2-a2b0ce0d514a",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.3.1 入选标准\t35",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&48b8df09-e00d-4f37-846c-b3b5b4553738",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.3.2 排除标准\t36",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&c7d92beb-dadc-4a75-bfc5-62e19133d600",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.3.3 从治疗或评估中移除受试者\t37",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&e2b00cee-09c9-4ce5-a30e-33948be45b75",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4 治疗\t38",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&7b1b13cb-e13c-4a91-83e4-5ef3fe002dd9",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.1 给予的治疗\t38",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&46a2a95b-6512-4077-991f-d0b5c0ff5431",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.2 研究药物信息\t39",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&553f59f9-3dcf-4850-9a06-1811d34ede64",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.3 受试者的治疗组分配方法\t39",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&09a10fa7-682f-4e83-b95c-d64e3a9f1751",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.4 研究中的剂量选择\t39",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&65253004-e5bc-4c59-aae4-192efae400d9",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.5 每位受试者的研究剂量和给药时间\t39",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&afbbf1bf-b6b6-41cf-aef7-5d0155a9831d",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.6 盲法\t40",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&47885289-78df-4d42-a00e-8d0effad819c",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.7 既往和伴随治疗\t40",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&ff4bac88-19b5-455d-b09e-e758b8d683a3",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.4.8 治疗依从性\t41",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&598baca5-d07f-4946-87cc-d4e67210cb44",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.5 疗效、安全性和药代动力学终点\t41",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&f895bfef-5180-4bd2-be3a-38c1d8192d4b",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.5.1 评估疗效、安全性和药代动力学终点指标和流程图\t41",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&c6f902f4-3983-4c99-806f-5c0a1a856fc5",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.5.2 衡量指标的适当性\t45",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&be40857d-f2a4-432d-a6b7-6698f88bc39e",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.5.3 主要疗效终点\t45",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&8e6610dc-0738-48c0-be14-7b406b49058f",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.5.4 药物浓度测定\t45",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&cbac3ade-7435-4dc7-885d-e77ab9ffcc67",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.6 数据质量保证\t45",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0306e17f-0994-4ef5-ba98-235edd0b7605",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.6.1临床试验过程的质量保证与控制\t46",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&66817d8c-143e-4b3f-a546-592caf59d4b3",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.6.2 启动访视\t47",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&7b61e7f8-6e42-44d4-8134-f0604a5110fd",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.6.3 实验室认可\t47",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&7e97e67b-a584-475b-8481-43218b8fffe7",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.6.4 数据管理\t47",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&689dc981-18c0-46d9-b4e9-a321740d3bfd",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.7 研究方案中计划的统计方法和样本量的确定\t48",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&3b2bb483-8cdc-4454-8e90-9b0b1e89ce26",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.7.1 统计分析计划\t48",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&9acbaa49-4778-44f0-8250-2c6b07914563",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.7.2 样本量的确定\t53",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&138a2915-2469-46a0-b8ec-fa27616d4f0c",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.8 研究过程或分析计划的变更\t53",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&96a99bd2-a5b6-4cb0-bae1-e42c5cc9bc7b",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.8.1 方案变更\t53",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&fc7f5926-c08a-451d-b101-17c6506abec9",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "9.8.2 统计分析计划（SAP）变更\t54",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&40cb74a3-4ff3-45dc-a257-128610b4241b",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10 研究对象\t55",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0c125f3f-ff13-42ce-87ca-3b1e16d29221",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.1 受试者分布\t55",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&591bdf8b-7874-4c6b-83f2-221f436eed14",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.2 研究方案偏离\t56",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&3bad5775-63be-4bcb-9d8e-1a3b488ee2c2",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.3 分析数据集\t56",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&21347ada-b39f-4e18-9cf9-9381e1d4be30",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.4 人口统计学和其他基线特征\t56",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&bca3d133-d571-4f91-92a4-5deb91c274f5",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.4.1 人口统计学和基线特征（mITT集）\t56",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&850912ba-6f6f-4a3f-b0ef-52aa75528f37",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.4.2肿瘤基线特征及既往抗肿瘤治疗（mITT集）\t58",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&5f43477f-9bf8-4de8-824d-886b0564d330",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.4.3 既往病史\t60",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&39f796dc-6569-4743-b1f2-2e7be7106850",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.4.4 既往和合并药物治疗\t61",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&efc3503e-c7d9-42b9-a69c-9a31de9172f9",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "10.5 治疗依从性的测量\t61",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&e3a11541-fc5f-4ed8-b383-28e32921266a",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11疗效评估和药代动力学评估\t61",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0340b7ae-93d0-4cb9-96e2-1e4e3cc280a1",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1疗效结果和个体受试者数据列表\t61",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&c97713ad-ed83-4675-b9b6-44e7985fdb6b",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1.1 疗效分析\t62",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&98b6506b-28db-4cbf-a5f9-2a808ab21f14",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1.2 统计/分析内容\t69",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&f52d1654-bdbe-453c-b852-b3d219905eeb",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1.3 个体疗效数据列表\t72",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&7f32bf3a-dde8-49cd-98e4-606f7ded5346",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1.4 药物剂量、药物浓度以及效应之间的关系\t72",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0359bcee-b771-44bf-9bf6-7eec74397386",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1.5 药物-药物和药物-疾病相互作用\t72",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&ac6c1a29-f0a4-4533-ac4e-4702e1c8a9c6",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1.6 按受试者列出\t72",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0902548d-308a-403b-bc3d-df1acacfce16",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "11.1.7 疗效结论\t72",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&1e5876ca-eb6d-431d-8b89-0a83c5503146",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12 安全性评价\t73",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&0b07599e-bc44-4cad-96f6-072aa142bbc6",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.1 暴露程度\t73",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&5ca91707-720a-4efb-a1f3-1fc36f5dbf95",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.2 不良事件（AE）\t74",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&46ee079c-2870-4471-b3b9-a7d02eebd239",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.2.1 治疗期间发生的不良事件（TEAE）概要\t74",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&f55697b0-d4c8-4bb5-b134-29465d838df2",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.2.2 不良事件列出\t75",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&b9ef66b7-f590-4ea2-9b9a-3d14643d746a",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.2.3不良事件分析\t76",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&ed323f99-574a-4410-af8f-9c5cd02c0b32",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.2.4 各受试者不良事件列表\t76",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&ed581567-46fe-4dc8-9ac4-31405ca46590",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.3 死亡、其他严重不良事件和其他重要不良事件\t77",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&578dde75-475c-44aa-a94c-886729a123b3",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.3.1 死亡、其他严重不良事件和其他重要不良事件列表\t77",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&bcd05c99-c256-489f-8e19-e2d5cda18664",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.3.2 死亡、其他严重不良事件和其他重要不良事件的叙述\t79",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&7a734bd1-06f1-4084-86d4-64d4579f74e8",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.3.3 死亡、其他严重不良事件和其他重要不良事件的分析和讨论\t79",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&26260c11-eca3-4f51-95c7-77ba8da968f1",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.4 临床实验室评估\t79",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&c0146693-6cec-4b4a-8afd-d54e1c032da5",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.4.1 各受试者的个例实验测量值列表（16.2.8）和各异常实验室值（14.3.4）\t79",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&08f03524-061c-4e7c-8c08-6d2b6e6c73c7",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.4.2 各实验室参数的评价\t79",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&1f826b25-1e81-4ba4-8f68-c77e78be0095",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.5 生命体征、体格检查发现和其他安全性相关观察结果（SS）\t81",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&102c3246-11a4-4c57-9458-3f339a27a3f1",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "12.6 安全性结论\t81",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&9b55f3ec-a371-4742-b26b-a644861d49f7",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "13 讨论和总体结论\t81",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&6996ddd9-f10f-4c1b-aeb5-03aa2e7fa590",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "13.1 讨论\t81",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&1c20ae4d-98ff-4124-9663-f45ccbdd1a9c",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "13.1.1 背景\t81",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&18aeb7ec-6ad4-4cf4-ad3f-5671d9b060d7",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "13.1.2 有效性\t82",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&717212d5-e823-474b-a2ad-95f8ea25983b",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "13.1.3 安全性\t83",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&cebd019e-68d6-461b-a070-803c86b99700",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "13.1.4 药代动力学\t83",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&be9325da-5d8a-4276-9783-b4e2b3a24658",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "13.2 结论\t83",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&54b7453d-bf99-4951-bc70-fa34c8074608",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "14 参考但不纳入文本的表格、图示和图表\t84",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&40a0bdb1-90f9-4953-aaa8-9608470514e4",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "15参考文献列表\t85",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&8921ebc4-d42b-4e0a-9e3a-719c2512a8a4",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    },
                    {
                        "type": "toc",
                        "content": "16 附录\t86",
                        "block_id": "90b827cc-b786-49f3-b98c-59ee9329b595&&ad05c740-1d1b-48eb-8fd3-6d396d8cba93",
                        "parent_block_id": "90b827cc-b786-49f3-b98c-59ee9329b595"
                    }
                ]

        doc1 = aw.Document()
        builder = aw.DocumentBuilder(doc1)

        # 遍历节(section)
        for section in sections:
            builder.paragraph_format.clear_formatting()
            if section["type"] == 'title':
                # 设置标题样式
                level = section.get('level', 1)
                builder.paragraph_format.style_identifier = getattr(
                    aw.StyleIdentifier, f"HEADING{level}"
                )
                # 添加标题内容
                new_run = builder.font
                # 设置西文和中文字体
                new_run.size = 12
                new_run.name = "Times New Roman"  # 设置西文是新罗马字体
                new_run.name_far_east = "MS Gothic"
                # 添加标题内容
                title = section.get('contents', '')  # 获取标题内容
                builder.writeln(title)  # 写入标题到文档

            elif section["type"] == 'toc':
                # 添加文本内容
                page_setup = doc.first_section.page_setup
                tab_stop_position = page_setup.page_width - page_setup.left_margin - page_setup.right_margin
                builder.paragraph_format.tab_stops.add(tab_stop_position, aw.TabAlignment.RIGHT, aw.TabLeader.DOTS)
                text_content = section.get('content', '')
                builder.writeln(text_content)

            elif section["type"] == 'text':
                # 添加文本内容
                new_run = builder.font
                # 设置西文和中文字体
                new_run.name = "Times New Roman"  # 设置西文是新罗马字体
                new_run.name_far_east = "MS Gothic"
                new_run.size = 12
                text_content = section.get('content', '')  # 获取文本内容
                builder.paragraph_format.character_unit_first_line_indent = 2
                builder.paragraph_format.line_spacing = 18
                builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
                builder.paragraph_format.line_unit_after = 0
                builder.paragraph_format.space_after = 0
                builder.writeln(text_content)  # 写入文本到文档

            elif section["type"] == 'table':
                # 添加表格内容
                table_data = section.get('contents', {})  # 获取表格内容
                table = builder.start_table()  # 开始一个新表格
                cell_data_index = 0
                # 遍历表格行
                for row_data in table_data.get('content', []):
                    # 遍历表格单元格
                    for cell_data in row_data.get('content', []):
                        new_cell = builder.insert_cell()  # 插入新的表格行
                        builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE

                        row_index = table.index_of(new_cell)
                        cell_index = new_cell.parent_row.index_of(new_cell)

                        previous_row = table.rows[row_index - 1]
                        if previous_row is not None:
                            previous_cell = previous_row.cells[cell_index]
                            if previous_cell is not None and \
                                    (previous_cell.cell_format.vertical_merge == aw.tables.CellMerge.FIRST or
                                     previous_cell.cell_format.vertical_merge == aw.tables.CellMerge.PREVIOUS):
                                for y in range(i, -1, -1):
                                    curr_row_span = table_data[y].get('content', [])[cell_data_index].get('attrs',
                                                                                                          {}).get(
                                        'rowspan', 1)
                                    if curr_row_span > 1:
                                        break
                                if curr_row_span >= table.rows.count:
                                    builder.cell_format.vertical_merge = aw.tables.CellMerge.PREVIOUS
                                else:
                                    builder.cell_format.vertical_merge = aw.tables.CellMerge.NONE

                        builder.cell_format.horizontal_merge = aw.tables.CellMerge.NONE
                        cell_content = ''  # 单元格内容初始化为空字符串
                        if isinstance(cell_data, dict):  # 检查单元格内容是否为字典类型
                            for paragraph_content in cell_data.get('content', []):
                                for paragraph in paragraph_content.get('content', []):
                                    text = paragraph["text"]  # 获取段落文本内容
                                    cell_content += text  # 将段落内容添加到单元格内容中
                        else:
                            cell_content = cell_data  # 如果不是字典类型，则直接使用内容

                        new_cell.cell_format.wrap_text = True  # 设置单元格内容自动换行
                        new_cell.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER  # 设置单元格内容垂直居中

                        builder.write(cell_content)  # 写入单元格内容

                        colspan = row_data.get('attrs', {}).get('colspan', 1)  # 获取列合并数
                        rowspan = row_data.get('attrs', {}).get('rowspan', 1)  # 获取行合并数

                        if colspan > 1 and rowspan > 1:
                            builder.cell_format.vertical_merge = aw.tables.CellMerge.FIRST
                            builder.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST  # 水平合并第一个单元格
                            for i in range(1, colspan):
                                builder.insert_cell()
                                builder.cell_format.horizontal_merge = aw.tables.CellMerge.PREVIOUS
                        elif colspan > 1:
                            builder.cell_format.horizontal_merge = aw.tables.CellMerge.FIRST  # 水平合并第一个单元格
                            for i in range(1, colspan):
                                builder.insert_cell()
                                builder.cell_format.horizontal_merge = aw.tables.CellMerge.PREVIOUS
                        elif rowspan > 1:
                            builder.cell_format.vertical_merge = aw.tables.CellMerge.FIRST  # 垂直合并第一个单元格

                        cell_data_index += 1

                    builder.end_row()
                    builder.end_table()

        # 保存文档
        doc1.save(ARTIFACTS_DIR + "test2.docx")

        return data


    def test_merged_cells(self):
        doc = aw.Document(MY_DIR + "merged.docx")

        with BytesIO() as docToHtml:
            doc.save(docToHtml, aw.SaveFormat.HTML)
            docToHtml.seek(0)
            xr = xml_tree.parse(docToHtml)
            html = xr.getroot()

            # extract the information about table col and rowspan
            table_dfn = xml_tree.Element("tables")
            for tab in html.findall(".//table"):
                table = xml_tree.SubElement(table_dfn, "table")
                for tr in tab.findall(".//tr"):
                    row = xml_tree.SubElement(table, "tr")
                    for td in tr.findall(".//td"):
                        cell = xml_tree.SubElement(row, "td")
                        for attr in td.attrib:
                            if "span" in attr:
                                cell.set(attr, td.attrib[attr])

    def test_tt1(self):
        doc = aw.Document(ARTIFACTS_DIR + "Result.docx")

        list = []
        para = doc.first_section.body.get_child(aw.NodeType.PARAGRAPH, 1, True)
        if "C-Heading" in para.as_paragraph().paragraph_format.style_name:
            current_node = para.next_sibling
            is_continue = True

            while current_node is not None and is_continue:
                if current_node.node_type == aw.NodeType.PARAGRAPH:
                    temp_para = current_node.as_paragraph()
                    if "C-Heading" in temp_para.paragraph_format.style_name:
                        is_continue = False
                        continue

                if current_node.node_type == aw.NodeType.TABLE:
                    list.append(current_node)

                current_node = current_node.next_sibling

        for node in list:
            node.remove()

        doc.save(ARTIFACTS_DIR + "GetTableRemove.docx")

    def test_another(self):
        doc = aw.Document(MY_DIR + "111 (1).docx")
        current_level = 0
        data = []
        stack = []
        paragraph = doc.get_child(aw.NodeType.PARAGRAPH, 4, True)
        print(paragraph.as_paragraph().get_text())
        for s in doc.sections:
            sect = s.as_section()
            for node in sect.body.get_child_nodes(aw.NodeType.ANY, False):
                if node.node_type == aw.NodeType.PARAGRAPH:
                    node = node.as_paragraph()
                    if node.paragraph_format.outline_level in [0, 1, 2, 3, 4, 5]:
                        level = int(node.paragraph_format.outline_level) + 1
                        if level > current_level:
                            # 如果级别更深，将当前标题添加到堆栈
                            stack.append((current_level, data))
                            data = []
                            current_level = level
                        elif level < current_level:
                            # 如果级别更浅，将堆栈中的项添加回数据
                            while stack and stack[-1][0] >= level:
                                old_level, old_data = stack.pop()
                                data = old_data + data
                                current_level = old_level
                        data.append(
                            {
                                "Title": node.get_text(),
                                "Content": [],
                                "Level": level,
                                "Table": [],
                                "Tbale_name": [],
                            }
                        )
                    else:
                        if data:
                            if node.get_text().startswith("表"):
                                data[-1]["Tbale_name"].append(
                                    node.get_text().strip("SEQ \* ARABIC").strip("SEQ")
                                )
                            if (
                                    node.get_text().startswith("表")
                                    or node.get_text().startswith("来源：")
                                    or node.get_text().startswith("图")
                            ):
                                pass
                            else:
                                data[-1]["Content"].append(node.get_text())
                if data:
                    if node.node_type == aw.NodeType.TABLE:
                        parent_node = node.as_table()
                        # able_content = aw_read_table(parent_node)

                        # data[-1]["Table"].append(able_content)
        while stack:
            old_level, old_data = stack.pop()
            data = old_data + data

    @staticmethod
    def find_last_related_paragraph(paragraphs, title_text):
        """
        找到标题后的最后一个相关段落。
        如果标题下面没有段落内容，返回标题本身。
        """
        title_paragraph = None

        # 找到目标标题
        for para in paragraphs:
            para = para.as_paragraph()  # 确保是段落
            para_text = para.get_text().strip()  # 去除空格
            if para_text == title_text:
                title_paragraph = para
                break

        if title_paragraph is None:
            # 返回文档的最后一个段落作为备用
            return paragraphs[-1]

        # 找到标题后的最后一个相关段落
        last_related_paragraph = title_paragraph  # 初始位置是标题

        # 遍历直到找到级别改变的段落
        while (
                last_related_paragraph.next_sibling
                and last_related_paragraph.next_sibling.as_paragraph().paragraph_format.outline_level
                > title_paragraph.paragraph_format.outline_level
        ):
            last_related_paragraph = last_related_paragraph.next_sibling

        return last_related_paragraph
    def test_generate_docx_with_result_laikai(self):
        header_list = ["DISPOSITION OF SUBJECTS"]
        table_list = [
                [
                    {
                        "id": "f32b82df656a4bd9a90121d95ea0d86c",
                        "is_header_and_footer": True,
                    },
                    {
                        "id": "2",
                        "is_header_and_footer": True,
                    },
                ],
            ]
        doc_main = aw.Document(MY_DIR + "111 (1).docx")
        builder_main = aw.DocumentBuilder(doc_main)
        paragraphs = doc_main.get_child_nodes(aw.NodeType.PARAGRAPH, True)
        for para in paragraphs:
            para = para.as_paragraph()
            para_content = para.to_string(aw.SaveFormat.TEXT)
            para_content = para_content.replace("\r", "")
            para_content = para_content.strip()
            if para_content in header_list or para_content.capitalize() in header_list:
                table_header = aw.Paragraph(doc_main)
                table_header.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
                table_header.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
                para = self.find_last_related_paragraph(paragraphs, para_content)
                previous_para = para
                layout_collector = aw.layout.LayoutCollector(doc_main)
                page_index = layout_collector.get_start_page_index(para)
                if page_index > 1:
                    while previous_para is not None:
                        if previous_para.node_type is aw.NodeType.PARAGRAPH \
                                and "C-Heading" in previous_para.as_paragraph().paragraph_format.style_name \
                                and layout_collector.get_start_page_index(previous_para) < page_index:
                            previous_para = previous_para
                            break
                        previous_para = previous_para.previous_pre_order(doc_main)

                    builder_main.move_to(previous_para)
                    builder_main.insert_paragraph()
                    builder_main.paragraph_format.clear_formatting()
                    builder_main.list_format.list = None
                    builder_main.insert_break(aw.BreakType.SECTION_BREAK_CONTINUOUS)
                    builder_main.current_paragraph.remove()

                if para_content in ["DISPOSITION OF SUBJECTS", "APPENDICES", "REFERENCES"]:
                    idx_num = header_list.index(para_content)
                    num_tables = len(table_list[idx_num])
                    logger.info(f"当前表格长度:{num_tables}")

                    for _index, info in enumerate(table_list[idx_num]):
                        table_id = info.get("id", "")
                        is_header_and_footer = info.get("is_header_and_footer", "")

                        # logger.warning(f"table_id in table list:{table_id}")
                        aw.Document(
                            os.path.join(MY_DIR, f"{table_id}.rtf")
                        ).save(os.path.join(ARTIFACTS_DIR, f"{table_id}.docx"))
                        # 判断是否需要页眉页脚 False为删除页眉页脚
                        if not is_header_and_footer:
                            document = aw.Document(
                                os.path.join(ARTIFACTS_DIR, f"{table_id}.docx")
                            )
                            for section in document.sections:
                                section.footer.is_linked_to_previous = True
                                section.header.is_linked_to_previous = True

                            document.save(
                                os.path.join(ARTIFACTS_DIR, f"{table_id}.docx")
                            )

                        doc_rtf = aw.Document(
                            os.path.join(ARTIFACTS_DIR, f"{table_id}.docx")
                        )

                        _tables = doc_rtf.get_child_nodes(aw.NodeType.TABLE, True)
                        _tables = [t for t in _tables]

                        for table in _tables[::-1]:
                            if table.get_ancestor(aw.NodeType.BODY):
                                i = _tables.index(table)
                                temp = _tables[i]
                                _tables[i] = _tables[i - 1]
                                _tables[i - 1] = temp
                        logger.info(f"插入表格: {table_id}")

                        all_tables_count = sum(p.parent_node.node_type == aw.NodeType.BODY for p in _tables)
                        curr_count = 1
                        for table in _tables:  # 将table 信息反向的插入到word文件中。TODO 表格美化
                            section = para.as_paragraph().parent_section
                            page_setup = section.page_setup
                            page_setup.orientation = aw.Orientation.LANDSCAPE
                            if table.get_ancestor(aw.NodeType.BODY):
                                table_index = _tables.index(table)
                                header_table = _tables[table_index - 1].as_table()
                                footer_table = _tables[table_index + 1].as_table()
                                header_table = header_table.clone(True).as_table()
                                footer_table = footer_table.clone(True).as_table()
                                table = table.clone(True).as_table()
                                imported_header = doc_main.import_node(header_table, True)
                                imported_footer = doc_main.import_node(footer_table, True)
                                imported_table = doc_main.import_node(table, True)

                                para = para.parent_node.insert_after(imported_header, para)
                                t1 = para.as_table()
                                for row in t1.rows:
                                    row.as_row().row_format.heading_format = True

                                para = para.parent_node.insert_after(imported_table, para)
                                t1 = para.as_table()
                                t1.rows[0].row_format.heading_format = True
                                t1.rows[1].row_format.heading_format = True

                                para = para.parent_node.insert_after(imported_footer, para)
                                para = para.parent_node.insert_after(aw.Paragraph(doc_main), para)

                                if curr_count < all_tables_count > 1:
                                    builder_main.move_to(para)
                                    builder_main.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
                                    para = builder_main.current_paragraph
                                    curr_count += 1

                                if curr_count == all_tables_count:
                                    builder_main.move_to(para)
                                    builder_main.insert_break(aw.BreakType.LINE_BREAK)

                    builder_main.insert_break(aw.BreakType.SECTION_BREAK_CONTINUOUS)
                    section = builder_main.current_section
                    page_setup = section.page_setup
                    page_setup.orientation = aw.Orientation.PORTRAIT

        doc_main.save(ARTIFACTS_DIR + "Result.docx")


    def test_test(self):

        document = aw.Document(MY_DIR + "t_ae_1(1).docx")
        dst_doc = aw.Document(MY_DIR + "Document.docx")
        paras = dst_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
        collection = document.get_child_nodes(aw.NodeType.TABLE, True)
        is_next_table = False
        combined_table = aw.tables.Table(dst_doc)

        for table in collection:
            if table.parent_node.node_type == aw.NodeType.BODY and not is_next_table:
                combined_table = table.clone(True).as_table()
                is_next_table = True

            if table.parent_node.node_type == aw.NodeType.BODY and is_next_table:
                table = table.as_table()
                if table.first_row.get_text() == combined_table.first_row.get_text():
                    table.rows[0].remove()
                    table.rows[1].remove()
                    table.last_row.remove()
                    second_table = table.clone(True).as_table()

                    while second_table.has_child_nodes:
                        combined_table.rows.add(second_table.first_row)

        cloned_table = combined_table.clone(True).as_table()
        cloned_table.preferred_width = aw.tables.PreferredWidth.from_percent(100)
        new_imported_table = dst_doc.import_node(cloned_table, True, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING).as_table()

        row_to_insert = new_imported_table.first_row.clone(True).as_row()
        cloned_run = row_to_insert.cells[0].as_cell().first_paragraph.runs[0].as_run().clone(True)
        cloned_run.as_run().text = "New text"
        for cell in row_to_insert.cells:
            cell = cell.as_cell()
            cell.remove_all_children()
            cell.ensure_minimum()

        row_to_insert.cells[0].as_cell().first_paragraph.runs.add(cloned_run)
        new_imported_table.rows.insert(0, row_to_insert)

        para = paras[3].as_paragraph()
        para.parent_node.append_child(new_imported_table)

        dst_doc.save(ARTIFACTS_DIR + 'TestSed.docx')

    def test_test_tiaoh(self):
        header_list = ["REFERENCES"]
        doc_main = aw.Document(MY_DIR + "111 (1).docx")
        builder = aw.DocumentBuilder(doc_main)
        paragraphs = doc_main.get_child_nodes(aw.NodeType.PARAGRAPH, True)
        for para in paragraphs:
            para = para.as_paragraph()
            para_content = para.to_string(aw.SaveFormat.TEXT)
            para_content = para_content.replace("\r", "")
            para_content = para_content.strip()  # 特殊地方，发现目录中有这个符号，暂时不知道符号是干啥的
            if para_content in header_list or para_content.capitalize() in header_list:
                table_header = aw.Paragraph(doc_main)
                table_header.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
                table_header.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
                if para_content in ["APPENDICES", "REFERENCES"]:
                    last_related_paragraph = para
                    while (
                            last_related_paragraph.next_sibling
                            and last_related_paragraph.next_sibling.as_paragraph().paragraph_format.outline_level
                            > para.paragraph_format.outline_level
                    ):
                        last_related_paragraph = last_related_paragraph.next_sibling

                    _doc = aw.Document(os.path.join(MY_DIR + "t_ae_1(1).docx"))
                    _tables = _doc.get_child_nodes(aw.NodeType.TABLE, True)
                    _tables = [t for t in _tables]
                    for table in _tables:
                        if table.get_ancestor(aw.NodeType.BODY):
                            i = _tables.index(table)
                            temp = _tables[i]
                            _tables[i] = _tables[i - 1]
                            _tables[i - 1] = temp

                    for table in _tables:  # 将table 信息反向的插入到word文件中。TODO 表格美化
                        node = table.clone(True)
                        imported_node = doc_main.import_node(node, True)
                        if imported_node.node_type == aw.NodeType.TABLE and \
                                table.parent_node.node_type is not aw.NodeType.HEADER_FOOTER:
                            imported_table = imported_node.as_table()
                            imported_table.preferred_width = aw.tables.PreferredWidth.from_percent(100)
                            for index, row in enumerate(imported_table.rows):
                                row = row.as_row()
                                # if (
                                #     "Source:" in row.get_text().strip()
                                #     or not row.get_text().strip()
                                # ):
                                #     row.remove()

                                for cell_index, cell in enumerate(row.cells):
                                    cell = cell.as_cell()
                                    cell.cell_format.vertical_alignment = (
                                        aw.tables.CellVerticalAlignment.BOTTOM
                                    )
                                    for paragraph in cell.paragraphs:
                                        paragraph = paragraph.as_paragraph()
                                        # 居中对齐
                                        for run in paragraph.runs:
                                            run = run.as_run()
                                            run.font.name = "Courier New"
                                            run.font.name_far_east = "宋体"
                                            run.font.size = 8

                            # curr_para = para.next_sibling.next_sibling.as_paragraph()


                            last_related_paragraph = last_related_paragraph.parent_node.insert_after(imported_table, last_related_paragraph)
                            # if table.parent_node.node_type == aw.NodeType.HEADER_FOOTER:
                            #     section = para.get_ancestor(aw.NodeType.SECTION).as_section()
                            #     section.headers_footers.clear()
                            #     footer = aw.HeaderFooter(doc_main, aw.HeaderFooterType.FOOTER_PRIMARY)
                            #     section.headers_footers.add(footer)
                            #     footer.append_paragraph(imported_table.get_text())

        doc_main.save(ARTIFACTS_DIR + "Result.docx")

    @staticmethod
    def insert_document(insertion_destination: aw.Node, doc_to_insert: aw.Document):
        """Inserts content of the external document after the specified node.
        Section breaks and section formatting of the inserted document are ignored.
       :param insertion_destination: Node in the destination document after which the content
            Should be inserted. This node should be a block level node (paragraph or table).
       :param doc_to_insert: The document to insert.
        """

        if insertion_destination.node_type not in (aw.NodeType.PARAGRAPH, aw.NodeType.TABLE):
            raise ValueError("The destination node should be either a paragraph or table.")

        destination_parent = insertion_destination.parent_node

        importer = aw.NodeImporter(doc_to_insert, insertion_destination.document,
                                   aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        # Loop through all block-level nodes in the section's body,
        # then clone and insert every node that is not the last empty paragraph of a section.
        for src_section in doc_to_insert.sections:
            for src_node in src_section.as_section().body.get_child_nodes(aw.NodeType.ANY, False):
                if src_node.node_type == aw.NodeType.PARAGRAPH:
                    para = src_node.as_paragraph()
                    if para.is_end_of_section and not para.has_child_nodes:
                        continue

                new_node = importer.import_node(src_node, True)

                destination_parent.insert_after(new_node, insertion_destination)
                insertion_destination = new_node


    # def test_tiaoh(self):
    #     header_list = ["研究事務局"]
    #
    #     doc_main = aw.Document(MY_DIR + "111.docx")
    #     paragraphs = doc_main.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    #     for para in paragraphs:
    #         para = para.as_paragraph()
    #         para_content = para.to_string(aw.SaveFormat.TEXT)
    #         para_content = para_content.replace("\r", "")
    #         para_content = para_content.strip()  # 特殊地方，发现目录中有这个符号，暂时不知道符号是干啥的
    #         if (para_content in header_list or para_content.capitalize() in header_list):  # 如果当前段落中有写作内容，那么找到内容，找到生成的结果
    #             print(para_content)
    #             table_header = aw.Paragraph(doc_main)
    #             table_header.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
    #             table_header.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
    #             try:
    #                 if para_content in ["APPENDICES", "REFERENCES"]:
    #                     para_content = para_content.capitalize()
    #
    #                 idx_num = header_list.index(para_content)
    #                 # 获取header对应的table
    #                 num_tables = len(table_list[idx_num])
    #
    #                 for _index, info in enumerate(table_list[idx_num]):
    #                     table_id = info.get("id", "")
    #                     is_header_and_footer = info.get("is_header_and_footer", "")
    #                     aw.Document(
    #                         os.path.join(settings.UPLOAD_PATH, f"{table_id}.rtf")
    #                     ).save(os.path.join(settings.UPLOAD_PATH, f"{table_id}.docx"))
    #                     # 判断是否需要页眉页脚 False为删除页眉页脚
    #                     if not is_header_and_footer:
    #                         document = aw.Document(
    #                             os.path.join(settings.UPLOAD_PATH, f"{table_id}.docx")
    #                         )
    #                         for section in document.sections:
    #                             section.footer.is_linked_to_previous = True
    #                             section.header.is_linked_to_previous = True
    #                         document.save(
    #                             os.path.join(settings.UPLOAD_PATH, f"{table_id}.docx")
    #                         )
    #                     _doc = aw.Document(
    #                         os.path.join(settings.UPLOAD_PATH, f"{table_id}.docx")
    #                     )
    #                     _tables = _doc.get_child_nodes(aw.NodeType.TABLE, True)
    #                     _tables = [t for t in _tables]
    #                     # 删除表格第一行内容
    #                     for (table) in _tables:  # 将table 信息反向的插入到word文件中。TODO 表格美化
    #                         table_clone = table.as_table().clone(True)
    #                         imported_table = doc_main.import_node(table_clone, True)
    #                         if imported_table.node_type == aw.NodeType.TABLE:
    #                             imported_table = imported_table.as_table()
    #                             imported_table.preferred_width = aw.tables.PreferredWidth.from_percent(100)
    #                             imported_table.first_row.remove()
    #                             for index, row in enumerate(imported_table.rows):
    #                                 row = row.as_row()
    #                                 for cell_index, cell in enumerate(row.cells):
    #                                     cell = cell.as_cell()
    #                                     cell.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.BOTTOM
    #                                     for paragraph in cell.paragraphs:
    #                                         paragraph = paragraph.as_paragraph()
    #                                         # 居中对齐
    #                                         for run in paragraph.runs:
    #                                             run = run.as_run()
    #                                             run.font.name = "Times New Roman"  # 设置西文是新罗马字体
    #                                             run.font.name_far_east = "宋体"
    #                                             run.font.size = 8
    #                             # 在插入段落标题之后插入段落内容
    #                             table_header.parent_node.insert_after(imported_table, table_header)
    #                             if _index < num_tables - 1:
    #                                 table_newline = aw.Paragraph(doc_main)
    #                                 run = aw.Run(doc_main, "")
    #                                 table_newline.append_child(run)
    #                                 imported_table.parent_node.insert_before(
    #                                     table_newline, imported_table
    #                                 )
    #                 else:
    #                     #logger.info(f"当前段落不需要插入表格内容")
    #             except:
    #                 print("没有找到header", traceback.format_exc())
    #     doc_main.save(ARTIFACTS_DIR + "")

    def test_forum_pdf(self):
        hex_color = "#98be25"
        rgb_color = drawing.ColorTranslator.from_html(hex_color)

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_paragraph()
        # Insert another textbox with specific margins.
        text_box_shape = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 120)
        text_box = text_box_shape.text_box
        text_box.internal_margin_top = 15
        text_box.internal_margin_bottom = 15
        text_box.internal_margin_left = 15
        text_box.internal_margin_right = 15

        # Set the outline color of the text box to be transparent (invisible)
        text_box_shape.stroke_color = drawing.Color.transparent

        # Move the builder cursor to the end of the existing paragraph
        builder.move_to(text_box_shape.last_paragraph)

        # Set font name, size, and alignment
        builder.font.name = "Lucida Bright"
        builder.font.size = 9
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        builder.writeln()

        # Write text
        builder.write("Capitolo")

        # Insert a line break
        builder.writeln()

        # Set font name, size, and alignment
        builder.font.name = "Lucida Bright"
        builder.font.size = 43
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        # Write text with different font and size
        builder.write("10")

        # Set the fill color for the textbox
        text_box_shape.fill_color = rgb_color

        builder.move_to(text_box_shape.parent_paragraph)

        builder.writeln()
        builder.writeln()

        builder.paragraph_format.style_name = "Heading 1"
        builder.write("Title of this chapter")

        # Save the document
        doc.save(ARTIFACTS_DIR + "learn1.docx")

    def test_constructor(self):

        # ExStart
        # ExFor:Document.__init__()
        # ExFor:Document.__init__(str,LoadOptions)
        # ExSummary:Shows how to create and load documents.
        # There are two ways of creating a Document object using Aspose.Words.
        # 1 -  Create a blank document:
        doc = aw.Document()

        # New Document objects by default come with the minimal set of nodes
        # required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
        doc.first_section.body.first_paragraph.append_child(aw.Run(doc, "Hello world!"))

        # 2 -  Load a document that exists in the local file system:
        doc = aw.Document(MY_DIR + "Document.docx")

        # Loaded documents will have contents that we can access and edit.
        self.assertEqual("Hello World!", doc.first_section.body.first_paragraph.get_text().strip())

        # Some operations that need to occur during loading, such as using a password to decrypt a document,
        # can be done by passing a LoadOptions object when loading the document.
        doc = aw.Document(MY_DIR + "Encrypted.docx", aw.loading.LoadOptions("docPassword"))

        self.assertEqual("Test encrypted document.", doc.first_section.body.first_paragraph.get_text().strip())
        # ExEnd

    def test_load_from_stream(self):

        # ExStart
        # ExFor:Document.__init__(BytesIO)
        # ExSummary:Shows how to load a document using a stream.
        with open(MY_DIR + "Document.docx", "rb") as stream:
            doc = aw.Document(stream)

            self.assertEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.get_text().strip())

        # ExEnd

    def test_load_from_web(self):

        # ExStart
        # ExFor:Document.__init__(BytesIO)
        # ExSummary:Shows how to load a document from a URL.
        # Create a URL that points to a Microsoft Word document.
        url = "https://filesamples.com/samples/document/docx/sample3.docx"

        # Download the document into a byte array, then load that array into a document using a memory stream.
        request_site = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        data_bytes = urlopen(request_site).read()

        with io.BytesIO(data_bytes) as byte_stream:
            doc = aw.Document(byte_stream)

            # At this stage, we can read and edit the document's contents and then save it to the local file system.
            self.assertEqual(
                "There are eight section headings in this document. At the beginning, \"Sample Document\" is a level 1 heading. " +
                "The main section headings, such as \"Headings\" and \"Lists\" are level 2 headings. " +
                "The Tables section contains two sub-headings, \"Simple Table\" and \"Complex Table,\" which are both level 3 headings.",
                doc.first_section.body.paragraphs[3].get_text().strip())

            doc.save(ARTIFACTS_DIR + "Document.load_from_web.docx")

        # ExEnd

    def test_convert_to_pdf(self):

        # ExStart
        # ExFor:Document.__init__(str)
        # ExFor:Document.save(str)
        # ExSummary:Shows how to open a document and convert it to .PDF.
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.save(ARTIFACTS_DIR + "Document.convert_to_pdf.pdf")
        # ExEnd

    def test_save_to_image_stream(self):

        # ExStart
        # ExFor:Document.save(BytesIO,SaveFormat)
        # ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.name = "Times New Roman"
        builder.font.size = 24
        builder.writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

        builder.insert_image(IMAGE_DIR + "Logo.jpg")

        with io.BytesIO() as stream:
            doc.save(stream, aw.SaveFormat.BMP)

            stream.seek(0, os.SEEK_SET)

            # Read the stream back into an image.
            with drawing.Image.from_stream(stream) as image:
                self.assertEqual(drawing.imaging.ImageFormat.bmp, image.raw_format)
                self.assertEqual(816, image.width)
                self.assertEqual(1056, image.height)

        # ExEnd

    # def test_open_type(self):

    #    #ExStart
    #    #ExFor:LayoutOptions.text_shaper_factory
    #    #ExSummary:Shows how to support OpenType features using the HarfBuzz text shaping engine.
    #    doc = aw.Document(MY_DIR + "OpenType text shaping.docx")

    #    # Aspose.Words can use externally provided text shaper objects,
    #    # which represent fonts and compute shaping information for text.
    #    # A text shaper factory is necessary for documents that use multiple fonts.
    #    # When the text shaper factory set, the layout uses OpenType features.
    #    # An Instance property returns a static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory.
    #    doc.layout_options.text_shaper_factory = aw.shaping.harfbuzz.HarfBuzzTextShaperFactory.instance

    #    # Currently, text shaping is performing when exporting to PDF or XPS formats.
    #    doc.save(ARTIFACTS_DIR + "Document.open_type.pdf")
    #    #ExEnd

    def test_detect_mobi_document_format(self):
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + "Document.mobi")
        self.assertEqual(info.load_format, aw.LoadFormat.MOBI)

    def test_detect_pdf_document_format(self):
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + "Pdf Document.pdf")
        self.assertEqual(info.load_format, aw.LoadFormat.PDF)

    def test_open_pdf_document(self):

        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        self.assertEqual(
            "Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\u000c",
            doc.range.text)

    def test_open_protected_pdf_document(self):

        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        save_options = aw.saving.PdfSaveOptions()
        save_options.encryption_details = aw.saving.PdfEncryptionDetails("Aspose", None)

        doc.save(ARTIFACTS_DIR + "Document.open_protected_pdf_document.pdf", save_options)

        load_options = aw.loading.PdfLoadOptions()
        load_options.password = "Aspose"
        load_options.load_format = aw.LoadFormat.PDF

        doc = aw.Document(ARTIFACTS_DIR + "Document.open_protected_pdf_document.pdf", load_options)

    def test_open_from_stream_with_base_uri(self):

        # ExStart
        # ExFor:Document.__init__(BytesIO,LoadOptions)
        # ExFor:LoadOptions.__init__()
        # ExFor:LoadOptions.base_uri
        # ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
        with open(MY_DIR + "Document.html", "rb") as stream:
            # Pass the URI of the base folder while loading it
            # so that any images with relative URIs in the HTML document can be found.
            load_options = aw.loading.LoadOptions()
            load_options.base_uri = IMAGE_DIR

            doc = aw.Document(stream, load_options)

            # Verify that the first shape of the document contains a valid image.
            shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

            self.assertTrue(shape.is_image)
            self.assertIsNotNone(shape.image_data.image_bytes)
            self.assertAlmostEqual(32.0, aw.ConvertUtil.point_to_pixel(shape.width), delta=0.01)
            self.assertAlmostEqual(32.0, aw.ConvertUtil.point_to_pixel(shape.height), delta=0.01)

        # ExEnd

    def test_insert_html_from_web_page(self):

        # ExStart
        # ExFor:Document.__init__(BytesIO,LoadOptions)
        # ExFor:LoadOptions.__init__(LoadFormat,str,str)
        # ExFor:LoadFormat
        # ExSummary:Shows how save a web page as a .docx file.
        url = "https://www.aspose.com/"

        with io.BytesIO(urlopen(url).read()) as stream:
            # The URL is used again as a "base_uri" to ensure that any relative image paths are retrieved correctly.
            options = aw.loading.LoadOptions(aw.LoadFormat.HTML, "", url)

            # Load the HTML document from stream and pass the LoadOptions object.
            doc = aw.Document(stream, options)

            # At this stage, we can read and edit the document's contents and then save it to the local file system.
            self.assertTrue(doc.get_text().find(
                "HYPERLINK \"https://products.aspose.com/words/family/\" \\o \"Aspose.Words\"") > 0)  # ExSkip

            doc.save(ARTIFACTS_DIR + "Document.insert_html_from_web_page.docx")

        # ExEnd

        self.verify_web_response_status_code(200, url)

    def test_load_encrypted(self):

        # ExStart
        # ExFor:Document.__init__(BytesIO,LoadOptions)
        # ExFor:Document.__init__(str,LoadOptions)
        # ExFor:LoadOptions
        # ExFor:LoadOptions.__init__(str)
        # ExSummary:Shows how to load an encrypted Microsoft Word document.

        # Aspose.Words throw an exception if we try to open an encrypted document without its password.
        with self.assertRaises(Exception):
            doc = aw.Document(MY_DIR + "Encrypted.docx")

        # When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
        options = aw.loading.LoadOptions("docPassword")

        # There are two ways of loading an encrypted document with a LoadOptions object.
        # 1 -  Load the document from the local file system by filename:
        doc = aw.Document(MY_DIR + "Encrypted.docx", options)
        self.assertEqual("Test encrypted document.", doc.get_text().strip())  # ExSkip

        # 2 -  Load the document from a stream:
        with open(MY_DIR + "Encrypted.docx", "rb") as stream:
            doc = aw.Document(stream, options)
            self.assertEqual("Test encrypted document.", doc.get_text().strip())  # ExSkip

        # ExEnd

    def test_temp_folder(self):

        # ExStart
        # ExFor:LoadOptions.temp_folder
        # ExSummary:Shows how to load a document using temporary files.
        # Note that such an approach can reduce memory usage but degrades speed
        load_options = aw.loading.LoadOptions()
        load_options.temp_folder = "C:\\TempFolder\\"

        # Ensure that the directory exists and load
        os.makedirs(load_options.temp_folder, exist_ok=True)

        doc = aw.Document(MY_DIR + "Document.docx", load_options)
        # ExEnd

    def test_convert_to_html(self):

        # ExStart
        # ExFor:Document.save(str,SaveFormat)
        # ExFor:SaveFormat
        # ExSummary:Shows how to convert from DOCX to HTML format.
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.save(ARTIFACTS_DIR + "Document.convert_to_html.html", aw.SaveFormat.HTML)
        # ExEnd

    def test_convert_to_mhtml(self):

        doc = aw.Document(MY_DIR + "Document.docx")
        doc.save(ARTIFACTS_DIR + "Document.convert_to_mhtml.mht")

    def test_convert_to_txt(self):

        doc = aw.Document(MY_DIR + "Document.docx")
        doc.save(ARTIFACTS_DIR + "Document.convert_to_txt.txt")

    def test_convert_to_epub(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")
        doc.save(ARTIFACTS_DIR + "Document.convert_to_epub.epub")

    def test_save_to_stream(self):

        # ExStart
        # ExFor:Document.save(BytesIO,SaveFormat)
        # ExSummary:Shows how to save a document to a stream.
        doc = aw.Document(MY_DIR + "Document.docx")

        with io.BytesIO() as dst_stream:
            doc.save(dst_stream, aw.SaveFormat.DOCX)

            # Verify that the stream contains the document.
            self.assertEqual("Hello World!\r\rHello Word!\r\r\rHello World!",
                             aw.Document(dst_stream).get_text().strip())

        # ExEnd

    ##ExStart
    ##ExFor:INodeChangingCallback
    ##ExFor:INodeChangingCallback.node_inserting
    ##ExFor:INodeChangingCallback.node_inserted
    ##ExFor:INodeChangingCallback.node_removing
    ##ExFor:INodeChangingCallback.node_removed
    ##ExFor:NodeChangingArgs
    ##ExFor:NodeChangingArgs.node
    ##ExFor:DocumentBase.node_changing_callback
    ##ExSummary:Shows how customize node changing with a callback.
    # def test_font_change_via_callback(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Set the node changing callback to custom implementation,
    #    # then add/remove nodes to get it to generate a log.
    #    callback = ExDocument.HandleNodeChangingFontChanger()
    #    doc.node_changing_callback = callback

    #    builder.writeln("Hello world!")
    #    builder.writeln("Hello again!")
    #    builder.insert_field(" HYPERLINK \"https://www.google.com/\" ")
    #    builder.insert_shape(aw.drawing.ShapeType.RECTANGLE, 300, 300)

    #    doc.range.fields[0].remove()

    #    print(callback.get_log())
    #    self._test_font_change_via_callback(callback.get_log()) #ExSkip

    # class HandleNodeChangingFontChanger(aw.INodeChangingCallback):
    #    """Logs the date and time of each node insertion and removal.
    #    Sets a custom font name/size for the text contents of Run nodes."""

    #    def __init__(self):

    #        self.log = io.StringIO()

    #    def node_inserted(self, args: aw.NodeChangingArgs):

    #        self.log.write(f"\tType:\t{args.node.node_type}\n")
    #        self.log.write(f"\tHash:\t{args.node.get_hash_code()}\n")

    #        if args.node.node_type == aw.NodeType.RUN:
    #            font = args.node.as_run().font
    #            self.log.write(f"\tFont:\tChanged from \"{font.Name}\" {font.Size}pt")

    #            font.size = 24
    #            font.name = "Arial"

    #            self.log.write(f" to \"{font.Name}\" {font.Size}pt\n")
    #            self.log.write(f"\tContents:\n\t\t\"{args.node.get_text()}\"\n")

    #    def node_inserting(self, args: aw.NodeChangingArgs):

    #        self.log.write(f"\n{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\tNode insertion:\n")

    #    def node_removed(self, args: aw.NodeChangingArgs):

    #        self.log.write(f"\tType:\t{args.node.node_type}\n")
    #        self.log.write(f"\tHash code:\t{hash(args.node)}\n")

    #    def node_removing(self, args: aw.NodeChangingArgs):

    #        self.log.write(f"\n{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\tNode removal:\n")

    #    def get_log(self) -> str:

    #        return self.log.getvalue()

    ##ExEnd

    # def _test_font_change_via_callback(self, log: str):

    #    self.assertEqual(10, log.count("insertion"))
    #    self.assertEqual(5, log.count("removal"))

    def test_append_document(self):

        # ExStart
        # ExFor:Document.append_document(Document,ImportFormatMode)
        # ExSummary:Shows how to append a document to the end of another document.
        src_doc = aw.Document()
        src_doc.first_section.body.append_paragraph("Source document text. ")

        dst_doc = aw.Document()
        dst_doc.first_section.body.append_paragraph("Destination document text. ")

        # Append the source document to the destination document while preserving its formatting,
        # then save the source document to the local file system.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        self.assertEqual(2, dst_doc.sections.count)  # ExSkip

        dst_doc.save(ARTIFACTS_DIR + "Document.append_document.docx")
        # ExEnd

        out_doc_text = aw.Document(ARTIFACTS_DIR + "Document.append_document.docx").get_text()

        self.assertTrue(out_doc_text.startswith(dst_doc.get_text()))
        self.assertTrue(out_doc_text.endswith(src_doc.get_text()))

    # The file path used below does not point to an existing file.
    def test_append_document_from_automation(self):

        doc = aw.Document()

        # We should call this method to clear this document of any existing content.
        doc.remove_all_children()

        record_count = 5
        for i in range(1, record_count + 1):
            src_doc = aw.Document()

            with self.assertRaises(Exception):
                src_doc = aw.Document("C:\\DetailsList.doc")

            # Append the source document at the end of the destination document.
            doc.append_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES)

            # Automation required you to insert a new section break at this point, however, in Aspose.Words we
            # do not need to do anything here as the appended document is imported as separate sections already

            # Unlink all headers/footers in this section from the previous section headers/footers
            # if this is the second document or above being appended.
            if i > 1:
                with self.assertRaises(Exception):
                    doc.sections[i].headers_footers.link_to_previous(False)

    def test_import_list(self):

        for is_keep_source_numbering in (True, False):
            with self.subTest(is_keep_source_numbering=is_keep_source_numbering):
                # ExStart
                # ExFor:ImportFormatOptions.keep_source_numbering
                # ExSummary:Shows how to import a document with numbered lists.
                src_doc = aw.Document(MY_DIR + "List source.docx")
                dst_doc = aw.Document(MY_DIR + "List destination.docx")

                self.assertEqual(4, dst_doc.lists.count)

                options = aw.ImportFormatOptions()

                # If there is a clash of list styles, apply the list format of the source document.
                # Set the "keep_source_numbering" property to "False" to not import any list numbers into the destination document.
                # Set the "keep_source_numbering" property to "True" import all clashing
                # list style numbering with the same appearance that it had in the source document.
                options.keep_source_numbering = is_keep_source_numbering

                dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, options)
                dst_doc.update_list_labels()

                if is_keep_source_numbering:
                    self.assertEqual(5, dst_doc.lists.count)
                else:
                    self.assertEqual(4, dst_doc.lists.count)
                # ExEnd

    def test_keep_source_numbering_same_list_ids(self):

        # ExStart
        # ExFor:ImportFormatOptions.keep_source_numbering
        # ExFor:NodeImporter.__init__(DocumentBase,DocumentBase,ImportFormatMode,ImportFormatOptions)
        # ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
        src_doc = aw.Document(MY_DIR + "List with the same definition identifier - source.docx")
        dst_doc = aw.Document(MY_DIR + "List with the same definition identifier - destination.docx")

        # Set the "keep_source_numbering" property to "True" to apply a different list definition ID
        # to identical styles as Aspose.Words imports them into destination documents.
        import_format_options = aw.ImportFormatOptions()
        import_format_options.keep_source_numbering = True

        dst_doc.append_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES, import_format_options)
        dst_doc.update_list_labels()
        # ExEnd

        para_text = dst_doc.sections[1].body.last_paragraph.get_text()

        self.assertTrue(para_text.startswith("13->13"))
        self.assertEqual("1.", dst_doc.sections[1].body.last_paragraph.list_label.label_string)

    def test_merge_pasted_lists(self):

        # ExStart
        # ExFor:ImportFormatOptions.merge_pasted_lists
        # ExSummary:Shows how to merge lists from a documents.
        src_doc = aw.Document(MY_DIR + "List item.docx")
        dst_doc = aw.Document(MY_DIR + "List destination.docx")

        options = aw.ImportFormatOptions()
        options.merge_pasted_lists = True

        # Set the "merge_pasted_lists" property to "True" pasted lists will be merged with surrounding lists.
        dst_doc.append_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES, options)

        dst_doc.save(ARTIFACTS_DIR + "Document.merge_pasted_lists.docx")
        # ExEnd

    def test_force_copy_styles(self):

        # ExStart
        # ExFor:ImportFormatOptions.force_copy_styles
        # ExSummary:Shows how to copy source styles with unique names forcibly.

        # Both documents contain MyStyle1 and MyStyle2, MyStyle3 exists only in a source document.
        src_doc = aw.Document(MY_DIR + "Styles source.docx");
        dst_doc = aw.Document(MY_DIR + "Styles destination.docx");

        options = aw.ImportFormatOptions()
        options.force_copy_styles = True
        dst_doc.append_document(src_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING, options)

        paras = dst_doc.sections[1].body.paragraphs

        self.assertEqual(paras[0].paragraph_format.style.name, "MyStyle1_0")
        self.assertEqual(paras[1].paragraph_format.style.name, "MyStyle2_0")
        self.assertEqual(paras[2].paragraph_format.style.name, "MyStyle3")
        # ExEnd

    def test_validate_individual_document_signatures(self):

        # ExStart
        # ExFor:CertificateHolder.certificate
        # ExFor:Document.digital_signatures
        # ExFor:DigitalSignature
        # ExFor:DigitalSignatureCollection
        # ExFor:DigitalSignature.is_valid
        # ExFor:DigitalSignature.comments
        # ExFor:DigitalSignature.sign_time
        # ExFor:DigitalSignature.signature_type
        # ExSummary:Shows how to validate and display information about each signature in a document.
        doc = aw.Document(MY_DIR + "Digitally signed.docx")

        for signature in doc.digital_signatures:
            print(f"\n{'Valid' if signature.is_valid else 'Invalid'} signature: ")
            print(f"\tReason:\t{signature.comments}")
            print(f"\tType:\t{signature.signature_type}")
            print(f"\tSign time:\t{signature.sign_time}")
            # System.Security.Cryptography.X509Certificates.X509Certificate2 is not supported. That is why the following information is not accesible.
            # print(f"\tSubject name:\t{signature.certificate_holder.certificate.subject_name}")
            # print(f"\tIssuer name:\t{signature.certificate_holder.certificate.issuer_name.name}")
            print()

        # ExEnd

        self.assertEqual(1, doc.digital_signatures.count)

        digital_sig = doc.digital_signatures[0]

        self.assertTrue(digital_sig.is_valid)
        self.assertEqual("Test Sign", digital_sig.comments)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG, digital_sig.signature_type)
        # System.Security.Cryptography.X509Certificates.X509Certificate2 is not supported. That is why the following information is not accesible.
        # self.assertTrue(digital_sig.certificate_holder.certificate.subject.contains("Aspose Pty Ltd"))
        # self.assertIsNotNone(digital_sig.certificate_holder.certificate.issuer_name.name is not None)
        # self.assertIn("VeriSign", digital_sig.certificate_holder.certificate.issuer_name.name)

    @unittest.skip("DigitalSignatureUtil.sing method is not supported")
    def test_digital_signature(self):

        # ExStart
        # ExFor:DigitalSignature.certificate_holder
        # ExFor:DigitalSignature.issuer_name
        # ExFor:DigitalSignature.subject_name
        # ExFor:DigitalSignatureCollection
        # ExFor:DigitalSignatureCollection.is_valid
        # ExFor:DigitalSignatureCollection.count
        # ExFor:DigitalSignatureCollection.__getitem__(int)
        # ExFor:DigitalSignatureUtil.sign(BytesIO,BytesIO,CertificateHolder)
        # ExFor:DigitalSignatureUtil.sign(str,str,CertificateHolder)
        # ExFor:DigitalSignatureType
        # ExFor:Document.digital_signatures
        # ExSummary:Shows how to sign documents with X.509 certificates.
        # Verify that a document is not signed.
        self.assertFalse(aw.FileFormatUtil.detect_file_format(MY_DIR + "Document.docx").has_digital_signature)

        # Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
        certificate_holder = aw.digitalsignatures.CertificateHolder.create(MY_DIR + "morzal.pfx", "aw", None)

        # There are two ways of saving a signed copy of a document to the local file system:
        # 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
        sign_options = aw.digitalsignatures.SignOptions()
        sign_options.sign_time = datetime.utcnow()
        aw.digitalsignatures.DigitalSignatureUtil.sign(
            MY_DIR + "Document.docx", ARTIFACTS_DIR + "Document.digital_signature.docx",
            certificate_holder, sign_options)

        self.assertTrue(aw.FileFormatUtil.detect_file_format(
            ARTIFACTS_DIR + "Document.digital_signature.docx").has_digital_signature)

        # 2 - Take a document from a stream and save a signed copy to another stream.
        with open(MY_DIR + "Document.docx", "rb") as in_doc:
            with open(ARTIFACTS_DIR + "Document.digital_signature.docx", "wb") as out_doc:
                aw.digitalsignatures.DigitalSignatureUtil.sign(in_doc, out_doc, certificate_holder)

        self.assertTrue(aw.FileFormatUtil.detect_file_format(
            ARTIFACTS_DIR + "Document.digital_signature.docx").has_digital_signature)

        # Please verify that all of the document's digital signatures are valid and check their details.
        signed_doc = aw.Document(ARTIFACTS_DIR + "Document.digital_signature.docx")
        digital_signature_collection = signed_doc.digital_signatures

        self.assertTrue(digital_signature_collection.is_valid)
        self.assertEqual(1, digital_signature_collection.count)
        self.assertEqual(aw.digitalsignatures.DigitalSignatureType.XML_DSIG,
                         digital_signature_collection[0].signature_type)
        self.assertEqual("CN=Morzal.Me", signed_doc.digital_signatures[0].issuer_name)
        self.assertEqual("CN=Morzal.Me", signed_doc.digital_signatures[0].subject_name)
        # ExEnd

    def test_append_all_documents_in_folder(self):

        # ExStart
        # ExFor:Document.append_document(Document,ImportFormatMode)
        # ExSummary:Shows how to append all the documents in a folder to the end of a template document.
        dst_doc = aw.Document()

        builder = aw.DocumentBuilder(dst_doc)
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.writeln("Template Document")
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
        builder.writeln("Some content here")
        self.assertEqual(5, dst_doc.styles.count)  # ExSkip
        self.assertEqual(1, dst_doc.sections.count)  # ExSkip

        # Append all unencrypted documents with the .doc extension
        # from our local file system directory to the base document.
        doc_files = glob.glob(MY_DIR + "*.doc")
        for file_name in doc_files:
            info = aw.FileFormatUtil.detect_file_format(file_name)
            if info.is_encrypted:
                continue

            src_doc = aw.Document(file_name)
            dst_doc.append_document(src_doc, aw.ImportFormatMode.USE_DESTINATION_STYLES)

        dst_doc.save(ARTIFACTS_DIR + "Document.append_all_documents_in_folder.doc")
        # ExEnd

        self.assertEqual(7, dst_doc.styles.count)
        self.assertEqual(9, dst_doc.sections.count)

    def test_join_runs_with_same_formatting(self):

        # ExStart
        # ExFor:Document.join_runs_with_same_formatting
        # ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        # Open a document that contains adjacent runs of text with identical formatting,
        # which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # If any number of these runs are adjacent with identical formatting,
        # then the document may be simplified.
        self.assertEqual(317, doc.get_child_nodes(aw.NodeType.RUN, True).count)

        # Combine such runs with this method and verify the number of run joins that will take place.
        self.assertEqual(121, doc.join_runs_with_same_formatting())

        # The number of joins and the number of runs we have after the join
        # should add up the number of runs we had initially.
        self.assertEqual(196, doc.get_child_nodes(aw.NodeType.RUN, True).count)
        # ExEnd

    def test_default_tab_stop(self):

        # ExStart
        # ExFor:Document.default_tab_stop
        # ExFor:ControlChar.TAB
        # ExFor:ControlChar.TAB_CHAR
        # ExSummary:Shows how to set a custom interval for tab stop positions.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Set tab stops to appear every 72 points (1 inch).
        builder.document.default_tab_stop = 72

        # Each tab character snaps the text after it to the next closest tab stop position.
        builder.writeln("Hello" + aw.ControlChar.TAB + "World!")
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        self.assertEqual(72, doc.default_tab_stop)

    def test_clone_document(self):

        # ExStart
        # ExFor:Document.clone()
        # ExSummary:Shows how to deep clone a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        # Cloning will produce a new document with the same contents as the original,
        # but with a unique copy of each of the original document's nodes.
        clone = doc.clone()

        self.assertEqual(doc.first_section.body.first_paragraph.runs[0].get_text(),
                         clone.first_section.body.first_paragraph.runs[0].text)
        self.assertIsNot(doc.first_section.body.first_paragraph.runs[0],
                         clone.first_section.body.first_paragraph.runs[0])
        # ExEnd

    def test_document_get_text_to_string(self):

        # ExStart
        # ExFor:CompositeNode.get_text
        # ExFor:Node.to_string(SaveFormat)
        # ExSummary:Shows the difference between calling the get_text and to_string methods on a node.
        doc = aw.Document()

        builder = aw.DocumentBuilder(doc)
        builder.insert_field("MERGEFIELD Field")

        # get_text will retrieve the visible text as well as field codes and special characters.
        self.assertEqual("\u0013MERGEFIELD Field\u0014«Field»\u0015\u000c", doc.get_text())

        # to_string will give us the document's appearance if saved to a passed save format.
        self.assertEqual("«Field»\r\n", doc.to_string(aw.SaveFormat.TEXT))
        # ExEnd

    def test_document_byte_array(self):

        doc = aw.Document(MY_DIR + "Document.docx")

        stream_out = io.BytesIO()
        doc.save(stream_out, aw.SaveFormat.DOCX)

        doc_bytes = stream_out.getvalue()

        stream_in = io.BytesIO(doc_bytes)

        load_doc = aw.Document(stream_in)
        self.assertEqual(doc.get_text(), load_doc.get_text())

    def test_protect_unprotect(self):

        # ExStart
        # ExFor:Document.protect(ProtectionType,str)
        # ExFor:Document.protection_type
        # ExFor:Document.unprotect()
        # ExFor:Document.unprotect(str)
        # ExSummary:Shows how to protect and unprotect a document.
        doc = aw.Document()
        doc.protect(aw.ProtectionType.READ_ONLY, "password")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)

        # If we open this document with Microsoft Word intending to edit it,
        # we will need to apply the password to get through the protection.
        doc.save(ARTIFACTS_DIR + "Document.protect_unprotect.docx")

        # Note that the protection only applies to Microsoft Word users opening our document.
        # We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
        protected_doc = aw.Document(ARTIFACTS_DIR + "Document.protect_unprotect.docx")

        self.assertEqual(aw.ProtectionType.READ_ONLY, protected_doc.protection_type)

        builder = aw.DocumentBuilder(protected_doc)
        builder.writeln("Text added to a protected document.")
        self.assertEqual("Text added to a protected document.", protected_doc.range.text.strip())  # ExSkip

        # There are two ways of removing protection from a document.
        # 1 - With no password:
        doc.unprotect()

        self.assertEqual(aw.ProtectionType.NO_PROTECTION, doc.protection_type)

        doc.protect(aw.ProtectionType.READ_ONLY, "NewPassword")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)

        doc.unprotect("WrongPassword")

        self.assertEqual(aw.ProtectionType.READ_ONLY, doc.protection_type)

        # 2 - With the correct password:
        doc.unprotect("NewPassword")

        self.assertEqual(aw.ProtectionType.NO_PROTECTION, doc.protection_type)
        # ExEnd

    def test_document_ensure_minimum(self):

        # ExStart
        # ExFor:Document.ensure_minimum
        # ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
        # A newly created document contains one child Section, which includes one child Body and one child Paragraph.
        # We can edit the document body's contents by adding nodes such as Runs or inline Shapes to that paragraph.
        doc = aw.Document()
        nodes = doc.get_child_nodes(aw.NodeType.ANY, True)

        self.assertEqual(aw.NodeType.SECTION, nodes[0].node_type)
        self.assertEqual(doc, nodes[0].parent_node)

        self.assertEqual(aw.NodeType.BODY, nodes[1].node_type)
        self.assertEqual(nodes[0], nodes[1].parent_node)

        self.assertEqual(aw.NodeType.PARAGRAPH, nodes[2].node_type)
        self.assertEqual(nodes[1], nodes[2].parent_node)

        # This is the minimal set of nodes that we need to be able to edit the document.
        # We will no longer be able to edit the document if we remove any of them.
        doc.remove_all_children()

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.ANY, True).count)

        # Call this method to make sure that the document has at least those three nodes so we can edit it again.
        doc.ensure_minimum()

        self.assertEqual(aw.NodeType.SECTION, nodes[0].node_type)
        self.assertEqual(aw.NodeType.BODY, nodes[1].node_type)
        self.assertEqual(aw.NodeType.PARAGRAPH, nodes[2].node_type)

        nodes[2].as_paragraph().runs.add(aw.Run(doc, "Hello world!"))
        # ExEnd

        self.assertEqual("Hello world!", doc.get_text().strip())

    def test_remove_macros_from_document(self):

        # ExStart
        # ExFor:Document.remove_macros
        # ExSummary:Shows how to remove all macros from a document.
        doc = aw.Document(MY_DIR + "Macro.docm")

        self.assertTrue(doc.has_macros)
        self.assertEqual("Project", doc.vba_project.name)

        # Remove the document's VBA project, along with all its macros.
        doc.remove_macros()

        self.assertFalse(doc.has_macros)
        self.assertIsNone(doc.vba_project)
        # ExEnd

    def test_get_page_count(self):

        # ExStart
        # ExFor:Document.page_count
        # ExSummary:Shows how to count the number of pages in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Page 1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write("Page 2")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write("Page 3")

        # Verify the expected page count of the document.
        self.assertEqual(3, doc.page_count)

        # Getting the page_count property invoked the document's page layout to calculate the value.
        # This operation will not need to be re-done when rendering the document to a fixed page save format,
        # such as .pdf. So you can save some time, especially with more complex documents.
        doc.save(ARTIFACTS_DIR + "Document.get_page_count.pdf")
        # ExEnd

    def test_get_updated_page_properties(self):

        # ExStart
        # ExFor:Document.update_word_count()
        # ExFor:Document.update_word_count(bool)
        # ExFor:BuiltInDocumentProperties.characters
        # ExFor:BuiltInDocumentProperties.words
        # ExFor:BuiltInDocumentProperties.paragraphs
        # ExFor:BuiltInDocumentProperties.lines
        # ExSummary:Shows how to update all list labels in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
        builder.write("Ut enim ad minim veniam, " +
                      "quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.")

        # Aspose.Words does not track document metrics like these in real time.
        self.assertEqual(0, doc.built_in_document_properties.characters)
        self.assertEqual(0, doc.built_in_document_properties.words)
        self.assertEqual(1, doc.built_in_document_properties.paragraphs)
        self.assertEqual(1, doc.built_in_document_properties.lines)

        # To get accurate values for three of these properties, we will need to update them manually.
        doc.update_word_count()

        self.assertEqual(196, doc.built_in_document_properties.characters)
        self.assertEqual(36, doc.built_in_document_properties.words)
        self.assertEqual(2, doc.built_in_document_properties.paragraphs)

        # For the line count, we will need to call a specific overload of the updating method.
        self.assertEqual(1, doc.built_in_document_properties.lines)

        doc.update_word_count(True)

        self.assertEqual(4, doc.built_in_document_properties.lines)
        # ExEnd

    def test_table_style_to_direct_formatting(self):

        # ExStart
        # ExFor:CompositeNode.get_child
        # ExFor:Document.expand_table_styles_to_direct_formatting
        # ExSummary:Shows how to apply the properties of a table's style directly to the table's elements.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        table = builder.start_table()
        builder.insert_cell()
        builder.write("Hello world!")
        builder.end_table()

        table_style = doc.styles.add(aw.StyleType.TABLE, "MyTableStyle1").as_table_style()
        table_style.row_stripe = 3
        table_style.cell_spacing = 5
        table_style.shading.background_pattern_color = drawing.Color.antique_white
        table_style.borders.color = drawing.Color.blue
        table_style.borders.line_style = aw.LineStyle.DOT_DASH

        table.style = table_style

        # This method concerns table style properties such as the ones we set above.
        doc.expand_table_styles_to_direct_formatting()

        doc.save(ARTIFACTS_DIR + "Document.table_style_to_direct_formatting.docx")
        # ExEnd

        self.verify_doc_package_file_contains_string("<w:tblStyleRowBandSize w:val=\"3\" />",
                                                     ARTIFACTS_DIR + "Document.table_style_to_direct_formatting.docx",
                                                     "word/document.xml")
        self.verify_doc_package_file_contains_string("<w:tblCellSpacing w:w=\"100\" w:type=\"dxa\" />",
                                                     ARTIFACTS_DIR + "Document.table_style_to_direct_formatting.docx",
                                                     "word/document.xml")
        self.verify_doc_package_file_contains_string(
            "<w:tblBorders><w:top w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:left w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:bottom w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:right w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideH w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideV w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /></w:tblBorders>",
            ARTIFACTS_DIR + "Document.table_style_to_direct_formatting.docx", "word/document.xml")

    def test_get_original_file_info(self):

        # ExStart
        # ExFor:Document.original_file_name
        # ExFor:Document.original_load_format
        # ExSummary:Shows how to retrieve details of a document's load operation.
        doc = aw.Document(MY_DIR + "Document.docx")

        self.assertEqual(MY_DIR + "Document.docx", doc.original_file_name)
        self.assertEqual(aw.LoadFormat.DOCX, doc.original_load_format)
        # ExEnd

    # WORDSNET-16099
    def test_footnote_columns(self):

        # ExStart
        # ExFor:FootnoteOptions
        # ExFor:FootnoteOptions.columns
        # ExSummary:Shows how to split the footnote section into a given number of columns.
        doc = aw.Document(MY_DIR + "Footnotes and endnotes.docx")
        self.assertEqual(0, doc.footnote_options.columns)  # ExSkip

        doc.footnote_options.columns = 2
        doc.save(ARTIFACTS_DIR + "Document.footnote_columns.docx")
        # ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Document.footnote_columns.docx")

        self.assertEqual(2, doc.first_section.page_setup.footnote_options.columns)

    def test_compare(self):

        # ExStart
        # ExFor:Document.compare(Document,str,datetime)
        # ExFor:RevisionCollection.accept_all
        # ExSummary:Shows how to compare documents.
        doc_original = aw.Document()
        builder = aw.DocumentBuilder(doc_original)
        builder.writeln("This is the original document.")

        doc_edited = aw.Document()
        builder = aw.DocumentBuilder(doc_edited)
        builder.writeln("This is the edited document.")

        # Comparing documents with revisions will throw an exception.
        if doc_original.revisions.count == 0 and doc_edited.revisions.count == 0:
            doc_original.compare(doc_edited, "authorName", datetime.now())

        # After the comparison, the original document will gain a new revision
        # for every element that is different in the edited document.
        self.assertEqual(2, doc_original.revisions.count)  # ExSkip
        for revision in doc_original.revisions:
            print(f"Revision type: {revision.revision_type}, on a node of type \"{revision.parent_node.node_type}\"")
            print(f"\tChanged text: \"{revision.parent_node.get_text()}\"")

        # Accepting these revisions will transform the original document into the edited document.
        doc_original.revisions.accept_all()

        self.assertEqual(doc_original.get_text(), doc_edited.get_text())
        # ExEnd

        doc_original = DocumentHelper.save_open(doc_original)
        self.assertEqual(0, doc_original.revisions.count)

    def test_compare_document_with_revisions(self):

        doc1 = aw.Document()
        builder = aw.DocumentBuilder(doc1)
        builder.writeln("Hello world! This text is not a revision.")

        doc_with_revision = aw.Document()
        builder = aw.DocumentBuilder(doc_with_revision)

        doc_with_revision.start_track_revisions("John Doe")
        builder.writeln("This is a revision.")

        with self.assertRaises(Exception):
            doc_with_revision.compare(doc1, "John Doe", datetime.now())

    def test_compare_options(self):

        # ExStart
        # ExFor:CompareOptions
        # ExFor:CompareOptions.ignore_formatting
        # ExFor:CompareOptions.ignore_case_changes
        # ExFor:CompareOptions.ignore_comments
        # ExFor:CompareOptions.ignore_tables
        # ExFor:CompareOptions.ignore_fields
        # ExFor:CompareOptions.ignore_footnotes
        # ExFor:CompareOptions.ignore_textboxes
        # ExFor:CompareOptions.ignore_headers_and_footers
        # ExFor:CompareOptions.target
        # ExFor:ComparisonTargetType
        # ExFor:Document.compare(Document,str,datetime,CompareOptions)
        # ExSummary:Shows how to filter specific types of document elements when making a comparison.
        # Create the original document and populate it with various kinds of elements.
        doc_original = aw.Document()
        builder = aw.DocumentBuilder(doc_original)

        # Paragraph text referenced with an endnote:
        builder.writeln("Hello world! This is the first paragraph.")
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Original endnote text.")

        # Table:
        builder.start_table()
        builder.insert_cell()
        builder.write("Original cell 1 text")
        builder.insert_cell()
        builder.write("Original cell 2 text")
        builder.end_table()

        # Textbox:
        text_box = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 150, 20)
        builder.move_to(text_box.first_paragraph)
        builder.write("Original textbox contents")

        # DATE field:
        builder.move_to(doc_original.first_section.body.append_paragraph(""))
        builder.insert_field(" DATE ")

        # Comment:
        new_comment = aw.Comment(doc_original, "John Doe", "J.D.", datetime.now())
        new_comment.set_text("Original comment.")
        builder.current_paragraph.append_child(new_comment)

        # Header:
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.writeln("Original header contents.")

        # Create a clone of our document and perform a quick edit on each of the cloned document's elements.
        doc_edited = doc_original.clone(True).as_document()
        first_paragraph = doc_edited.first_section.body.first_paragraph

        first_paragraph.runs[0].text = "hello world! this is the first paragraph, after editing."
        first_paragraph.paragraph_format.style = doc_edited.styles.get_by_style_identifier(aw.StyleIdentifier.HEADING1)
        doc_edited.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote().first_paragraph.runs[
            1].text = "Edited endnote text."
        doc_edited.get_child(aw.NodeType.TABLE, 0, True).as_table().first_row.cells[1].first_paragraph.runs[
            0].text = "Edited Cell 2 contents"
        doc_edited.get_child(aw.NodeType.SHAPE, 0, True).as_shape().first_paragraph.runs[
            0].text = "Edited textbox contents"
        doc_edited.range.fields[0].as_field_date().use_lunar_calendar = True
        doc_edited.get_child(aw.NodeType.COMMENT, 0, True).as_comment().first_paragraph.runs[0].text = "Edited comment."
        doc_edited.first_section.headers_footers.header_primary.first_paragraph.runs[0].text = "Edited header contents."

        # Comparing documents creates a revision for every edit in the edited document.
        # A CompareOptions object has a series of flags that can suppress revisions
        # on each respective type of element, effectively ignoring their change.
        compare_options = aw.comparing.CompareOptions()
        compare_options.ignore_formatting = False
        compare_options.ignore_case_changes = False
        compare_options.ignore_comments = False
        compare_options.ignore_tables = False
        compare_options.ignore_fields = False
        compare_options.ignore_footnotes = False
        compare_options.ignore_textboxes = False
        compare_options.ignore_headers_and_footers = False
        compare_options.target = aw.comparing.ComparisonTargetType.NEW

        doc_original.compare(doc_edited, "John Doe", datetime.now(), compare_options)
        doc_original.save(ARTIFACTS_DIR + "Document.compare_options.docx")
        # ExEnd

        doc_original = aw.Document(ARTIFACTS_DIR + "Document.compare_options.docx")

        self.verify_footnote(aw.notes.FootnoteType.ENDNOTE, True, "",
                             "OriginalEdited endnote text.",
                             doc_original.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())

    def test_ignore_dml_unique_id(self):

        for is_ignore_dml_unique_id in (False, True):
            with self.subTest(is_ignore_dml_unique_id=is_ignore_dml_unique_id):
                # ExStart
                # ExFor:CompareOptions.ignore_dml_unique_id
                # ExSummary:Shows how to compare documents ignoring DML unique ID.
                doc_a = aw.Document(MY_DIR + "DML unique ID original.docx")
                doc_b = aw.Document(MY_DIR + "DML unique ID compare.docx")

                # By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
                # If we are ignoring DML's unique ID, and revisions count were 0.
                compare_options = aw.comparing.CompareOptions()
                compare_options.ignore_dml_unique_id = is_ignore_dml_unique_id

                doc_a.compare(doc_b, "Aspose.Words", datetime.now(), compare_options)

                self.assertEqual(0 if is_ignore_dml_unique_id else 2, doc_a.revisions.count)
                # ExEnd

    def test_remove_external_schema_references(self):

        # ExStart
        # ExFor:Document.remove_external_schema_references
        # ExSummary:Shows how to remove all external XML schema references from a document.
        doc = aw.Document(MY_DIR + "External XML schema.docx")

        doc.remove_external_schema_references()
        # ExEnd

    def test_track_revisions(self):

        # ExStart
        # ExFor:Document.start_track_revisions(str)
        # ExFor:Document.start_track_revisions(str,datetime)
        # ExFor:Document.stop_track_revisions
        # ExSummary:Shows how to track revisions while editing a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Editing a document usually does not count as a revision until we begin tracking them.
        builder.write("Hello world! ")

        self.assertEqual(0, doc.revisions.count)
        self.assertFalse(doc.first_section.body.paragraphs[0].runs[0].is_insert_revision)

        doc.start_track_revisions("John Doe")

        builder.write("Hello again! ")

        self.assertEqual(1, doc.revisions.count)
        self.assertTrue(doc.first_section.body.paragraphs[0].runs[1].is_insert_revision)
        self.assertEqual("John Doe", doc.revisions[0].author)
        self.assertAlmostEqual(doc.revisions[0].date_time, datetime.now(tz=timezone.utc), delta=timedelta(seconds=1))

        # Stop tracking revisions to not count any future edits as revisions.
        doc.stop_track_revisions()
        builder.write("Hello again! ")

        self.assertEqual(1, doc.revisions.count)
        self.assertFalse(doc.first_section.body.paragraphs[0].runs[2].is_insert_revision)

        # Creating revisions gives them a date and time of the operation.
        # We can disable this by passing "datetime.min" when we start tracking revisions.
        doc.start_track_revisions("John Doe", datetime.min)
        builder.write("Hello again! ")

        self.assertEqual(2, doc.revisions.count)
        self.assertEqual("John Doe", doc.revisions[1].author)
        self.assertEqual(datetime.min, doc.revisions[1].date_time)

        # We can accept/reject these revisions programmatically
        # by calling methods such as "Document.accept_all_revisions", or each revision's "accept" method.
        # In Microsoft Word, we can process them manually via "Review" -> "Changes".
        doc.save(ARTIFACTS_DIR + "Document.track_revisions.docx")
        # ExEnd

    def test_accept_all_revisions(self):

        # ExStart
        # ExFor:Document.accept_all_revisions
        # ExSummary:Shows how to accept all tracking changes in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Edit the document while tracking changes to create a few revisions.
        doc.start_track_revisions("John Doe")
        builder.write("Hello world! ")
        builder.write("Hello again! ")
        builder.write("This is another revision.")
        doc.stop_track_revisions()

        self.assertEqual(3, doc.revisions.count)

        # We can iterate through every revision and accept/reject it as a part of our document.
        # If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
        doc.accept_all_revisions()

        self.assertEqual(0, doc.revisions.count)
        self.assertEqual("Hello world! Hello again! This is another revision.", doc.get_text().strip())
        # ExEnd

    def test_get_revised_properties_of_list(self):

        # ExStart
        # ExFor:RevisionsView
        # ExFor:Document.revisions_view
        # ExSummary:Shows how to switch between the revised and the original view of a document.
        doc = aw.Document(MY_DIR + "Revisions at list levels.docx")
        doc.update_list_labels()

        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual("1.", paragraphs[0].list_label.label_string)
        self.assertEqual("a.", paragraphs[1].list_label.label_string)
        self.assertEqual("", paragraphs[2].list_label.label_string)

        # View the document object as if all the revisions are accepted. Currently supports list labels.
        doc.revisions_view = aw.RevisionsView.FINAL

        self.assertEqual("", paragraphs[0].list_label.label_string)
        self.assertEqual("1.", paragraphs[1].list_label.label_string)
        self.assertEqual("a.", paragraphs[2].list_label.label_string)
        # ExEnd

        doc.revisions_view = aw.RevisionsView.ORIGINAL
        doc.accept_all_revisions()

        self.assertEqual("a.", paragraphs[0].list_label.label_string)
        self.assertEqual("", paragraphs[1].list_label.label_string)
        self.assertEqual("b.", paragraphs[2].list_label.label_string)

    def test_update_thumbnail(self):

        # ExStart
        # ExFor:Document.update_thumbnail()
        # ExFor:Document.update_thumbnail(ThumbnailGeneratingOptions)
        # ExFor:ThumbnailGeneratingOptions
        # ExFor:ThumbnailGeneratingOptions.generate_from_first_page
        # ExFor:ThumbnailGeneratingOptions.thumbnail_size
        # ExSummary:Shows how to update a document's thumbnail.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        builder.insert_image(IMAGE_DIR + "Logo.jpg")

        # There are two ways of setting a thumbnail image when saving a document to .epub.
        # 1 -  Use the document's first page:
        doc.update_thumbnail()
        doc.save(ARTIFACTS_DIR + "Document.update_thumbnail.first_page.epub")

        # 2 -  Use the first image found in the document:
        options = aw.rendering.ThumbnailGeneratingOptions()
        self.assertEqual(drawing.Size(600, 900), options.thumbnail_size)  # ExSkip
        self.assertTrue(options.generate_from_first_page)  # ExSkip
        options.thumbnail_size = drawing.Size(400, 400)
        options.generate_from_first_page = False

        doc.update_thumbnail(options)
        doc.save(ARTIFACTS_DIR + "Document.update_thumbnail.first_image.epub")
        # ExEnd

    def test_hyphenation_options(self):

        # ExStart
        # ExFor:Document.hyphenation_options
        # ExFor:HyphenationOptions
        # ExFor:HyphenationOptions.auto_hyphenation
        # ExFor:HyphenationOptions.consecutive_hyphen_limit
        # ExFor:HyphenationOptions.hyphenation_zone
        # ExFor:HyphenationOptions.hyphenate_caps
        # ExSummary:Shows how to configure automatic hyphenation.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.size = 24
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

        doc.hyphenation_options.auto_hyphenation = True
        doc.hyphenation_options.consecutive_hyphen_limit = 2
        doc.hyphenation_options.hyphenation_zone = 720
        doc.hyphenation_options.hyphenate_caps = True

        doc.save(ARTIFACTS_DIR + "Document.hyphenation_options.docx")
        # ExEnd

        self.assertTrue(doc.hyphenation_options.auto_hyphenation)
        self.assertEqual(2, doc.hyphenation_options.consecutive_hyphen_limit)
        self.assertEqual(720, doc.hyphenation_options.hyphenation_zone)
        self.assertTrue(doc.hyphenation_options.hyphenate_caps)

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "Document.hyphenation_options.docx",
                                                    GOLDS_DIR + "Document.HyphenationOptions Gold.docx"))

    def test_hyphenation_options_default_values(self):

        doc = aw.Document()
        doc = DocumentHelper.save_open(doc)

        self.assertEqual(False, doc.hyphenation_options.auto_hyphenation)
        self.assertEqual(0, doc.hyphenation_options.consecutive_hyphen_limit)
        self.assertEqual(360, doc.hyphenation_options.hyphenation_zone)  # 0.25 inch
        self.assertTrue(doc.hyphenation_options.hyphenate_caps)

    def test_hyphenation_options_exceptions(self):

        doc = aw.Document()

        doc.hyphenation_options.consecutive_hyphen_limit = 0
        with self.assertRaises(Exception):
            doc.hyphenation_options.hyphenation_zone = 0

        with self.assertRaises(Exception):
            doc.hyphenation_options.consecutive_hyphen_limit = -1

        doc.hyphenation_options.hyphenation_zone = 360

    def test_ooxml_compliance_version(self):

        # ExStart
        # ExFor:Document.compliance
        # ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
        # The compliance version varies between documents created by different versions of Microsoft Word.
        doc = aw.Document(MY_DIR + "Document.doc")

        self.assertEqual(doc.compliance, aw.saving.OoxmlCompliance.ECMA376_2006)

        doc = aw.Document(MY_DIR + "Document.docx")

        self.assertEqual(doc.compliance, aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL)
        # ExEnd

    @unittest.skip("WORDSNET-20342")
    def test_image_save_options(self):

        # ExStart
        # ExFor:Document.save(str,SaveOptions)
        # ExFor:SaveOptions.use_anti_aliasing
        # ExFor:SaveOptions.use_high_quality_rendering
        # ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
        doc = aw.Document(MY_DIR + "Rendering.docx")
        builder = aw.DocumentBuilder(doc)

        builder.font.size = 60
        builder.writeln("Some text.")

        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        self.assertFalse(options.use_anti_aliasing)  # ExSkip
        self.assertFalse(options.use_high_quality_rendering)  # ExSkip

        doc.save(ARTIFACTS_DIR + "Document.image_save_options.default.jpg", options)

        options.use_anti_aliasing = True
        options.use_high_quality_rendering = True

        doc.save(ARTIFACTS_DIR + "Document.image_save_options.high_quality.jpg", options)
        # ExEnd

        self.verify_image(794, 1122, ARTIFACTS_DIR + "Document.image_save_options.default.jpg")
        self.verify_image(794, 1122, ARTIFACTS_DIR + "Document.image_save_options.high_quality.jpg")

    def test_cleanup(self):

        # ExStart
        # ExFor:Document.cleanup()
        # ExSummary:Shows how to remove unused custom styles from a document.
        doc = aw.Document()

        doc.styles.add(aw.StyleType.LIST, "MyListStyle1")
        doc.styles.add(aw.StyleType.LIST, "MyListStyle2")
        doc.styles.add(aw.StyleType.CHARACTER, "MyParagraphStyle1")
        doc.styles.add(aw.StyleType.CHARACTER, "MyParagraphStyle2")

        # Combined with the built-in styles, the document now has eight styles.
        # A custom style counts as "used" while applied to some part of the document,
        # which means that the four styles we added are currently unused.
        self.assertEqual(8, doc.styles.count)

        # Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
        builder = aw.DocumentBuilder(doc)
        builder.font.style = doc.styles.get_by_name("MyParagraphStyle1")
        builder.writeln("Hello world!")

        builder.list_format.list = doc.lists.add(doc.styles.get_by_name("MyListStyle1"))
        builder.writeln("Item 1")
        builder.writeln("Item 2")

        doc.cleanup()

        self.assertEqual(6, doc.styles.count)

        # Removing every node that a custom style is applied to marks it as "unused" again.
        # Run the "cleanup" method again to remove them.
        doc.first_section.body.remove_all_children()
        doc.cleanup()

        self.assertEqual(4, doc.styles.count)
        # ExEnd

    def test_automatically_update_styles(self):

        # ExStart
        # ExFor:Document.automatically_update_styles
        # ExSummary:Shows how to attach a template to a document.
        doc = aw.Document()

        # Microsoft Word documents by default come with an attached template called "Normal.dotm".
        # There is no default template for blank Aspose.Words documents.
        self.assertEqual("", doc.attached_template)

        # Attach a template, then set the flag to apply style changes
        # within the template to styles in our document.
        doc.attached_template = MY_DIR + "Business brochure.dotx"
        doc.automatically_update_styles = True

        doc.save(ARTIFACTS_DIR + "Document.automatically_update_styles.docx")
        # ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Document.automatically_update_styles.docx")

        self.assertTrue(doc.automatically_update_styles)
        self.assertEqual(MY_DIR + "Business brochure.dotx", doc.attached_template)
        self.assertTrue(os.path.exists(doc.attached_template))

    def test_default_template(self):

        # ExStart
        # ExFor:Document.attached_template
        # ExFor:Document.automatically_update_styles
        # ExFor:SaveOptions.create_save_options(str)
        # ExFor:SaveOptions.default_template
        # ExSummary:Shows how to set a default template for documents that do not have attached templates.
        doc = aw.Document()

        # Enable automatic style updating, but do not attach a template document.
        doc.automatically_update_styles = True

        self.assertEqual("", doc.attached_template)

        # Since there is no template document, the document had nowhere to track style changes.
        # Use a SaveOptions object to automatically set a template
        # if a document that we are saving does not have one.
        options = aw.saving.SaveOptions.create_save_options("Document.default_template.docx")
        options.default_template = MY_DIR + "Business brochure.dotx"

        doc.save(ARTIFACTS_DIR + "Document.default_template.docx", options)
        # ExEnd

        self.assertTrue(os.path.exists(options.default_template))

    def test_use_substitutions(self):

        # ExStart
        # ExFor:FindReplaceOptions.use_substitutions
        # ExFor:FindReplaceOptions.legacy_mode
        # ExSummary:Shows how to recognize and use substitutions within replacement patterns.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Jason gave money to Paul.")

        options = aw.replacing.FindReplaceOptions()
        options.use_substitutions = True

        # Using legacy mode does not support many advanced features, so we need to set it to 'False'.
        options.legacy_mode = False

        doc.range.replace_regex(r"([A-z]+) gave money to ([A-z]+)", r"$2 took money from $1", options)

        self.assertEqual(doc.get_text(), "Paul took money from Jason.\f")
        # ExEnd

    def test_set_invalidate_field_types(self):

        # ExStart
        # ExFor:Document.normalize_field_types
        # ExSummary:Shows how to get the keep a field's type up to date with its field code.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field("DATE", None)

        # Aspose.Words automatically detects field types based on field codes.
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)

        # Manually change the raw text of the field, which determines the field code.
        field_text = doc.first_section.body.first_paragraph.get_child_nodes(aw.NodeType.RUN, True)[0].as_run()
        self.assertEqual("DATE", field_text.text)  # ExSkip
        field_text.text = "PAGE"

        # Changing the field code has changed this field to one of a different type,
        # but the field's type properties still display the old type.
        self.assertEqual("PAGE", field.get_field_code())
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.start.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.separator.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.end.field_type)

        # Update those properties with this method to display current value.
        doc.normalize_field_types()

        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.type)
        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.start.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.separator.field_type)
        self.assertEqual(aw.fields.FieldType.FIELD_PAGE, field.end.field_type)
        # ExEnd

    def test_layout_options_revisions(self):

        # ExStart
        # ExFor:Document.layout_options
        # ExFor:LayoutOptions
        # ExFor:LayoutOptions.revision_options
        # ExFor:RevisionColor
        # ExFor:RevisionOptions
        # ExFor:RevisionOptions.inserted_text_color
        # ExFor:RevisionOptions.show_revision_bars
        # ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a revision, then change the color of all revisions to green.
        builder.writeln("This is not a revision.")
        doc.start_track_revisions("John Doe", datetime.now())
        self.assertEqual(aw.layout.RevisionColor.BY_AUTHOR,
                         doc.layout_options.revision_options.inserted_text_color)  # ExSkip
        self.assertTrue(doc.layout_options.revision_options.show_revision_bars)  # ExSkip
        builder.writeln("This is a revision.")
        doc.stop_track_revisions()
        builder.writeln("This is not a revision.")

        # Remove the bar that appears to the left of every revised line.
        doc.layout_options.revision_options.inserted_text_color = aw.layout.RevisionColor.BRIGHT_GREEN
        doc.layout_options.revision_options.show_revision_bars = False

        doc.save(ARTIFACTS_DIR + "Document.layout_options_revisions.pdf")
        # ExEnd

    def test_layout_options_hidden_text(self):

        for show_hidden_text in (False, True):
            with self.subTest(show_hidden_text=show_hidden_text):
                # ExStart
                # ExFor:Document.layout_options
                # ExFor:LayoutOptions
                # ExFor:LayoutOptions.show_hidden_text
                # ExSummary:Shows how to hide text in a rendered output document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                self.assertFalse(doc.layout_options.show_hidden_text)  # ExSkip

                # Insert hidden text, then specify whether we wish to omit it from a rendered document.
                builder.writeln("This text is not hidden.")
                builder.font.hidden = True
                builder.writeln("This text is hidden.")

                doc.layout_options.show_hidden_text = show_hidden_text

                doc.save(ARTIFACTS_DIR + "Document.layout_options_hidden_text.pdf")
                # ExEnd

        # pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "Document.layout_options_hidden_text.pdf")
        # text_absorber = aspose.pdf.text.TextAbsorber()
        # text_absorber.visit(pdf_doc)

        # if show_hidden_text:
        #    self.assertEqual("This text is not hidden.\nThis text is hidden.", text_absorber.text)
        # else:
        #    self.assertEqual("This text is not hidden.", text_absorber.text)

    def test_layout_options_paragraph_marks(self):

        for show_paragraph_marks in (False, True):
            with self.subTest(show_paragraph_marks=show_paragraph_marks):
                # ExStart
                # ExFor:Document.layout_options
                # ExFor:LayoutOptions
                # ExFor:LayoutOptions.show_paragraph_marks
                # ExSummary:Shows how to show paragraph marks in a rendered output document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                self.assertFalse(doc.layout_options.show_paragraph_marks)  # ExSkip

                # Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
                # with a pilcrow (¶) symbol when we render the document.
                builder.writeln("Hello world!")
                builder.writeln("Hello again!")

                doc.layout_options.show_paragraph_marks = show_paragraph_marks

                doc.save(ARTIFACTS_DIR + "Document.layout_options_paragraph_marks.pdf")
                # ExEnd

            # pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "Document.layout_options_paragraph_marks.pdf")
            # text_absorber = aspose.pdf.text.TextAbsorber()
            # text_absorber.visit(pdf_doc)

            # self.assertEqual("Hello world!¶\nHello again!¶\n¶" if show_paragraph_marks "Hello world!\nHello again!", text_absorber.text)

    def test_update_page_layout(self):

        # ExStart
        # ExFor:StyleCollection.__getitem__(str)
        # ExFor:SectionCollection.__getitem__(int)
        # ExFor:Document.update_page_layout
        # ExFor:PageSetup.margins
        # ExSummary:Shows when to recalculate the page layout of the document.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # Saving a document to PDF, to an image, or printing for the first time will automatically
        # cache the layout of the document within its pages.
        doc.save(ARTIFACTS_DIR + "Document.update_page_layout.1.pdf")

        # Modify the document in some way.
        doc.styles.get_by_name("Normal").font.size = 6
        doc.sections[0].page_setup.orientation = aw.Orientation.LANDSCAPE
        doc.sections[0].page_setup.margins = aw.Margins.MIRRORED

        # In the current version of Aspose.Words, modifying the document does not automatically rebuild
        # the cached page layout. If we wish for the cached layout
        # to stay up to date, we will need to update it manually.
        doc.update_page_layout()

        doc.save(ARTIFACTS_DIR + "Document.update_page_layout.2.pdf")
        # ExEnd

    def test_doc_package_custom_parts(self):

        # ExStart
        # ExFor:CustomPart
        # ExFor:CustomPart.content_type
        # ExFor:CustomPart.relationship_type
        # ExFor:CustomPart.is_external
        # ExFor:CustomPart.data
        # ExFor:CustomPart.name
        # ExFor:CustomPart.clone
        # ExFor:CustomPartCollection
        # ExFor:CustomPartCollection.add(CustomPart)
        # ExFor:CustomPartCollection.clear
        # ExFor:CustomPartCollection.clone
        # ExFor:CustomPartCollection.count
        # ExFor:CustomPartCollection.__iter__
        # ExFor:CustomPartCollection.__getitem__(int)
        # ExFor:CustomPartCollection.remove_at(int)
        # ExFor:Document.package_custom_parts
        # ExSummary:Shows how to access a document's arbitrary custom parts collection.
        doc = aw.Document(MY_DIR + "Custom parts OOXML package.docx")

        self.assertEqual(2, doc.package_custom_parts.count)

        # Clone the second part, then add the clone to the collection.
        cloned_part = doc.package_custom_parts[1].clone()
        doc.package_custom_parts.add(cloned_part)
        self._test_doc_package_custom_parts(doc.package_custom_parts)  # ExSkip

        self.assertEqual(3, doc.package_custom_parts.count)

        # Enumerate over the collection and print every part.
        for index, part in enumerate(doc.package_custom_parts):
            print(f"Part index {index}:")
            print(f"\tName:\t\t\t\t{part.name}")
            print(f"\tContent type:\t\t{part.content_type}")
            print(f"\tRelationship type:\t{part.relationship_type}")
            if part.is_external:
                print("\tSourced from outside the document")
            else:
                print(f"\tStored within the document, length: {len(part.data)} bytes")

        # We can remove elements from this collection individually, or all at once.
        doc.package_custom_parts.remove_at(2)

        self.assertEqual(2, doc.package_custom_parts.count)

        doc.package_custom_parts.clear()

        self.assertEqual(0, doc.package_custom_parts.count)
        # ExEnd

    def _test_doc_package_custom_parts(self, parts: aw.markup.CustomPartCollection):

        self.assertEqual(3, parts.count)

        self.assertEqual("/payload/payload_on_package.test", parts[0].name)
        self.assertEqual("mytest/somedata", parts[0].content_type)
        self.assertEqual("http://mytest.payload.internal", parts[0].relationship_type)
        self.assertEqual(False, parts[0].is_external)
        self.assertEqual(18, len(parts[0].data))

        self.assertEqual("http://www.aspose.com/Images/aspose-logo.jpg", parts[1].name)
        self.assertEqual("", parts[1].content_type)
        self.assertEqual("http://mytest.payload.external", parts[1].relationship_type)
        self.assertTrue(parts[1].is_external)
        self.assertEqual(0, len(parts[1].data))

        self.assertEqual("http://www.aspose.com/Images/aspose-logo.jpg", parts[2].name)
        self.assertEqual("", parts[2].content_type)
        self.assertEqual("http://mytest.payload.external", parts[2].relationship_type)
        self.assertTrue(parts[2].is_external)
        self.assertEqual(0, len(parts[2].data))

    def test_shade_form_data(self):

        for use_grey_shading in (False, True):
            with self.subTest(use_grey_shading=use_grey_shading):
                # ExStart
                # ExFor:Document.shade_form_data
                # ExSummary:Shows how to apply gray shading to form fields.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                self.assertTrue(doc.shade_form_data)  # ExSkip

                builder.write("Hello world! ")
                builder.insert_text_input("My form field", aw.fields.TextFormFieldType.REGULAR, "",
                                          "Text contents of form field, which are shaded in grey by default.", 0)

                # We can turn the grey shading off, so the bookmarked text will blend in with the other text.
                doc.shade_form_data = use_grey_shading
                doc.save(ARTIFACTS_DIR + "Document.shade_form_data.docx")
                # ExEnd

    def test_versions_count(self):

        # ExStart
        # ExFor:Document.versions_count
        # ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
        doc = aw.Document(MY_DIR + "Versions.doc")

        # We can read this property of a document, but we cannot preserve it while saving.
        self.assertEqual(4, doc.versions_count)

        doc.save(ARTIFACTS_DIR + "Document.versions_count.doc")
        doc = aw.Document(ARTIFACTS_DIR + "Document.versions_count.doc")

        self.assertEqual(0, doc.versions_count)
        # ExEnd

    def test_write_protection(self):

        # ExStart
        # ExFor:Document.write_protection
        # ExFor:WriteProtection
        # ExFor:WriteProtection.is_write_protected
        # ExFor:WriteProtection.read_only_recommended
        # ExFor:WriteProtection.set_password(str)
        # ExFor:WriteProtection.validate_password(str)
        # ExSummary:Shows how to protect a document with a password.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world! This document is protected.")
        self.assertFalse(doc.write_protection.is_write_protected)  # ExSkip
        self.assertFalse(doc.write_protection.read_only_recommended)  # ExSkip

        # Enter a password up to 15 characters in length, and then verify the document's protection status.
        doc.write_protection.set_password("MyPassword")
        doc.write_protection.read_only_recommended = True

        self.assertTrue(doc.write_protection.is_write_protected)
        self.assertTrue(doc.write_protection.validate_password("MyPassword"))

        # Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
        doc.save(ARTIFACTS_DIR + "Document.write_protection.docx")
        doc = aw.Document(ARTIFACTS_DIR + "Document.write_protection.docx")

        self.assertTrue(doc.write_protection.is_write_protected)

        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()
        builder.writeln("Writing text in a protected document.")

        self.assertEqual("Hello world! This document is protected." +
                         "\rWriting text in a protected document.", doc.get_text().strip())
        # ExEnd
        self.assertTrue(doc.write_protection.read_only_recommended)
        self.assertTrue(doc.write_protection.validate_password("MyPassword"))
        self.assertFalse(doc.write_protection.validate_password("wrongpassword"))

    def test_remove_personal_information(self):

        for save_without_personal_info in (False, True):
            with self.subTest(save_without_personal_info=save_without_personal_info):
                # ExStart
                # ExFor:Document.remove_personal_information
                # ExSummary:Shows how to enable the removal of personal information during a manual save.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert some content with personal information.
                doc.built_in_document_properties.author = "John Doe"
                doc.built_in_document_properties.company = "Placeholder Inc."

                doc.start_track_revisions(doc.built_in_document_properties.author, datetime.now())
                builder.write("Hello world!")
                doc.stop_track_revisions()

                # This flag is equivalent to File -> Options -> Trust Center -> Trust Center Settings... ->
                # Privacy Options -> "Remove personal information from file properties on save" in Microsoft Word.
                doc.remove_personal_information = save_without_personal_info

                # This option will not take effect during a save operation made using Aspose.Words.
                # Personal data will be removed from our document with the flag set when we save it manually using Microsoft Word.
                doc.save(ARTIFACTS_DIR + "Document.remove_personal_information.docx")
                doc = aw.Document(ARTIFACTS_DIR + "Document.remove_personal_information.docx")

                self.assertEqual(save_without_personal_info, doc.remove_personal_information)
                self.assertEqual("John Doe", doc.built_in_document_properties.author)
                self.assertEqual("Placeholder Inc.", doc.built_in_document_properties.company)
                self.assertEqual("John Doe", doc.revisions[0].author)
                # ExEnd

    def test_show_comments(self):

        # ExStart
        # ExFor:LayoutOptions.comment_display_mode
        # ExFor:CommentDisplayMode
        # ExSummary:Shows how to show comments when saving a document to a rendered format.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        comment = aw.Comment(doc, "John Doe", "J.D.", datetime.now())
        comment.set_text("My comment.")
        builder.current_paragraph.append_child(comment)

        # SHOW_IN_ANNOTATIONS is only available in Pdf1.7 and Pdf1.5 formats.
        # In other formats, it will work similarly to Hide.
        doc.layout_options.comment_display_mode = aw.layout.CommentDisplayMode.SHOW_IN_ANNOTATIONS

        doc.save(ARTIFACTS_DIR + "Document.show_comments_in_annotations.pdf")

        # Note that it's required to rebuild the document page layout (via Document.update_page_layout() method)
        # after changing the "Document.layout_options" values.
        doc.layout_options.comment_display_mode = aw.layout.CommentDisplayMode.SHOW_IN_BALLOONS
        doc.update_page_layout()

        doc.save(ARTIFACTS_DIR + "Document.show_comments_in_balloons.pdf")
        # ExEnd

        # pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "Document.show_comments_in_balloons.pdf")
        # text_absorber = aspose.pdf.text.TextAbsorber()
        # text_absorber.visit(pdf_doc)

        # self.assertEqual(
        #    "Hello world!                                                                    Commented [J.D.1]:  My comment.",
        #    text_absorber.text)

    def test_copy_template_styles_via_document(self):

        # ExStart
        # ExFor:Document.copy_styles_from_template(Document)
        # ExSummary:Shows how to copies styles from the template to a document via Document.
        template = aw.Document(MY_DIR + "Rendering.docx")
        target = aw.Document(MY_DIR + "Document.docx")

        self.assertEqual(18, template.styles.count)  # ExSkip
        self.assertEqual(12, target.styles.count)  # ExSkip

        target.copy_styles_from_template(template)
        self.assertEqual(22, target.styles.count)  # ExSkip

        # ExEnd

    def test_copy_template_styles_via_document_new(self):

        # ExStart
        # ExFor:Document.copy_styles_from_template(Document)
        # ExFor:Document.copy_styles_from_template(str)
        # ExSummary:Shows how to copy styles from one document to another.
        # Create a document, and then add styles that we will copy to another document.
        template = aw.Document()

        style = template.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle1")
        style.font.name = "Times New Roman"
        style.font.color = drawing.Color.navy

        style = template.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle2")
        style.font.name = "Arial"
        style.font.color = drawing.Color.deep_sky_blue

        style = template.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle3")
        style.font.name = "Courier New"
        style.font.color = drawing.Color.royal_blue

        self.assertEqual(7, template.styles.count)

        # Create a document which we will copy the styles to.
        target = aw.Document()

        # Create a style with the same name as a style from the template document and add it to the target document.
        style = target.styles.add(aw.StyleType.PARAGRAPH, "TemplateStyle3")
        style.font.name = "Calibri"
        style.font.color = drawing.Color.orange

        self.assertEqual(5, target.styles.count)

        # There are two ways of calling the method to copy all the styles from one document to another.
        # 1 -  Passing the template document object:
        target.copy_styles_from_template(template)

        # Copying styles adds all styles from the template document to the target
        # and overwrites existing styles with the same name.
        self.assertEqual(7, target.styles.count)

        self.assertEqual("Courier New", target.styles.get_by_name("TemplateStyle3").font.name)
        self.assertEqual(drawing.Color.royal_blue.to_argb(),
                         target.styles.get_by_name("TemplateStyle3").font.color.to_argb())

        # 2 -  Passing the local system filename of a template document:
        target.copy_styles_from_template(MY_DIR + "Rendering.docx")

        self.assertEqual(21, target.styles.count)
        # ExEnd

    def test_read_macros_from_existing_document(self):

        # ExStart
        # ExFor:Document.vba_project
        # ExFor:VbaModuleCollection
        # ExFor:VbaModuleCollection.count
        # ExFor:VbaModuleCollection.__getitem__(int)
        # ExFor:VbaModuleCollection.__getitem__(string)
        # ExFor:VbaModuleCollection.remove
        # ExFor:VbaModule
        # ExFor:VbaModule.name
        # ExFor:VbaModule.source_code
        # ExFor:VbaProject
        # ExFor:VbaProject.name
        # ExFor:VbaProject.modules
        # ExFor:VbaProject.code_page
        # ExFor:VbaProject.is_signed
        # ExSummary:Shows how to access a document's VBA project information.
        doc = aw.Document(MY_DIR + "VBA project.docm")

        # A VBA project contains a collection of VBA modules.
        vba_project = doc.vba_project
        self.assertTrue(vba_project.is_signed)  # ExSkip
        if vba_project.is_signed:
            print(
                f"Project name: {vba_project.name} signed; Project code page: {vba_project.code_page}; Modules count: {vba_project.modules.count}\n")
        else:
            print(
                f"Project name: {vba_project.name} not signed; Project code page: {vba_project.code_page}; Modules count: {vba_project.modules.count}\n")

        vba_modules = doc.vba_project.modules

        self.assertEqual(vba_modules.count, 3)

        for module in vba_modules:
            print(f"Module name: {module.name};\nModule code:\n{module.source_code}\n")

        # Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
        vba_modules[0].source_code = "Your VBA code..."
        vba_modules.get_by_name("Module1").source_code = "Your VBA code..."

        # Remove a module from the collection.
        vba_modules.remove(vba_modules[2])
        # ExEnd

        self.assertEqual("AsposeVBAtest", vba_project.name)
        self.assertEqual(2, vba_project.modules.count)
        self.assertEqual(1251, vba_project.code_page)
        self.assertFalse(vba_project.is_signed)

        self.assertEqual("ThisDocument", vba_modules[0].name)
        self.assertEqual("Your VBA code...", vba_modules[0].source_code)

        self.assertEqual("Module1", vba_modules[1].name)
        self.assertEqual("Your VBA code...", vba_modules[1].source_code)

    def test_save_output_parameters(self):

        # ExStart
        # ExFor:SaveOutputParameters
        # ExFor:SaveOutputParameters.content_type
        # ExSummary:Shows how to access output parameters of a document's save operation.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
        parameters = doc.save(ARTIFACTS_DIR + "Document.save_output_parameters.doc")

        self.assertEqual("application/msword", parameters.content_type)

        # This property changes depending on the save format.
        parameters = doc.save(ARTIFACTS_DIR + "Document.save_output_parameters.pdf")

        self.assertEqual("application/pdf", parameters.content_type)
        # ExEnd

    def test_sub_document(self):

        # ExStart
        # ExFor:SubDocument
        # ExFor:SubDocument.node_type
        # ExSummary:Shows how to access a master document's subdocument.
        doc = aw.Document(MY_DIR + "Master document.docx")

        sub_documents = doc.get_child_nodes(aw.NodeType.SUB_DOCUMENT, True)
        self.assertEqual(1, sub_documents.count)  # ExSkip

        # This node serves as a reference to an external document, and its contents cannot be accessed.
        sub_document = sub_documents[0].as_sub_document()

        self.assertFalse(sub_document.is_composite)
        # ExEnd

    def test_create_web_extension(self):

        # ExStart
        # ExFor:BaseWebExtensionCollection.add()
        # ExFor:BaseWebExtensionCollection.clear
        # ExFor:TaskPane
        # ExFor:TaskPane.dock_state
        # ExFor:TaskPane.is_visible
        # ExFor:TaskPane.width
        # ExFor:TaskPane.is_locked
        # ExFor:TaskPane.web_extension
        # ExFor:TaskPane.row
        # ExFor:WebExtension
        # ExFor:WebExtension.reference
        # ExFor:WebExtension.properties
        # ExFor:WebExtension.bindings
        # ExFor:WebExtension.is_frozen
        # ExFor:WebExtensionReference.id
        # ExFor:WebExtensionReference.version
        # ExFor:WebExtensionReference.store_type
        # ExFor:WebExtensionReference.store
        # ExFor:WebExtensionPropertyCollection
        # ExFor:WebExtensionBindingCollection
        # ExFor:WebExtensionProperty.__init__(str,str)
        # ExFor:WebExtensionBinding.__init__(str,WebExtensionBindingType,str)
        # ExFor:WebExtensionStoreType
        # ExFor:WebExtensionBindingType
        # ExFor:TaskPaneDockState
        # ExFor:TaskPaneCollection
        # ExSummary:Shows how to add a web extension to a document.
        doc = aw.Document()

        # Create task pane with "MyScript" add-in, which will be used by the document,
        # then set its default location.
        my_script_task_pane = aw.webextensions.TaskPane()
        doc.web_extension_task_panes.add(my_script_task_pane)
        my_script_task_pane.dock_state = aw.webextensions.TaskPaneDockState.RIGHT
        my_script_task_pane.is_visible = True
        my_script_task_pane.width = 300
        my_script_task_pane.is_locked = True

        # If there are multiple task panes in the same docking location, we can set this index to arrange them.
        my_script_task_pane.row = 1

        # Create an add-in called "MyScript Math Sample", which the task pane will display within.
        web_extension = my_script_task_pane.web_extension

        # Set application store reference parameters for our add-in, such as the ID.
        web_extension.reference.id = "WA104380646"
        web_extension.reference.version = "1.0.0.0"
        web_extension.reference.store_type = aw.webextensions.WebExtensionStoreType.OMEX
        web_extension.reference.store = "en-US"
        web_extension.properties.add(aw.webextensions.WebExtensionProperty("MyScript", "MyScript Math Sample"))
        web_extension.bindings.add(
            aw.webextensions.WebExtensionBinding("MyScript", aw.webextensions.WebExtensionBindingType.TEXT,
                                                 "104380646"))

        # Allow the user to interact with the add-in.
        web_extension.is_frozen = False

        # We can access the web extension in Microsoft Word via Developer -> Add-ins.
        doc.save(ARTIFACTS_DIR + "Document.create_web_extension.docx")

        # Remove all web extension task panes at once like this.
        doc.web_extension_task_panes.clear()

        self.assertEqual(0, doc.web_extension_task_panes.count)
        # ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Document.create_web_extension.docx")
        my_script_task_pane = doc.web_extension_task_panes[0]

        self.assertEqual(aw.webextensions.TaskPaneDockState.RIGHT, my_script_task_pane.dock_state)
        self.assertTrue(my_script_task_pane.is_visible)
        self.assertEqual(300.0, my_script_task_pane.width)
        self.assertTrue(my_script_task_pane.is_locked)
        self.assertEqual(1, my_script_task_pane.row)
        web_extension = my_script_task_pane.web_extension

        self.assertEqual("WA104380646", web_extension.reference.id)
        self.assertEqual("1.0.0.0", web_extension.reference.version)
        self.assertEqual(aw.webextensions.WebExtensionStoreType.OMEX, web_extension.reference.store_type)
        self.assertEqual("en-US", web_extension.reference.store)

        self.assertEqual("MyScript", web_extension.properties[0].name)
        self.assertEqual("MyScript Math Sample", web_extension.properties[0].value)

        self.assertEqual("MyScript", web_extension.bindings[0].id)
        self.assertEqual(aw.webextensions.WebExtensionBindingType.TEXT, web_extension.bindings[0].binding_type)
        self.assertEqual("104380646", web_extension.bindings[0].app_ref)

        self.assertFalse(web_extension.is_frozen)

    def test_get_web_extension_info(self):

        # ExStart
        # ExFor:BaseWebExtensionCollection
        # ExFor:BaseWebExtensionCollection.__iter__
        # ExFor:BaseWebExtensionCollection.remove(int)
        # ExFor:BaseWebExtensionCollection.count
        # ExFor:BaseWebExtensionCollection.__getitem__(int)
        # ExSummary:Shows how to work with a document's collection of web extensions.
        doc = aw.Document(MY_DIR + "Web extension.docx")

        self.assertEqual(1, doc.web_extension_task_panes.count)

        # print all properties of the document's web extension.
        web_extension_property_collection = doc.web_extension_task_panes[0].web_extension.properties
        for web_extension_property in web_extension_property_collection:
            print(f"Binding name: {web_extension_property.name}; Binding value: {web_extension_property.value}")

        # Remove the web extension.
        doc.web_extension_task_panes.remove(0)

        self.assertEqual(0, doc.web_extension_task_panes.count)
        # ExEnd

    def test_epub_cover(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
        doc.built_in_document_properties.author = "John Doe"
        doc.built_in_document_properties.title = "My Book Title"

        # The thumbnail we specify here can become the cover image.
        with open(IMAGE_DIR + "Transparent background logo.png", "rb") as file:
            image = file.read()

        doc.built_in_document_properties.thumbnail = image

        doc.save(ARTIFACTS_DIR + "Document.epub_cover.epub")

    def test_text_watermark(self):

        # ExStart
        # ExFor:Watermark.set_text(str)
        # ExFor:Watermark.set_text(str,TextWatermarkOptions)
        # ExFor:Watermark.remove
        # ExFor:TextWatermarkOptions.font_family
        # ExFor:TextWatermarkOptions.font_size
        # ExFor:TextWatermarkOptions.color
        # ExFor:TextWatermarkOptions.layout
        # ExFor:TextWatermarkOptions.is_semitrasparent
        # ExFor:WatermarkLayout
        # ExFor:WatermarkType
        # ExSummary:Shows how to create a text watermark.
        doc = aw.Document()

        # Add a plain text watermark.
        doc.watermark.set_text("Aspose Watermark")

        # If we wish to edit the text formatting using it as a watermark,
        # we can do so by passing a TextWatermarkOptions object when creating the watermark.
        text_watermark_options = aw.TextWatermarkOptions()
        text_watermark_options.font_family = "Arial"
        text_watermark_options.font_size = 36
        text_watermark_options.color = drawing.Color.black
        text_watermark_options.layout = aw.WatermarkLayout.DIAGONAL
        text_watermark_options.is_semitrasparent = False

        doc.watermark.set_text("Aspose Watermark", text_watermark_options)

        doc.save(ARTIFACTS_DIR + "Document.text_watermark.docx")

        # We can remove a watermark from a document like this.
        if doc.watermark.type == aw.WatermarkType.TEXT:
            doc.watermark.remove()
        # ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Document.text_watermark.docx")

        self.assertEqual(aw.WatermarkType.TEXT, doc.watermark.type)

    @unittest.skip("drawing.Image type isn't supported yet")
    def test_image_watermark(self):

        # ExStart
        # ExFor:Watermark.set_image(Image,ImageWatermarkOptions)
        # ExFor:ImageWatermarkOptions.scale
        # ExFor:ImageWatermarkOptions.is_washout
        # ExSummary:Shows how to create a watermark from an image in the local file system.
        doc = aw.Document()

        # Modify the image watermark's appearance with an ImageWatermarkOptions object,
        # then pass it while creating a watermark from an image file.
        image_watermark_options = aw.ImageWatermarkOptions()
        image_watermark_options.scale = 5
        image_watermark_options.is_washout = False

        doc.watermark.set_image(drawing.Image.from_file(IMAGE_DIR + "Logo.jpg"), image_watermark_options)

        doc.save(ARTIFACTS_DIR + "Document.image_watermark.docx")
        # ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Document.image_watermark.docx")

        self.assertEqual(aw.WatermarkType.IMAGE, doc.watermark.type)

    def test_spelling_and_grammar_errors(self):

        for show_errors in (False, True):
            with self.subTest(show_errors=show_errors):
                # ExStart
                # ExFor:Document.show_grammatical_errors
                # ExFor:Document.show_spelling_errors
                # ExSummary:Shows how to show/hide errors in the document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert two sentences with mistakes that would be picked up
                # by the spelling and grammar checkers in Microsoft Word.
                builder.writeln("There is a speling error in this sentence.")
                builder.writeln("Their is a grammatical error in this sentence.")

                # If these options are enabled, then spelling errors will be underlined
                # in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
                doc.show_grammatical_errors = show_errors
                doc.show_spelling_errors = show_errors

                doc.save(ARTIFACTS_DIR + "Document.spelling_and_grammar_errors.docx")
                # ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Document.spelling_and_grammar_errors.docx")

                self.assertEqual(show_errors, doc.show_grammatical_errors)
                self.assertEqual(show_errors, doc.show_spelling_errors)

    def test_granularity_compare_option(self):

        for granularity in (aw.comparing.Granularity.CHAR_LEVEL,
                            aw.comparing.Granularity.WORD_LEVEL):
            with self.subTest(granularity=granularity):
                # ExStart
                # ExFor:CompareOptions.granularity
                # ExFor:Granularity
                # ExSummary:Shows to specify a granularity while comparing documents.
                doc_a = aw.Document()
                builder_a = aw.DocumentBuilder(doc_a)
                builder_a.writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit")

                doc_b = aw.Document()
                builder_b = aw.DocumentBuilder(doc_b)
                builder_b.writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit")

                # Specify whether changes are tracking
                # by character ('Granularity.CHAR_LEVEL'), or by word ('Granularity.WORD_LEVEL').
                compare_options = aw.comparing.CompareOptions()
                compare_options.granularity = granularity

                doc_a.compare(doc_b, "author", datetime.now(), compare_options)

                # The first document's collection of revision groups contains all the differences between documents.
                groups = doc_a.revisions.groups
                self.assertEqual(5, groups.count)
                # ExEnd

                if granularity == aw.comparing.Granularity.CHAR_LEVEL:
                    self.assertEqual(aw.RevisionType.DELETION, groups[0].revision_type)
                    self.assertEqual("Alpha ", groups[0].text)

                    self.assertEqual(aw.RevisionType.DELETION, groups[1].revision_type)
                    self.assertEqual(",", groups[1].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[2].revision_type)
                    self.assertEqual("s", groups[2].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[3].revision_type)
                    self.assertEqual("- \"", groups[3].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[4].revision_type)
                    self.assertEqual("\"", groups[4].text)
                else:
                    self.assertEqual(aw.RevisionType.DELETION, groups[0].revision_type)
                    self.assertEqual("Alpha Lorem", groups[0].text)

                    self.assertEqual(aw.RevisionType.DELETION, groups[1].revision_type)
                    self.assertEqual(",", groups[1].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[2].revision_type)
                    self.assertEqual("Lorems", groups[2].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[3].revision_type)
                    self.assertEqual("- \"", groups[3].text)

                    self.assertEqual(aw.RevisionType.INSERTION, groups[4].revision_type)
                    self.assertEqual("\"", groups[4].text)

    def test_ignore_printer_metrics(self):

        # ExStart
        # ExFor:LayoutOptions.ignore_printer_metrics
        # ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        doc.layout_options.ignore_printer_metrics = False

        doc.save(ARTIFACTS_DIR + "Document.ignore_printer_metrics.docx")
        # ExEnd

    def test_extract_pages(self):

        # ExStart
        # ExFor:Document.extract_pages
        # ExSummary:Shows how to get specified range of pages from the document.
        doc = aw.Document(MY_DIR + "Layout entities.docx")

        doc = doc.extract_pages(0, 2)

        doc.save(ARTIFACTS_DIR + "Document.extract_pages.docx")
        # ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Document.extract_pages.docx")
        self.assertEqual(doc.page_count, 2)

    def test_spelling_or_grammar(self):

        for check_spelling_grammar in (True, False):
            with self.subTest(check_spelling_grammar=check_spelling_grammar):
                # ExStart
                # ExFor:Document.spelling_checked
                # ExFor:Document.grammar_checked
                # ExSummary:Shows how to set spelling or grammar verifying.
                doc = aw.Document()

                # The string with spelling errors.
                doc.first_section.body.first_paragraph.runs.add(
                    aw.Run(doc, "The speeling in this documentz is all broked."))

                # Spelling/Grammar check start if we set properties to False.
                # We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
                # Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
                doc.spelling_checked = check_spelling_grammar
                doc.grammar_checked = check_spelling_grammar

                doc.save(ARTIFACTS_DIR + "Document.spelling_or_grammar.docx")
                # ExEnd

    def test_allow_embedding_post_script_fonts(self):

        # ExStart
        # ExFor:SaveOptions.allow_embedding_post_script_fonts
        # ExSummary:Shows how to save the document with PostScript font.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.name = "PostScriptFont"
        builder.writeln("Some text with PostScript font.")

        # Load the font with PostScript to use in the document.
        with open(FONTS_DIR + "AllegroOpen.otf", "rb") as file:
            otf = aw.fonts.MemoryFontSource(file.read())

        doc.font_settings = aw.fonts.FontSettings()
        doc.font_settings.set_fonts_sources([otf])

        # Embed TrueType fonts.
        doc.font_infos.embed_true_type_fonts = True

        # Allow embedding PostScript fonts while embedding TrueType fonts.
        # Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
        save_options = aw.saving.SaveOptions.create_save_options(aw.SaveFormat.DOCX)
        save_options.allow_embedding_post_script_fonts = True

        doc.save(ARTIFACTS_DIR + "Document.allow_embedding_post_script_fonts.docx", save_options)
        # ExEnd

    def test_frameset(self):

        # ExStart
        # ExFor:Document.frameset
        # ExFor:Frameset
        # ExFor:Frameset.frame_default_url
        # ExFor:Frameset.is_frame_link_to_file
        # ExFor:Frameset.child_framesets
        # ExSummary:Shows how to access frames on-page.
        # Document contains several frames with links to other documents.
        doc = aw.Document(MY_DIR + "Frameset.docx")

        # We can check the default URL (a web page URL or local document) or if the frame is an external resource.
        self.assertEqual("https://file-examples-com.github.io/uploads/2017/02/file-sample_100kB.docx",
                         doc.frameset.child_framesets[0].child_framesets[0].frame_default_url)
        self.assertTrue(doc.frameset.child_framesets[0].child_framesets[0].is_frame_link_to_file)

        self.assertEqual("Document.docx", doc.frameset.child_framesets[1].frame_default_url)
        self.assertFalse(doc.frameset.child_framesets[1].is_frame_link_to_file)

        # Change properties for one of our frames.
        doc.frameset.child_framesets[0].child_framesets[
            0].frame_default_url = "https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx"
        doc.frameset.child_framesets[0].child_framesets[0].is_frame_link_to_file = False
        # ExEnd

        doc = DocumentHelper.save_open(doc)

        self.assertEqual(
            "https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx",
            doc.frameset.child_framesets[0].child_framesets[0].frame_default_url)
        self.assertFalse(doc.frameset.child_framesets[0].child_framesets[0].is_frame_link_to_file)

    def test_open_azw(self):
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + "Azw3 document.azw3")
        self.assertEqual(info.load_format, aw.LoadFormat.AZW3)

        doc = aw.Document(MY_DIR + "Azw3 document.azw3")
        self.assertIn("Hachette Book Group USA", doc.get_text())

    def test_open_epub(self):
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + "Epub document.epub")
        self.assertEqual(info.load_format, aw.LoadFormat.EPUB)
        doc = aw.Document(MY_DIR + "Epub document.epub")
        self.assertTrue(doc.get_text().find("Down the Rabbit-Hole") != -1)

    def test_open_xml(self):
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + "Mail merge data - Customers.xml")
        self.assertEqual(info.load_format, aw.LoadFormat.XML)
        doc = aw.Document(MY_DIR + "Mail merge data - Purchase order.xml")
        self.assertTrue(doc.get_text().find("Ellen Adams\r123 Maple Street") != -1)

    def test_move_to_structured_document_tag(self):
        # ExStart
        # ExFor:DocumentBuilder.move_to_structured_document_tag(int, int)
        # ExFor:DocumentBuilder.move_to_structured_document_tag(StructuredDocumentTag, int)
        # ExFor:DocumentBuilder.IsAtEndOfStructuredDocumentTag
        # ExFor:DocumentBuilder.CurrentStructuredDocumentTag
        # ExSummary:Shows how to move cursor of DocumentBuilder inside a structured document tag.
        doc = aw.Document(MY_DIR + "Structured document tags.docx")
        builder = aw.DocumentBuilder(doc)

        # There is a several ways to move the cursor:
        # 1 -  Move to the first character of structured document tag by index.
        builder.move_to_structured_document_tag(1, 1)

        # 2 -  Move to the first character of structured document tag by object.
        tag = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG, 2, True).as_structured_document_tag()
        builder.move_to_structured_document_tag(tag, 1)
        builder.write(" New text.")

        self.assertEqual("R New text.ichText", tag.get_text().strip())

        # 3 -  Move to the end of the second structured document tag.
        builder.move_to_structured_document_tag(1, -1)
        self.assertTrue(builder.is_at_end_of_structured_document_tag)

        # Get currently selected structured document tag.
        builder.current_structured_document_tag.color = drawing.Color.green

        doc.save(ARTIFACTS_DIR + "Document.MoveToStructuredDocumentTag.docx")
        # ExEnd

    def test_include_textboxes_footnotes_endnotes_in_stat(self):

        # ExStart
        # ExFor:Document.include_textboxes_footnotes_endnotes_in_stat
        # ExSummary:Shows how to include or exclude textboxes, footnotes and endnotes from word count statistics.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Lorem ipsum")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "sit amet")

        # By default option is set to 'false'.
        doc.update_word_count()

        # Words count without textboxes, footnotes and endnotes.
        self.assertEqual(2, doc.built_in_document_properties.words)

        # Words count with textboxes, footnotes and endnotes.
        doc.include_textboxes_footnotes_endnotes_in_stat = True
        doc.update_word_count()

        self.assertEqual(4, doc.built_in_document_properties.words)
        # ExEnd

    def test_set_justification_mode(self):
        # ExStart
        # ExFor:Document.justification_mode
        # ExFor:JustificationMode
        # ExSummary:Shows how to manage character spacing control.
        doc = aw.Document(MY_DIR + "Document.docx")
        justification_mode = doc.justification_mode
        if justification_mode == aw.settings.JustificationMode.EXPAND:
            doc.justification_mode = aw.settings.JustificationMode.COMPRESS

        doc.save(ARTIFACTS_DIR + "Document.SetJustificationMode.docx")
        # ExEnd

    def test_adjust_sentence_and_word_spacing(self):
        # ExStart
        # ExFor:ImportFormatOptions.adjust_sentence_and_word_spacing
        # ExSummary:Shows how to adjust sentence and word spacing automatically.

        srcDoc = aw.Document()
        dstDoc = aw.Document()

        builder = aw.DocumentBuilder(srcDoc)
        builder.write("Dolor sit amet.")

        builder = aw.DocumentBuilder(dstDoc)
        builder.write("Lorem ipsum.")

        options = aw.ImportFormatOptions()
        options.adjust_sentence_and_word_spacing = True

        builder.insert_document(srcDoc, aw.ImportFormatMode.USE_DESTINATION_STYLES, options)

        self.assertEqual("Lorem ipsum. Dolor sit amet.", dstDoc.first_section.body.first_paragraph.get_text().strip())
        # ExEnd

    def test_page_is_in_color(self):

        # ExStart
        # ExFor: PageInfo.colored
        # ExSummary:Shows how to check whether the page is in color or not.
        doc = aw.Document(MY_DIR + "Document.docx")

        # Check that the first page of the document is not colored.
        self.assertFalse(doc.get_page_info(0).colored)
        # ExEnd

    def test_insert_document_inline(self):

        # ExStart
        # ExFor:DocumentBuilder.insert_document_inline(Document, ImportFormatMode, ImportFormatOptions)
        # ExSummary:Shows how to insert a document inline at the cursor position.

        src_doc = aw.DocumentBuilder()
        src_doc.write("[src content]")

        # Create destination document.

        dst_doc = aw.DocumentBuilder()
        dst_doc.write("Before ")
        dst_doc.insert_node(aw.BookmarkStart(dst_doc.document, "src_place"))
        dst_doc.insert_node(aw.BookmarkEnd(dst_doc.document, "src_place"))
        dst_doc.write(" after")

        self.assertEqual("Before  after", dst_doc.document.get_text().strip())

        # Insert source document into destination inline.
        dst_doc.move_to_bookmark("src_place")
        dst_doc.insert_document_inline(src_doc.document, aw.ImportFormatMode.USE_DESTINATION_STYLES,
                                       aw.ImportFormatOptions())

        self.assertEqual("Before [src content] after", dst_doc.document.get_text().strip())
        # ExEnd
