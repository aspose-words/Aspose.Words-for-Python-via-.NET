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
import uuid
import aspose.pydrawing as drawing
from io import BytesIO
import asyncio
import websockets

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR, FONTS_DIR, GOLDS_DIR
from document_helper import DocumentHelper

class ExDocument(ApiExampleBase):

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
        return_data = {
            "type": "table",
            "attrs": {
                "id": id
            },
            "content": table_data
        }
        return return_data

    def test_aw_extract_headings_and_contents_table_dict_id(self):
        doc = aw.Document(MY_DIR + "Part - Copy.docx")
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
                                "block_id": str(block_id),
                                "Content": [],
                                "Level": level,
                                "Table": [],
                                "Tbale_name": [],
                            }
                        )
                    elif data:
                        if not node.get_ancestor(aw.NodeType.TABLE):
                            if node.get_text().strip() and "Toc" not in node.get_text()\
                                            and "TOC" not in node.get_text():
                                data[-1]["Content"].append(
                                    {"type": "text",
                                     "content": node.get_text().strip().replace("  SEQ 表 \* ARABIC ", ''),
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
                        tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
                        _able_content = self.aw_read_table_id(parent_node, data[-1]["block_id"] + '&&' + str(block_id1))
                        data[-1]["Content"].append(
                            {"type": "table",
                             "content": _able_content,
                             "block_id": data[-1]["block_id"] + '&&' + str(block_id1),
                             "parent_block_id": data[-1]["block_id"]})
        while stack:
            old_level, old_data = stack.pop()
            data = old_data + data
        return data

    def test_new_template(self):

        sections = [
                    {
                        "type": "title",
                        "Title": "3 个例临床研究报告目录\r",
                        "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e",
                        "Content": [
                            {
                                "type": "text",
                                "content": "目录",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&5b7c7edb-1b91-4f32-8e21-112ba1b3420e",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "1 标题页\t1",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "2 概要\t5",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "3 个例临床研究报告目录\t22",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "表格目录\t25",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "图表目录\t26",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "4 缩略语和术语定义表\t28",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "5 伦理学\t29",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "5.1 独立伦理委员会（IEC）\t29",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "5.2 研究的伦理行为\t29",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "5.3 受试者知情与同意\t29",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "6 研究者和研究管理机构\t30",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "7 简介\t31",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "8 研究目标\t32",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "8.1 主要研究目的\t32",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "8.2 次要研究目的\t32",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "8.3 探索性研究目的\t32",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9 研究计划\t32",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.1 整体研究计划\t32",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.1.1 研究示意图\t33",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.2 研究设计讨论\t33",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.2.1患者群体的选择\t33",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.2.2主要终点采用IRC评估的ORR\t34",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.2.3 疗效评估标准的选择\t34",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.3 研究人群的选择\t35",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.3.1 入选标准\t35",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.3.2 排除标准\t36",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.3.3 从治疗或评估中移除受试者\t37",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4 治疗\t38",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.1 给予的治疗\t38",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.2 研究药物信息\t39",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.3 受试者的治疗组分配方法\t39",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.4 研究中的剂量选择\t39",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.5 每位受试者的研究剂量和给药时间\t39",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.6 盲法\t40",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.7 既往和伴随治疗\t40",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.4.8 治疗依从性\t41",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.5 疗效、安全性和药代动力学终点\t41",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.5.1 评估疗效、安全性和药代动力学终点指标和流程图\t41",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.5.2 衡量指标的适当性\t45",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.5.3 主要疗效终点\t45",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.5.4 药物浓度测定\t45",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.6 数据质量保证\t45",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.6.1临床试验过程的质量保证与控制\t46",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.6.2 启动访视\t47",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.6.3 实验室认可\t47",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.6.4 数据管理\t47",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.7 研究方案中计划的统计方法和样本量的确定\t48",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.7.1 统计分析计划\t48",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.7.2 样本量的确定\t53",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.8 研究过程或分析计划的变更\t53",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.8.1 方案变更\t53",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "9.8.2 统计分析计划（SAP）变更\t54",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10 研究对象\t55",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.1 受试者分布\t55",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.2 研究方案偏离\t56",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.3 分析数据集\t56",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.4 人口统计学和其他基线特征\t56",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.4.1 人口统计学和基线特征（mITT集）\t56",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.4.2肿瘤基线特征及既往抗肿瘤治疗（mITT集）\t58",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.4.3 既往病史\t60",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.4.4 既往和合并药物治疗\t61",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "10.5 治疗依从性的测量\t61",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11疗效评估和药代动力学评估\t61",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1疗效结果和个体受试者数据列表\t61",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1.1 疗效分析\t62",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1.2 统计/分析内容\t69",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1.3 个体疗效数据列表\t72",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1.4 药物剂量、药物浓度以及效应之间的关系\t72",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1.5 药物-药物和药物-疾病相互作用\t72",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1.6 按受试者列出\t72",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "11.1.7 疗效结论\t72",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12 安全性评价\t73",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.1 暴露程度\t73",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.2 不良事件（AE）\t74",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.2.1 治疗期间发生的不良事件（TEAE）概要\t74",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.2.2 不良事件列出\t75",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.2.3不良事件分析\t76",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.2.4 各受试者不良事件列表\t76",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.3 死亡、其他严重不良事件和其他重要不良事件\t77",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.3.1 死亡、其他严重不良事件和其他重要不良事件列表\t77",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.3.2 死亡、其他严重不良事件和其他重要不良事件的叙述\t79",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.3.3 死亡、其他严重不良事件和其他重要不良事件的分析和讨论\t79",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.4 临床实验室评估\t79",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.4.1 各受试者的个例实验测量值列表（16.2.8）和各异常实验室值（14.3.4）\t79",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.4.2 各实验室参数的评价\t79",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.5 生命体征、体格检查发现和其他安全性相关观察结果（SS）\t81",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "12.6 安全性结论\t81",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "13 讨论和总体结论\t81",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "13.1 讨论\t81",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "13.1.1 背景\t81",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "13.1.2 有效性\t82",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "13.1.3 安全性\t83",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "13.1.4 药代动力学\t83",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "13.2 结论\t83",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "14 参考但不纳入文本的表格、图示和图表\t84",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "15参考文献列表\t85",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            },
                            {
                                "type": "toc",
                                "content": "16 附录\t86",
                                "block_id": "082a0d84-d609-426e-86d8-b6c179abff0e&&236e40de-2c70-440c-a0ae-b67ccda10862",
                                "parent_block_id": "082a0d84-d609-426e-86d8-b6c179abff0e"
                            }
                        ],
                        "Level": 1,
                        "Table": [],
                        "Tbale_name": []
                    },
                    {
                        "Title": "3.1表格目录\r",
                        "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b",
                        "Content": [
                            {
                                "type": "toc",
                                "content": "表1 研究药物信息\t38",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 2方案修订情况\t52",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 3 受试者分布（所有入组受试者)\t54",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表4重大方案偏离汇总（所有入组受试者）\t55",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 5 人口统计学和基线特征（mITT集）\t56",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 6 肿瘤基线特征（mITT集）\t57",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 7 基因突变情况特征（mITT集）\t58",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 8既往抗肿瘤治疗（mITT集）\t58",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 9 基于PRC标准IRC评估的最佳总体疗效总结（mITT集）\t61",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 10 基于PRC标准研究者评估的最佳总体疗效总结（mITT集）\t63",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 11 IRC和研究者基于PRC所评估最佳总体疗效的一致性分析（mITT集）\t64",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 12 基于RECIST1.1标准IRC评估的最佳总体疗效总结（mITT集）\t65",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 13 基于RECIST1.1标准研究者评估的最佳总体疗效总结（mITT集）\t66",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 14 IRC和研究者基于RECIST1.1所评估最佳总体疗效的一致性分析（mITT）\t67",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 15 经确认的客观缓解率的亚组分析（由IRC、研究者基于PRC评估，mITT)\t70",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 16 研究药物暴露情况(SS)\t72",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 17 所有不良事件汇总（SS）\t74",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 18 其他严重不良事件（SS)\t76",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 19 按系统器官分类和首选术语分类的导致剂量降低的TEAE（SS）\t77",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            },
                            {
                                "type": "toc",
                                "content": "表 20 按系统器官分类和首选术语分类的导致暂停用药的TEAE（SS）\t77",
                                "block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b&&8d480b06-15f4-4456-a766-beaafd3750ae",
                                "parent_block_id": "7edfcd81-19fb-471b-9024-2bdc34618d3b"
                            }
                        ],
                        "Level": 2,
                        "Table": [],
                        "Tbale_name": []
                    },
                    {
                        "Title": "图表目录\r",
                        "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742",
                        "Content": [
                            {
                                "type": "toc",
                                "content": "图1总体研究设计图\t32",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "toc",
                                "content": "图 2 IRC基于PRC标准评估靶病灶SUV总和较基线变化最佳百分比的瀑布图（mITT集）\t62",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "toc",
                                "content": "图 3 IRC基于PRC标准评估靶病灶SUV总和较基线变化百分比的蜘蛛图（mITT集）\t62",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "toc",
                                "content": "图 4 IRC基于PRC标准评估治疗持续时间及肿瘤整体疗效游泳图（mITT集）\t63",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "toc",
                                "content": "图 5 总健康状况平均线性转换评分随时间变化图（mITT集）\t68",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "toc",
                                "content": "图 6 cfDNA中MAPK通路基因突变频率（AF%）的动态变化曲线图(mITT)\t68",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "toc",
                                "content": "图7 C2D1给药后平均血药浓度-时间曲线\t71",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "toc",
                                "content": "图 8 不同访视的平均谷浓度-时间曲线\t71",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&f657a024-b1de-49b0-a709-3c36b35801a7",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            },
                            {
                                "type": "text",
                                "content": "",
                                "block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742&&c1f01b9e-ce11-4c6d-ba24-a255d42a2cfd",
                                "parent_block_id": "e4441e9a-de49-43d7-8f11-ca9f04b4c742"
                            }
                        ],
                        "Level": 2,
                        "Table": [],
                        "Tbale_name": []
                    }
                ]

        # 创建一个新的文档
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

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
                new_run.bold = True
                # 添加标题内容
                title = section.get('contents', '').strip()  # 获取标题内容
                builder.writeln(title)  # 写入标题到文档

            elif section["type"] == 'text':
                # 添加文本内容
                new_run = builder.font
                # 设置西文和中文字体
                new_run.name = "Times New Roman"  # 设置西文是新罗马字体
                new_run.bold = False
                new_run.name_far_east = "宋体"
                new_run.size = 12
                text_content = section.get('contents', '').strip()  # 获取文本内容
                builder.paragraph_format.character_unit_first_line_indent = 2
                builder.paragraph_format.line_spacing = 18
                builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
                builder.paragraph_format.line_unit_after = 0
                builder.paragraph_format.space_after = 0
                if text_content.startswith("表"):
                    builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
                    new_run.bold = True
                builder.writeln(text_content)  # 写入文本到文档
            elif section["type"] == 'toc':
                # 添加文本内容
                page_setup = doc.first_section.page_setup
                tab_stop_position = page_setup.page_width - page_setup.left_margin - page_setup.right_margin
                builder.paragraph_format.tab_stops.add(tab_stop_position, aw.TabAlignment.RIGHT, aw.TabLeader.DOTS)
                text_content = section.get('content', '')
                builder.writeln(text_content)
            elif section["type"] == 'image':
                # 插入图片
                image_path = section.get('contents', '').strip()
                builder.insert_image(image_path)
            elif section["type"] == 'table':
                # 添加表格内容
                table = builder.start_table()  # 开始一个新表格
                new_run = builder.font
                new_run.bold = False
                new_run.size = 8
                #  记录rowspan的数据
                table_data = section["contents"]
                for i in range(len(table_data)):
                    cell_data_index = 0
                    for row_data in table_data[i].get('content', []):
                        # 遍历表格单元格
                        for cell_data in row_data.get('content', []):
                            new_cell = builder.insert_cell()  # 插入新的表格行
                            new_cell.cell_format.borders.line_style = aw.LineStyle.DASH_LARGE_GAP
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
                builder.insert_break(aw.BreakType.LINE_BREAK)

                for row in table.rows:
                    row = row.as_row()
                    if row.is_first_row:
                        for cell in row.cells:
                            cell = cell.as_cell()
                            cell.cell_format.borders.bottom.line_style = aw.LineStyle.SINGLE
                        for cell in row.next_row.cells:
                            cell = cell.as_cell()
                            cell.cell_format.borders.top.line_style = aw.LineStyle.SINGLE

                    if row.is_last_row:
                        for cell in row.cells:
                            cell = cell.as_cell()
                            cell.cell_format.borders.top.line_style = aw.LineStyle.SINGLE
                        for cell in row.previous_row.cells:
                            cell = cell.as_cell()
                            cell.cell_format.borders.bottom.line_style = aw.LineStyle.SINGLE

        # 保存文档
        doc.save(ARTIFACTS_DIR + "output1.docx")

