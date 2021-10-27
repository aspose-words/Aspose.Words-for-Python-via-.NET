import unittest
import os
import sys

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, JSON_DIR

import aspose.words as aw

class Lists(DocsExamplesBase):

    def test_create_bulleted_list(self):

        #ExStart:BulletedList
        doc = aw.Document(MY_DIR + "Reporting engine template - Bulleted list.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "clients.json"), "clients")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.create_bulleted_list.docx")
        #ExEnd:BulletedList

    def test_common_list(self):

        #ExStart:CommonList
        doc = aw.Document(MY_DIR + "Reporting engine template - Common master detail.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "managers.json", opt), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.common_list.docx")
        #ExEnd:CommonList

    def test_in_paragraph_list(self):

        #ExStart:InParagraphList
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("<<foreach [in clients]>><<[IndexOf() !=0 ? ”, ”:  ””]>><<[Name]>><</foreach>>")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "clients.json"), "clients")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.in_paragraph_list.docx")
        #ExEnd:InParagraphList

    def test_in_table_list(self):

        #ExStart:InTableList
        doc = aw.Document(MY_DIR + "Reporting engine template - Contextual object member access.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "managers.json"), "Managers")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.in_table_list.docx")
        #ExEnd:InTableList

    def test_multicolored_numbered_list(self):

        #ExStart:MulticoloredNumberedList
        doc = aw.Document(MY_DIR + "Reporting engine template - Multicolored numbered list.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "clients.json"), "clients")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.multicolored_numbered_list.doc")
        #ExEnd:MulticoloredNumberedList

    def test_numbered_list(self):

        #ExStart:NumberedList
        doc = aw.Document(MY_DIR + "Reporting engine template - Numbered list.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(JSON_DIR + "clients.json"), "clients")

        doc.save(ARTIFACTS_DIR + "ReportingEngine.numbered_list.docx")
        #ExEnd:NumberedList
