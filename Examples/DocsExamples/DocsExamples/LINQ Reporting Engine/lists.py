import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class Lists(docs_base.DocsExamplesBase):

    def test_create_bulleted_list(self):

        #ExStart:BulletedList
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Bulleted list.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "clients.json"), "clients")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.create_bulleted_list.docx")
        #ExEnd:BulletedList


    def test_common_list(self):

        #ExStart:CommonList
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Common master detail.docx")

        opt = aw.reporting.JsonDataLoadOptions()
        opt.always_generate_root_object = True
        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json", opt), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.common_list.docx")
        #ExEnd:CommonList


    def test_in_paragraph_list(self):

        #ExStart:InParagraphList
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("<<foreach [in clients]>><<[IndexOf() !=0 ? ”, ”:  ””]>><<[Name]>><</foreach>>")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "clients.json"), "clients")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.in_paragraph_list.docx")
        #ExEnd:InParagraphList


    def test_in_table_list(self):

        #ExStart:InTableList
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Contextual object member access.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json"), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.in_table_list.docx")
        #ExEnd:InTableList


    def test_multicolored_numbered_list(self):

        #ExStart:MulticoloredNumberedList
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Multicolored numbered list.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "clients.json"), "clients")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.multicolored_numbered_list.doc")
        #ExEnd:MulticoloredNumberedList


    def test_numbered_list(self):

        #ExStart:NumberedList
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Numbered list.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "clients.json"), "clients")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.numbered_list.docx")
        #ExEnd:NumberedList


if __name__ == '__main__':
    unittest.main()
