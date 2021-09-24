import unittest
import os
import sys
import io

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class BaseOperations(docs_base.DocsExamplesBase):
    
    def test_hello_world(self) :
        
        #ExStart:HelloWorld
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        builder.write("<<[sender.name]>> says: <<[sender.message]>>")

        json_data_source = aw.reporting.JsonDataSource(io.BytesIO(b"{\"Name\":\"LINQ Reporting Engine\",\"Message\":\"Hello World\"}"))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, json_data_source, "sender")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.hello_world.docx")
        #ExEnd:HelloWorld
        

    def test_single_row(self) :
        
        #ExStart:SingleRow
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Table row.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json"), "Managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.single_row.docx")
        #ExEnd:SingleRow
        

    def test_common_master_detail(self) :
        
        #ExStart:CommonMasterDetail
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Common master detail.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "managers.json"), "managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.common_master_detail.docx")
        #ExEnd:CommonMasterDetail
        

    def test_conditional_blocks(self) :
        
        #ExStart:ConditionalBlocks
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Table row conditional blocks.docx")

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, aw.reporting.JsonDataSource(docs_base.json_dir + "clients.json"), "clients")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.conditional_block.docx")
        #ExEnd:ConditionalBlocks
        

    def test_setting_background_color(self) :
        
        #ExStart:SettingBackgroundColor
        doc = aw.Document(docs_base.my_dir + "Reporting engine template - Background color.docx")

        json = b"""[
              {
                "Name": "Black",
                "Color": "black",
              },
              {
                "Name": "Red",
                "Color": "red"
              },
              {
                "Name": "Green",
                "Color": "green"
              }
            ]"""
            
        json_data_source = aw.reporting.JsonDataSource(io.BytesIO(json))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, json_data_source, "Colors")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.back_color.docx")
        #ExEnd:SettingBackgroundColor
        
    

if __name__ == '__main__':
    unittest.main()