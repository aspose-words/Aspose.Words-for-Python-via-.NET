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

    def test_xml_data_source(self) :

        xml = b"""<Person>
                    <Name>John Doe</Name>
                    <Age>30</Age>
                    <Birth>1989-04-01 4:00:00 pm</Birth>
                    <Child>Ann Doe</Child>
                    <Child>Charles Doe</Child>
                </Person>"""

        template_md = b"""**Name**: <<[Name]>>, **Age**: <<[Age]>>, **Date of Birth**:
<<[Birth]:"dd.MM.yyyy">>

**Children**:
<<foreach [in Child]>><<[Child_Text]>>
<</foreach>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)
        data_source = aw.reporting.XmlDataSource(io.BytesIO(xml))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source)

        doc.save(docs_base.artifacts_dir + "ReportingEngine.xml_data_source.docx")


    def test_xml_data_source_seq(self) :

        xml = b"""<Persons>
                    <Person>
                        <Name>John Doe</Name>
                        <Age>30</Age>
                        <Birth>1989-04-01 4:00:00 pm</Birth>
                    </Person>
                    <Person>
                        <Name>Jane Doe</Name>
                        <Age>27</Age>
                        <Birth>1992-01-31 07:00:00 am</Birth>
                    </Person>
                    <Person>
                        <Name>John Smith</Name>
                        <Age>51</Age>
                        <Birth>1968-03-08 1:00:00 pm</Birth>
                    </Person>
                </Persons>"""

        template_md = b"""<<foreach [in persons]>>**Name:** <<[Name]>>, **Age:** <<[Age]>>, **Date of Birth:** <<[Birth]:"dd.MM.yyyy">>
<</foreach>>

**Average age:** <<[persons.Average(p => p.Age)]>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)
        data_source = aw.reporting.XmlDataSource(io.BytesIO(xml))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source, "persons")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.xml_data_source_seq.docx")

    def test_xml_data_source_complex(self) :

        xml = b"""<Managers>
                    <Manager>
                        <Name>John Smith</Name>
                        <Contract>
                            <Client>
                                <Name>A Company</Name>
                            </Client>
                            <Price>1200000</Price>
                        </Contract>
                        <Contract>
                            <Client>
                                <Name>B Ltd.</Name>
                            </Client>
                            <Price>750000</Price>
                        </Contract>
                        <Contract>
                            <Client>
                                <Name>C &amp; D</Name>
                            </Client>
                            <Price>350000</Price>
                        </Contract>
                    </Manager>
                    <Manager>
                        <Name>Tony Anderson</Name>
                        <Contract>
                            <Client>
                                <Name>E Corp.</Name>
                            </Client>
                            <Price>650000</Price>
                        </Contract>
                        <Contract>
                            <Client>
                                <Name>F &amp; Partners</Name>
                            </Client>
                            <Price>550000</Price>
                        </Contract>
                    </Manager>
                    <Manager>
                        <Name>July James</Name>
                        <Contract>
                            <Client>
                                <Name>G &amp; Co.</Name>
                            </Client>
                            <Price>350000</Price>
                        </Contract>
                        <Contract>
                            <Client>
                                <Name>H Group</Name>
                            </Client>
                            <Price>250000</Price>
                        </Contract>
                        <Contract>
                            <Client>
                                <Name>I &amp; Sons</Name>
                            </Client>
                            <Price>100000</Price>
                        </Contract>
                        <Contract>
                            <Client>
                                <Name>J Ent.</Name>
                            </Client>
                            <Price>100000</Price>
                        </Contract>
                    </Manager>
                </Managers>"""

        template_md = b"""<<foreach [in managers]>>**Manager:** <<[Name]>>
**Contracts:**
<<foreach [in Contract]>>- <<[Client.Name]>> ($<<[Price]>>)
<</foreach>>
<</foreach>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)

        data_source = aw.reporting.XmlDataSource(io.BytesIO(xml))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source, "managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.xml_data_source_complex.docx")


    def test_json_data_source(self) :

        json = b"""{
                        Name: "John Doe",
                        Age: 30,
                        Birth: "1989-04-01 4:00:00 pm",
                        Child: [ "Ann Doe", "Charles Doe" ]
                    }"""

        template_md = b"""**Name**: <<[Name]>>, **Age**: <<[Age]>>, **Date of Birth**:
<<[Birth]:"dd.MM.yyyy">>

**Children**:
<<foreach [in Child]>><<[Child_Text]>>
<</foreach>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)
        data_source = aw.reporting.JsonDataSource(io.BytesIO(json))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source)

        doc.save(docs_base.artifacts_dir + "ReportingEngine.json_data_source.docx")


    def test_json_data_source_seq(self) :

        json = b"""[
                    {
                        Name: "John Doe",
                        Age: 30,
                        Birth: "1989-04-01 4:00:00 pm"
                    },
                    {
                        Name: "Jane Doe",
                        Age: 27,
                        Birth: "1992-01-31 07:00:00 am"
                    },
                    {
                        Name: "John Smith",
                        Age: 51,
                        Birth: "1968-03-08 1:00:00 pm"
                    }
                ]"""

        template_md = b"""<<foreach [in persons]>>**Name:** <<[Name]>>, **Age:** <<[Age]>>, **Date of Birth:** <<[Birth]:"dd.MM.yyyy">>
<</foreach>>

**Average age:** <<[persons.Average(p => p.Age)]>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)
        data_source = aw.reporting.JsonDataSource(io.BytesIO(json))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source, "persons")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.json_data_source_seq.docx")

    def test_json_data_source_complex(self) :

        json = b"""[
                    {
                        Name: "John Smith",
                        Contract:
                        [
                            {
                                Client:
                                {
                                    Name: "A Company"
                                },
                                Price: 1200000
                            },
                            {
                                Client:
                                {
                                    Name: "B Ltd."
                                },
                                Price: 750000
                            },
                            {
                                Client:
                                {
                                    Name: "C & D"
                                },
                                Price: 350000
                            }
                        ]
                    },
                    {
                        Name: "Tony Anderson",
                        Contract:
                        [
                            {
                                Client:
                                {
                                    Name: "E Corp."
                                },
                                Price: 650000
                            },
                            {
                                Client:
                                {
                                    Name: "F & Partners"
                                },
                                Price: 550000
                            }
                        ]
                    },
                    {
                        Name: "July James",
                        Contract:
                        [
                            {
                                Client:
                                {
                                    Name: "G & Co."
                                },
                                Price: 350000
                            },
                            {
                                Client:
                                {
                                    Name: "H Group"
                                },
                                Price: 250000
                            },
                            {
                                Client:
                                {
                                    Name: "I & Sons"
                                },
                                Price: 100000
                            },
                            {
                                Client:
                                {
                                    Name: "J Ent."
                                },
                                Price: 100000
                            }
                        ]
                    }
                ]"""

        template_md = b"""<<foreach [in managers]>>**Manager:** <<[Name]>>
**Contracts:**
<<foreach [in Contract]>>- <<[Client.Name]>> ($<<[Price]>>)
<</foreach>>
<</foreach>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)

        data_source = aw.reporting.JsonDataSource(io.BytesIO(json))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source, "managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.json_data_source_complex.docx")


    def test_exact_date_time_parse_formats(self) :

        formats = [ "MM/dd/yyyy" ]
        options = aw.reporting.JsonDataLoadOptions()
        options.exact_date_time_parse_formats = formats

#=========================================================================================================================================================

    def test_csv_data_source_seq(self) :

        csv = b"""John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm"""

        template_md = b"""<<foreach [in persons]>>**Name:** <<[Column1]>>, **Age:** <<[Column2]>>, **Date of Birth:** <<[Column3]:"dd.MM.yyyy">>
<</foreach>>

**Average age:** <<[persons.Average(p => p.Column2)]>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)
        data_source = aw.reporting.CsvDataSource(io.BytesIO(csv))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source, "persons")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.csv_data_source_seq.docx")

    def test_csv_data_source_seq_column_names(self) :

        csv = b"""Name,Age,Birth
John Doe,30,1989-04-01 4:00:00 pm
Jane Doe,27,1992-01-31 07:00:00 am
John Smith,51,1968-03-08 1:00:00 pm"""

        template_md = b"""<<foreach [in persons]>>**Name:** <<[Name]>>, **Age:** <<[Age]>>, **Date of Birth:** <<[Birth]:"dd.MM.yyyy">>
<</foreach>>

**Average age:** <<[persons.Average(p => p.Age)]>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)
        options = aw.reporting.CsvDataLoadOptions(True)
        data_source = aw.reporting.CsvDataSource(io.BytesIO(csv), options)

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source, "persons")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.csv_data_source_seq_column_names.docx")

    def test_csv_data_source_complex(self) :

        csv = b"""[
                    {
                        Name: "John Smith",
                        Contract:
                        [
                            {
                                Client:
                                {
                                    Name: "A Company"
                                },
                                Price: 1200000
                            },
                            {
                                Client:
                                {
                                    Name: "B Ltd."
                                },
                                Price: 750000
                            },
                            {
                                Client:
                                {
                                    Name: "C & D"
                                },
                                Price: 350000
                            }
                        ]
                    },
                    {
                        Name: "Tony Anderson",
                        Contract:
                        [
                            {
                                Client:
                                {
                                    Name: "E Corp."
                                },
                                Price: 650000
                            },
                            {
                                Client:
                                {
                                    Name: "F & Partners"
                                },
                                Price: 550000
                            }
                        ]
                    },
                    {
                        Name: "July James",
                        Contract:
                        [
                            {
                                Client:
                                {
                                    Name: "G & Co."
                                },
                                Price: 350000
                            },
                            {
                                Client:
                                {
                                    Name: "H Group"
                                },
                                Price: 250000
                            },
                            {
                                Client:
                                {
                                    Name: "I & Sons"
                                },
                                Price: 100000
                            },
                            {
                                Client:
                                {
                                    Name: "J Ent."
                                },
                                Price: 100000
                            }
                        ]
                    }
                ]"""

        template_md = b"""<<foreach [in managers]>>**Manager:** <<[Name]>>
**Contracts:**
<<foreach [in Contract]>>- <<[Client.Name]>> ($<<[Price]>>)
<</foreach>>
<</foreach>>"""

        load_options = aw.loading.LoadOptions()
        load_options.load_format = aw.LoadFormat.MARKDOWN

        doc = aw.Document(io.BytesIO(template_md), load_options)

        data_source = aw.reporting.CsvDataSource(io.BytesIO(csv))

        engine = aw.reporting.ReportingEngine()
        engine.build_report(doc, data_source, "managers")

        doc.save(docs_base.artifacts_dir + "ReportingEngine.csv_data_source_complex.docx")


if __name__ == '__main__':
    unittest.main()