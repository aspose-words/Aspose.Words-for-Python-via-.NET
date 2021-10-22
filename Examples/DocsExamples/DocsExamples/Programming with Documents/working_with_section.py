import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithSection(docs_base.DocsExamplesBase):

    def test_add_section(self) :

        #ExStart:AddSection
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        builder.writeln("Hello2")

        section_to_add = aw.Section(doc)
        doc.sections.add(section_to_add)
        #ExEnd:AddSection


    def test_delete_section(self) :

        #ExStart:DeleteSection
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello2")
        doc.append_child(aw.Section(doc))

        doc.sections.remove_at(0)
        #ExEnd:DeleteSection


    def test_delete_all_sections(self) :

        #ExStart:DeleteAllSections
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello2")
        doc.append_child(aw.Section(doc))

        doc.sections.clear()
        #ExEnd:DeleteAllSections


    def test_append_section_content(self) :

        #ExStart:AppendSectionContent
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello22")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello3")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello45")

        # This is the section that we will append and prepend to.
        section = doc.sections[2]

        # This copies the content of the 1st section and inserts it at the beginning of the specified section.
        section_to_prepend = doc.sections[0]
        section.prepend_content(section_to_prepend)

        # This copies the content of the 2nd section and inserts it at the end of the specified section.
        section_to_append = doc.sections[1]
        section.append_content(section_to_append)
        #ExEnd:AppendSectionContent


    def test_clone_section(self) :

        #ExStart:CloneSection
        doc = aw.Document(docs_base.my_dir + "Document.docx")
        clone_section = doc.sections[0].clone()
        #ExEnd:CloneSection


    def test_copy_section(self) :

        #ExStart:CopySection
        src_doc = aw.Document(docs_base.my_dir + "Document.docx")
        dst_doc = aw.Document()

        source_section = src_doc.sections[0]
        new_section = dst_doc.import_node(source_section, True).as_section()
        dst_doc.sections.add(new_section)

        dst_doc.save(docs_base.artifacts_dir + "WorkingWithSection.copy_section.docx")
        #ExEnd:CopySection


    def test_delete_header_footer_content(self) :

        #ExStart:DeleteHeaderFooterContent
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        section = doc.sections[0]
        section.clear_headers_footers()
        #ExEnd:DeleteHeaderFooterContent


    def test_delete_section_content(self) :

        #ExStart:DeleteSectionContent
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        section = doc.sections[0]
        section.clear_content()
        #ExEnd:DeleteSectionContent


    def test_modify_page_setup_in_all_sections(self) :

        #ExStart:ModifyPageSetupInAllSections
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello1")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello22")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello3")
        doc.append_child(aw.Section(doc))
        builder.writeln("Hello45")

        # It is important to understand that a document can contain many sections,
        # and each section has its page setup. In this case, we want to modify them all.
        for child in doc :
            child.as_section().page_setup.paper_size = aw.PaperSize.LETTER

        doc.save(docs_base.artifacts_dir + "WorkingWithSection.modify_page_setup_in_all_sections.doc")
        #ExEnd:ModifyPageSetupInAllSections


    def test_sections_access_by_index(self) :

        #ExStart:SectionsAccessByIndex
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        section = doc.sections[0]
        section.page_setup.left_margin = 90 # 3.17 cm
        section.page_setup.right_margin = 90 # 3.17 cm
        section.page_setup.top_margin = 72 # 2.54 cm
        section.page_setup.bottom_margin = 72 # 2.54 cm
        section.page_setup.header_distance = 35.4 # 1.25 cm
        section.page_setup.footer_distance = 35.4 # 1.25 cm
        section.page_setup.text_columns.spacing = 35.4 # 1.25 cm
        #ExEnd:SectionsAccessByIndex


if __name__ == '__main__':
    unittest.main()
