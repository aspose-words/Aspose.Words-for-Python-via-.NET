import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
import aspose.pydrawing as drawing

class WorkingWithList(docs_base.DocsExamplesBase):

    def test_restart_list_at_each_section(self) :

        #ExStart:RestartListAtEachSection
        doc = aw.Document()

        doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)

        list = doc.lists[0]
        list.is_restart_at_each_section = True

        builder = aw.DocumentBuilder(doc)
        builder.list_format.list = list

        for i in range(1, 45) :

            builder.writeln(f"List Item {i}")

            if i == 15 :
                builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)


        # IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.ecma_376.
        options = aw.saving.OoxmlSaveOptions()
        options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL

        doc.save(docs_base.artifacts_dir + "WorkingWithList.restart_list_at_each_section.docx", options)
        #ExEnd:RestartListAtEachSection


    def test_specify_list_level(self) :

        #ExStart:SpecifyListLevel
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a numbered list based on one of the Microsoft Word list templates
        # and apply it to the document builder's current paragraph.
        builder.list_format.list = doc.lists.add(aw.lists.ListTemplate.NUMBER_ARABIC_DOT)

        # There are nine levels in this list, let's try them all.
        for i in range(0, 9) :

            builder.list_format.list_level_number = i
            builder.writeln(f"Level {i}")


        # Create a bulleted list based on one of the Microsoft Word list templates
        # and apply it to the document builder's current paragraph.
        builder.list_format.list = doc.lists.add(aw.lists.ListTemplate.BULLET_DIAMONDS)

        for i in range(0, 9) :

            builder.list_format.list_level_number = i
            builder.writeln(f"Level {i}")


        # This is a way to stop list formatting.
        builder.list_format.list = None

        builder.document.save(docs_base.artifacts_dir + "WorkingWithList.specify_list_level.docx")
        #ExEnd:SpecifyListLevel


    def test_restart_list_number(self) :

        #ExStart:RestartListNumber
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a list based on a template.
        list1 = doc.lists.add(aw.lists.ListTemplate.NUMBER_ARABIC_PARENTHESIS)
        list1.list_levels[0].font.color = drawing.Color.red
        list1.list_levels[0].alignment = aw.lists.ListLevelAlignment.RIGHT

        builder.writeln("List 1 starts below:")
        builder.list_format.list = list1
        builder.writeln("Item 1")
        builder.writeln("Item 2")
        builder.list_format.remove_numbers()

        # To reuse the first list, we need to restart numbering by creating a copy of the original list formatting.
        list2 = doc.lists.add_copy(list1)

        # We can modify the new list in any way, including setting a new start number.
        list2.list_levels[0].start_at = 10

        builder.writeln("List 2 starts below:")
        builder.list_format.list = list2
        builder.writeln("Item 1")
        builder.writeln("Item 2")
        builder.list_format.remove_numbers()

        builder.document.save(docs_base.artifacts_dir + "WorkingWithList.restart_list_number.docx")
        #ExEnd:RestartListNumber



if __name__ == '__main__':
    unittest.main()
