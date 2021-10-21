import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithMarkdownSaveOptions(docs_base.DocsExamplesBase):

    def test_export_into_markdown_with_table_content_alignment(self) :

        #ExStart:ExportIntoMarkdownWithTableContentAlignment
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_cell()
        builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
        builder.write("Cell1")
        builder.insert_cell()
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.write("Cell2")

        # Makes all paragraphs inside the table to be aligned.
        save_options = aw.saving.MarkdownSaveOptions()

        save_options.table_content_alignment = aw.saving.TableContentAlignment.LEFT

        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdownSaveOptions.left_table_content_alignment.md", save_options)

        save_options.table_content_alignment = aw.saving.TableContentAlignment.RIGHT
        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdownSaveOptions.right_table_content_alignment.md", save_options)

        save_options.table_content_alignment = aw.saving.TableContentAlignment.CENTER
        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdownSaveOptions.center_table_content_alignment.md", save_options)

        # The alignment in this case will be taken from the first paragraph in corresponding table column.
        save_options.table_content_alignment = aw.saving.TableContentAlignment.AUTO
        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdownSaveOptions.auto_table_content_alignment.md", save_options)
        #ExEnd:ExportIntoMarkdownWithTableContentAlignment


    def test_set_images_folder(self) :

        #ExStart:SetImagesFolder
        doc = aw.Document(docs_base.my_dir + "Image bullet points.docx")

        save_options = aw.saving.MarkdownSaveOptions()
        save_options.images_folder = docs_base.artifacts_dir + "Images"

        doc.save(docs_base.artifacts_dir + "WorkingWithMarkdownSaveOptions.set_images_folder.md", save_options)
        #ExEnd:SetImagesFolder




if __name__ == '__main__':
    unittest.main()