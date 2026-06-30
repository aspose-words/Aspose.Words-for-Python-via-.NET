import io

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithMarkdownSaveOptions(DocsExamplesBase):

    def test_markdown_table_content_alignment(self):

        #ExStart:MarkdownTableContentAlignment
        #GistId:19de942ef8827201c1dca99f76c59133
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
        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdownSaveOptions.left_table_content_alignment.md", save_options)

        save_options.table_content_alignment = aw.saving.TableContentAlignment.RIGHT
        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdownSaveOptions.right_table_content_alignment.md", save_options)

        save_options.table_content_alignment = aw.saving.TableContentAlignment.CENTER
        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdownSaveOptions.center_table_content_alignment.md", save_options)

        # The alignment in this case will be taken from the first paragraph in corresponding table column.
        save_options.table_content_alignment = aw.saving.TableContentAlignment.AUTO
        doc.save(ARTIFACTS_DIR + "WorkingWithMarkdownSaveOptions.auto_table_content_alignment.md", save_options)
        #ExEnd:MarkdownTableContentAlignment

    def test_images_folder(self):

        #ExStart:ImagesFolder
        #GistId:51b4cb9c451832f23527892e19c7bca6
        doc = aw.Document(MY_DIR + "Image bullet points.docx")

        save_options = aw.saving.MarkdownSaveOptions()
        save_options.images_folder = ARTIFACTS_DIR + "Images"

        with io.BytesIO() as stream:
            doc.save(stream, save_options)
        #ExEnd:ImagesFolder
