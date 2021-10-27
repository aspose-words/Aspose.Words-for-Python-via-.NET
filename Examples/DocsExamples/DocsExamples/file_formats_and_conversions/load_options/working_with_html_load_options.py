import io

from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithHtmlLoadOptions(DocsExamplesBase):

    def test_preferred_control_type(self):

        #ExStart:LoadHtmlElementsWithPreferredControlType
        html = """
            <html>
                <select name='ComboBox' size='1'>
                    <option value='val1'>item1</option>
                    <option value='val2'></option>
                </select>
            </html>
        """

        print(html)

        load_options = aw.loading.HtmlLoadOptions()
        load_options.preferred_control_type = aw.loading.HtmlControlType.STRUCTURED_DOCUMENT_TAG

        doc = aw.Document(io.BytesIO(html.encode("utf-8")), load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlLoadOptions.preferred_control_type.docx", aw.SaveFormat.DOCX)
        #ExEnd:LoadHtmlElementsWithPreferredControlType
