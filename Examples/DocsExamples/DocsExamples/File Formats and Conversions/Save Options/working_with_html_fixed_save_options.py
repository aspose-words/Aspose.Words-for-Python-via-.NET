import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithHtmlFixedSaveOptions(DocsExamplesBase):

    def test_use_font_from_target_machine(self):

        #ExStart:UseFontFromTargetMachine
        doc = aw.Document(MY_DIR + "Bullet points with alternative font.docx")

        save_options = aw.saving.HtmlFixedSaveOptions()
        save_options.use_target_machine_fonts = True

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlFixedSaveOptions.use_font_from_target_machine.html", save_options)
        #ExEnd:UseFontFromTargetMachine


    def test_write_all_css_rules_in_single_file(self):

        #ExStart:WriteAllCssRulesInSingleFile
        doc = aw.Document(MY_DIR + "Document.docx")

        # Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
        # All CSS rules are written into single file "styles.css.
        save_options = aw.saving.HtmlFixedSaveOptions()
        save_options.save_font_face_css_separately = False

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlFixedSaveOptions.write_all_css_rules_in_single_file.html", save_options)
        #ExEnd:WriteAllCssRulesInSingleFile


if __name__ == '__main__':
    unittest.main()
