import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithHtmlFixedSaveOptions(docs_base.DocsExamplesBase):

    def test_use_font_from_target_machine(self) :

        #ExStart:UseFontFromTargetMachine
        doc = aw.Document(docs_base.my_dir + "Bullet points with alternative font.docx")

        saveOptions = aw.saving.HtmlFixedSaveOptions()
        saveOptions.use_target_machine_fonts = True

        doc.save(docs_base.artifacts_dir + "WorkingWithHtmlFixedSaveOptions.use_font_from_target_machine.html", saveOptions)
        #ExEnd:UseFontFromTargetMachine


    def test_write_all_css_rules_in_single_file(self) :

        #ExStart:WriteAllCssRulesInSingleFile
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        # Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
        # All CSS rules are written into single file "styles.css.
        saveOptions = aw.saving.HtmlFixedSaveOptions()
        saveOptions.save_font_face_css_separately = False

        doc.save(docs_base.artifacts_dir + "WorkingWithHtmlFixedSaveOptions.write_all_css_rules_in_single_file.html", saveOptions)
        #ExEnd:WriteAllCssRulesInSingleFile


if __name__ == '__main__':
    unittest.main()