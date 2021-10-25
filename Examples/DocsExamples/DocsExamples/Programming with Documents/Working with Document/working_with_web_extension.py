import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithWebExtension(DocsExamplesBase):

    def test_using_web_extension_task_panes(self):

        #ExStart:UsingWebExtensionTaskPanes
        doc = aw.Document()

        task_pane = aw.webextensions.TaskPane()
        doc.web_extension_task_panes.add(task_pane)

        task_pane.dock_state = aw.webextensions.TaskPaneDockState.RIGHT
        task_pane.is_visible = True
        task_pane.width = 300

        task_pane.web_extension.reference.id = "wa102923726"
        task_pane.web_extension.reference.version = "1.0.0.0"
        task_pane.web_extension.reference.store_type = aw.webextensions.WebExtensionStoreType.OMEX
        task_pane.web_extension.reference.store = "th-TH"
        task_pane.web_extension.properties.add(aw.webextensions.WebExtensionProperty("mailchimpCampaign", "mailchimpCampaign"))
        task_pane.web_extension.bindings.add(aw.webextensions.WebExtensionBinding("UnnamedBinding_0_1506535429545",
            aw.webextensions.WebExtensionBindingType.TEXT, "194740422"))

        doc.save(ARTIFACTS_DIR + "WorkingWithWebExtension.using_web_extension_task_panes.docx")
        #ExEnd:UsingWebExtensionTaskPanes

        #ExStart:GetListOfAddins
        doc = aw.Document(ARTIFACTS_DIR + "WorkingWithWebExtension.using_web_extension_task_panes.docx")

        print("Task panes sources:\n")

        for task_pane_info in doc.web_extension_task_panes:

            reference = task_pane_info.web_extension.reference
            print(f'Provider: "{reference.store}", version: "{reference.version}", catalog identifier: "{reference.id}"')

        #ExEnd:GetListOfAddins


if __name__ == '__main__':
    unittest.main()
