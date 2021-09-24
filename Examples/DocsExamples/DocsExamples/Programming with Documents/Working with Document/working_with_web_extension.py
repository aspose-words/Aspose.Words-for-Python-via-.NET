import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithWebExtension(docs_base.DocsExamplesBase):
    
        def test_using_web_extension_task_panes(self) :
        
            #ExStart:UsingWebExtensionTaskPanes
            doc = aw.Document()

            taskPane = aw.webextensions.TaskPane()
            doc.web_extension_task_panes.add(taskPane)

            taskPane.dock_state = aw.webextensions.TaskPaneDockState.RIGHT
            taskPane.is_visible = True
            taskPane.width = 300

            taskPane.web_extension.reference.id = "wa102923726"
            taskPane.web_extension.reference.version = "1.0.0.0"
            taskPane.web_extension.reference.store_type = aw.webextensions.WebExtensionStoreType.OMEX
            taskPane.web_extension.reference.store = "th-TH"
            taskPane.web_extension.properties.add(aw.webextensions.WebExtensionProperty("mailchimpCampaign", "mailchimpCampaign"))
            taskPane.web_extension.bindings.add(aw.webextensions.WebExtensionBinding("UnnamedBinding_0_1506535429545",
                aw.webextensions.WebExtensionBindingType.TEXT, "194740422"))

            doc.save(docs_base.artifacts_dir + "WorkingWithWebExtension.using_web_extension_task_panes.docx")
            #ExEnd:UsingWebExtensionTaskPanes
            
            #ExStart:GetListOfAddins
            doc = aw.Document(docs_base.artifacts_dir + "WorkingWithWebExtension.using_web_extension_task_panes.docx")
            
            print("Task panes sources:\n")

            for taskPaneInfo in doc.web_extension_task_panes :
            
                reference = taskPaneInfo.web_extension.reference
                print(f"Provider: \"{reference.store}\", version: \"{reference.version}\", catalog identifier: \"{reference.id}\"")
            
            #ExEnd:GetListOfAddins
        
    

if __name__ == '__main__':
        unittest.main()