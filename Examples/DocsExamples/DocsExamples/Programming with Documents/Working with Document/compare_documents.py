import unittest
import os
import sys
from datetime import date, datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class CompareDocument(docs_base.DocsExamplesBase):
    
        def test_compare_for_equal(self) :
        
            #ExStart:CompareForEqual
            docA = aw.Document(docs_base.my_dir + "Document.docx")
            docB = docA.clone().as_document()
            
            # DocA now contains changes as revisions.
            docA.compare(docB, "user", datetime.today())

            print("Documents are equal" if (docA.revisions.count == 0) else "Documents are not equal")
            #ExEnd:CompareForEqual                     
        

        def test_compare_options(self) :
        
            #ExStart:CompareOptions
            docA = aw.Document(docs_base.my_dir + "Document.docx")
            docB = docA.clone()

            options = aw.comparing.CompareOptions()
            
            options.ignore_formatting = True
            options.ignore_headers_and_footers = True
            options.ignore_case_changes = True
            options.ignore_tables = True
            options.ignore_fields = True
            options.ignore_comments = True
            options.ignore_textboxes = True
            options.ignore_footnotes = True
            

            docA.compare(docB, "user", datetime.today(), options)

            print("Documents are equal" if (docA.revisions.count == 0) else "Documents are not equal")
            #ExEnd:CompareOptions                     
        

        def test_comparison_target(self) :
        
            #ExStart:ComparisonTarget
            docA = aw.Document(docs_base.my_dir + "Document.docx")
            docB = docA.clone()

            # Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box.
            options = aw.comparing.CompareOptions()
            options.ignore_formatting = True
            options.target = aw.comparing.ComparisonTargetType.NEW 

            docA.compare(docB, "user", datetime.today(), options)
            #ExEnd:ComparisonTarget
        

        def test_comparison_granularity(self) :
        
            #ExStart:ComparisonGranularity
            builderA = aw.DocumentBuilder(aw.Document())
            builderB = aw.DocumentBuilder(aw.Document())

            builderA.writeln("This is A simple word")
            builderB.writeln("This is B simple words")

            compareOptions = aw.comparing.CompareOptions()
            compareOptions.granularity = aw.comparing.Granularity.CHAR_LEVEL 

            builderA.document.compare(builderB.document, "author", datetime.today(), compareOptions)
            #ExEnd:ComparisonGranularity      
        
    

if __name__ == '__main__':
        unittest.main()