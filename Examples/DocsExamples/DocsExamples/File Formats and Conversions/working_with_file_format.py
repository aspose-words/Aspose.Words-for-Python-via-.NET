import unittest
import os
import shutil
import sys

base_dir = os.path.abspath(os.curdir) + "/"
print("This is base_dir:" + base_dir)
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithFileFormat(docs_base.DocsExamplesBase):
    
    def test_detect_file_format(self) :
        
        #ExStart:CheckFormatCompatibility
        supported_dir = docs_base.artifacts_dir + "Supported"
        unknown_dir = docs_base.artifacts_dir + "Unknown"
        encrypted_dir = docs_base.artifacts_dir + "Encrypted"
        pre97_dir = docs_base.artifacts_dir + "Pre97"

        # Create the directories if they do not already exist.
        if not os.path.exists(supported_dir):
            os.makedirs(supported_dir)
        if not os.path.exists(unknown_dir):
            os.makedirs(unknown_dir)
        if not os.path.exists(encrypted_dir):
            os.makedirs(encrypted_dir)
        if not os.path.exists(pre97_dir):
            os.makedirs(pre97_dir)

        #ExStart:GetListOfFilesInFolder
        fileList =  (file for file in os.listdir(docs_base.my_dir)
           if (os.path.isfile(os.path.join(docs_base.my_dir, file)) and not file.endswith("Corrupted document.docx")))

        #ExEnd:GetListOfFilesInFolder
        for fileName in fileList:

            name_only = fileName
            fileName = os.path.join(docs_base.my_dir, name_only)

            print(name_only)
            #ExStart:DetectFileFormat
            info = aw.FileFormatUtil.detect_file_format(fileName)

            lf = info.load_format
            # Display the document type
            if lf == aw.LoadFormat.DOC:
                print("\tMicrosoft Word 97-2003 document.")
            elif lf == aw.LoadFormat.DOT:
                print("\tMicrosoft Word 97-2003 template.")
            elif lf ==  aw.LoadFormat.DOCX:
                print("\tOffice Open XML WordprocessingML Macro-Free Document.")
            elif lf == aw.LoadFormat.DOCM:
                print("\tOffice Open XML WordprocessingML Macro-Enabled Document.")
            elif lf == aw.LoadFormat.DOTX:
                print("\tOffice Open XML WordprocessingML Macro-Free Template.")
            elif lf == aw.LoadFormat.DOTM:
                print("\tOffice Open XML WordprocessingML Macro-Enabled Template.")
            elif lf == aw.LoadFormat.FLAT_OPC:
                print("\tFlat OPC document.")
            elif lf == aw.LoadFormat.RTF:
                print("\tRTF format.")
            elif lf == aw.LoadFormat.WORD_ML:
                print("\tMicrosoft Word 2003 WordprocessingML format.")
            elif lf == aw.LoadFormat.HTML:
                print("\tHTML format.")
            elif lf == aw.LoadFormat.MHTML:
                print("\tMHTML (Web archive) format.")
            elif lf == aw.LoadFormat.ODT:
                print("\tOpenDocument Text.")
            elif lf == aw.LoadFormat.OTT:
                print("\tOpenDocument Text Template.")
            elif lf == aw.LoadFormat.DOC_PRE_WORD60:
                print("\tMS Word 6 or Word 95 format.")
            elif lf == aw.LoadFormat.UNKNOWN:
                print("\tUnknown format.")
                
            #ExEnd:DetectFileFormat

            if info.is_encrypted:
                
                print("\tAn encrypted document.")
                shutil.copyfile(fileName, os.path.join(encrypted_dir, name_only))
                
            else:
                
                if lf == aw.LoadFormat.DOC_PRE_WORD60:
                    shutil.copyfile(fileName, os.path.join(pre97_dir, name_only))
                elif lf == aw.LoadFormat.UNKNOWN:
                    shutil.copyfile(fileName, os.path.join(unknown_dir, name_only))
                else: 
                    shutil.copyfile(fileName, os.path.join(supported_dir, name_only))
            
        #ExEnd:CheckFormatCompatibility
        

    def test_detect_document_signatures(self) :
        
        #ExStart:DetectDocumentSignatures
        info = aw.FileFormatUtil.detect_file_format(docs_base.my_dir + "Digitally signed.docx")

        if info.has_digital_signature:
            print("Document has digital signatures, they will be lost if you open/save this document with Aspose.words.")
            
        #ExEnd:DetectDocumentSignatures            
        

    def test_verify_encrypted_document(self) :
        
        #ExStart:VerifyEncryptedDocument
        info = aw.FileFormatUtil.detect_file_format(docs_base.my_dir + "Encrypted.docx")
        print(info.is_encrypted)
        #ExEnd:VerifyEncryptedDocument
        
    
if __name__ == '__main__':
    unittest.main()    