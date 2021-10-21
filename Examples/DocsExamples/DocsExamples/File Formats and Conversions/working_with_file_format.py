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
        file_list =  (file for file in os.listdir(docs_base.my_dir)
           if (os.path.isfile(os.path.join(docs_base.my_dir, file)) and not file.endswith("Corrupted document.docx")))

        #ExEnd:GetListOfFilesInFolder
        for file_name in file_list:

            name_only = file_name
            file_name = os.path.join(docs_base.my_dir, name_only)

            print(name_only)
            #ExStart:DetectFileFormat
            info = aw.FileFormatUtil.detect_file_format(file_name)

            load_format = info.load_format
            # Display the document type
            if load_format == aw.LoadFormat.DOC:
                print("\tMicrosoft Word 97-2003 document.")
            elif load_format == aw.LoadFormat.DOT:
                print("\tMicrosoft Word 97-2003 template.")
            elif load_format ==  aw.LoadFormat.DOCX:
                print("\tOffice Open XML WordprocessingML Macro-Free Document.")
            elif load_format == aw.LoadFormat.DOCM:
                print("\tOffice Open XML WordprocessingML Macro-Enabled Document.")
            elif load_format == aw.LoadFormat.DOTX:
                print("\tOffice Open XML WordprocessingML Macro-Free Template.")
            elif load_format == aw.LoadFormat.DOTM:
                print("\tOffice Open XML WordprocessingML Macro-Enabled Template.")
            elif load_format == aw.LoadFormat.FLAT_OPC:
                print("\tFlat OPC document.")
            elif load_format == aw.LoadFormat.RTF:
                print("\tRTF format.")
            elif load_format == aw.LoadFormat.WORD_ML:
                print("\tMicrosoft Word 2003 WordprocessingML format.")
            elif load_format == aw.LoadFormat.HTML:
                print("\tHTML format.")
            elif load_format == aw.LoadFormat.MHTML:
                print("\tMHTML (Web archive) format.")
            elif load_format == aw.LoadFormat.ODT:
                print("\tOpenDocument Text.")
            elif load_format == aw.LoadFormat.OTT:
                print("\tOpenDocument Text Template.")
            elif load_format == aw.LoadFormat.DOC_PRE_WORD60:
                print("\tMS Word 6 or Word 95 format.")
            elif load_format == aw.LoadFormat.UNKNOWN:
                print("\tUnknown format.")

            #ExEnd:DetectFileFormat

            if info.is_encrypted:

                print("\tAn encrypted document.")
                shutil.copyfile(file_name, os.path.join(encrypted_dir, name_only))

            else:

                if load_format == aw.LoadFormat.DOC_PRE_WORD60:
                    shutil.copyfile(file_name, os.path.join(pre97_dir, name_only))
                elif load_format == aw.LoadFormat.UNKNOWN:
                    shutil.copyfile(file_name, os.path.join(unknown_dir, name_only))
                else:
                    shutil.copyfile(file_name, os.path.join(supported_dir, name_only))

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