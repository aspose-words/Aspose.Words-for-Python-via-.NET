import unittest
import os
import shutil
import sys

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithFileFormat(DocsExamplesBase):

    def test_detect_file_format(self):

        #ExStart:CheckFormatCompatibility
        supported_dir = ARTIFACTS_DIR + "Supported"
        unknown_dir = ARTIFACTS_DIR + "Unknown"
        encrypted_dir = ARTIFACTS_DIR + "Encrypted"
        pre97_dir = ARTIFACTS_DIR + "Pre97"

        # Create the directories if they do not already exist.
        for dirname in (supported_dir, unknown_dir, encrypted_dir, pre97_dir):
            os.makedirs(dirname, exist_ok=True)

        #ExStart:GetListOfFilesInFolder
        file_list = [file for file in os.listdir(MY_DIR)
                     if os.path.isfile(os.path.join(MY_DIR, file) and not file.endswith("Corrupted document.docx"))]
        #ExEnd:GetListOfFilesInFolder

        for name_only in file_list:
            file_name = os.path.join(MY_DIR, name_only)
            print(name_only)

            #ExStart:DetectFileFormat
            info = aw.FileFormatUtil.detect_file_format(file_name)

            load_format = info.load_format
            # Display the document type
            if load_format == aw.LoadFormat.DOC:
                print("\tMicrosoft Word 97-2003 document.")
            elif load_format == aw.LoadFormat.DOT:
                print("\tMicrosoft Word 97-2003 template.")
            elif load_format == aw.LoadFormat.DOCX:
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
            elif load_format == aw.LoadFormat.DOC_PRE_WORD60:
                shutil.copyfile(file_name, os.path.join(pre97_dir, name_only))
            elif load_format == aw.LoadFormat.UNKNOWN:
                shutil.copyfile(file_name, os.path.join(unknown_dir, name_only))
            else:
                shutil.copyfile(file_name, os.path.join(supported_dir, name_only))

        #ExEnd:CheckFormatCompatibility

    def test_detect_document_signatures(self):

        #ExStart:DetectDocumentSignatures
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + "Digitally signed.docx")

        if info.has_digital_signature:
            print("Document has digital signatures, they will be lost if you open/save this document with Aspose.words.")

        #ExEnd:DetectDocumentSignatures

    def test_verify_encrypted_document(self):

        #ExStart:VerifyEncryptedDocument
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + "Encrypted.docx")
        print(info.is_encrypted)
        #ExEnd:VerifyEncryptedDocument
