import aspose.words as aw

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExMossRtf2Docx:

    @staticmethod
    def convert_rtf_to_docx(in_file_name: str, out_file_name: str):

        # Load an RTF file into Aspose.Words.
        doc = aw.Document(in_file_name)

        # Save the document in the OOXML format.
        doc.save(out_file_name, aw.SaveFormat.DOCX)
