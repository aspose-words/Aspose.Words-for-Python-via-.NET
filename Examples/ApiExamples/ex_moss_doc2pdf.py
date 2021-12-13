"""DOC2PDF document converter for SharePoint.
Uses Aspose.Words to perform the conversion."""

import sys
import logging
from typing import List, NamedTuple

import aspose.words as aw

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class Options(NamedTuple):
    in_file_name: str
    out_file_name: str


def convert_doc2pdf(in_file_name: str, out_file_name: str):
    # You can load not only DOC here, but any format supported by
    # Aspose.Words: DOC, DOCX, RTF, WordML, HTML, MHTML, ODT etc.
    doc = aw.Document(in_file_name)

    doc.save(out_file_name, aw.saving.PdfSaveOptions())


def main():

    # Although SharePoint passes "-log <filename>" to us and we are
    # supposed to log there, we will use our hardcoded path to the log file for the sake of simplicity.
    #
    # Make sure there are permissions to write into this folder.
    # The document converter will be called under the document
    # conversion account (not sure what name), so for testing purposes,
    # I would give the Users group write permissions into this folder.
    logging.basicConfig(filename=r"C:\Aspose2Pdf\log.txt")
    
    logger = logging.getLogger("doc2pdf")

    logger.info("Started")
    logger.info('Command line:', ' '.join(sys.argv[1:]))

    options = parse_command_line(sys.argv[1:])

    # Uncomment the code below when you have purchased a license for Aspose.Words.
    #
    # You need to deploy the license in the same folder as your
    # executable, alternatively you can add the license file as an
    # embedded resource to your project.
    #
    # Set license for Aspose.Words.
    # words_license = aw.License()
    # words_license.set_license("Aspose.Total.lic");

    convert_doc2pdf(options.in_file_name, options.out_file_name)


def parse_command_line(args: List[str]) -> Options:

    in_file_name = None
    out_file_name = None

    i = 0
    while i < len(args):
        token = args[i].lower()
        if token == "-in":
            i += 1
            in_file_name = args[i]

        elif token == "-out":
            i += 1
            out_file_name = args[i]

        elif token == "-config":
            # Skip the name of the config file and do nothing.
            i += 1

        elif token == "-log":
            # Skip the name of the log file and do nothing.
            i += 1
        
        else:
            raise Exception("Unknown command line argument: " + token)

        i += 1

    return Options(in_file_name, out_file_name)

    
if __name__ == '__main__':
    main()
