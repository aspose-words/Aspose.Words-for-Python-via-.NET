import io
import unittest
import sys
import aspose.words as aw
import aspose.pydrawing as drawing
import aspose.email as ae
from aspose.email.clients.smtp import SmtpClient
from aspose.words.replacing import FindReplaceOptions
from aspose.words.saving import XlsxSaveOptions, CompressionLevel

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, IMAGES_DIR

class BaseConversions(DocsExamplesBase):

    def test_doc_to_docx(self):

        #ExStart:LoadAndSave
        #GistId:1ea924a385a086092413b7fc5ea0f5e9
        #ExStart:OpenDocument
        doc = aw.Document(MY_DIR + "Document.doc")
        #ExEnd:OpenDocument

        doc.save(ARTIFACTS_DIR + "BaseConversions.doc_to_docx.docx")
        #ExEnd:LoadAndSave

    def test_docx_to_rtf(self):

        #ExStart:LoadAndSaveToStream
        #GistId:1ea924a385a086092413b7fc5ea0f5e9
        #ExStart:OpenFromStream
        #GistId:59e45f5041ff6b356c5165164c019a76
        # Read only access is enough for Aspose.Words to load a document.
        stream = io.FileIO(MY_DIR + "Document.docx")

        doc = aw.Document(stream)
        # You can close the stream now, it is no longer needed because the document is in memory.
        stream.close()
        #ExEnd:OpenFromStream

        # ... do something with the document.

        # Convert the document to a different format and save to stream.
        dst_stream = io.BytesIO()
        doc.save(dst_stream, aw.SaveFormat.RTF)
        #ExEnd:LoadAndSaveToStream

        with open(ARTIFACTS_DIR + "BaseConversions.docx_to_rtf.rtf", "wb") as output:
            output.write(dst_stream.getbuffer())

    def test_docx_to_pdf(self):

        #ExStart:DocxToPdf
        #GistId:36a49a29062268dc5e6d3134163f8d99
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_pdf.pdf")
        #ExEnd:DocxToPdf

    def test_docx_to_byte(self):

        #ExStart:DocxToByte
        #GistId:9278593292345acbef67679a2afb4286
        doc = aw.Document(MY_DIR + "Document.docx")

        out_stream = io.BytesIO()
        doc.save(out_stream, aw.SaveFormat.DOCX)

        doc_bytes = out_stream.getbuffer()
        in_stream = io.BytesIO(doc_bytes)

        doc_from_bytes = aw.Document(in_stream)
        #ExEnd:DocxToByte

    def test_docx_to_epub(self):

        #ExStart:DocxToEpub
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_epub.epub")
        #ExEnd:DocxToEpub

    def test_docx_to_html(self):

        #ExStart:DocxToHtml
        #GistId:c5d01a71f65e540e5e77650b846b24cc
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_html.html")
        #ExEnd:DocxToHtml

    @unittest.skip("This test should be run manually with a real SMTP server")
    def test_docx_to_mhtml_and_sending_email(self):

        #ExStart:DocxToMhtml
        #GistId:16caeb7d9781fa093a70df33f232dbaa
        doc = aw.Document(MY_DIR + "Document.docx")

        stream = io.BytesIO()
        doc.save(stream, aw.SaveFormat.MHTML)

        # Rewind the stream to the beginning so Aspose.Email can read it.
        stream.seek(0)

        # Create an Aspose.Email MIME email message from the stream.
        message = ae.MailMessage.load(stream, ae.MhtmlLoadOptions())
        message.from_address = ae.MailAddress("your_from@email.com")
        message.to.append(ae.MailAddress("your_to@email.com"))
        message.subject = "Aspose.Words + Aspose.Email MHTML Test Message"

        # Send the message using Aspose.Email.
        client = SmtpClient()
        client.host = "your_smtp.com"
        client.send(message)
        #ExEnd:DocxToMhtml

    def test_docx_to_markdown(self):

        #ExStart:DocxToMarkdown
        #GistId:461290170d82b0922d265fa7bc854942
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Some text!")

        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_markdown.md")
        #ExEnd:DocxToMarkdown

    def test_docx_to_txt(self):

        #ExStart:DocxToTxt
        #GistId:399801c9a5e656ed05aa2d7ac5ebc41e
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_txt.txt")
        #ExEnd:DocxToTxt

    def test_docx_to_xlsx(self):

        #ExStart:DocxToXlsx
        #GistId:1f614ac0f43c0f1392c52366850cf211
        doc = aw.Document(MY_DIR + "Document.docx")
        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_xlsx.xlsx")
        #ExEnd:DocxToXlsx

    def test_txt_to_docx(self):

        #ExStart:TxtToDocx
        # The encoding of the text file is automatically detected.
        doc = aw.Document(MY_DIR + "English text.txt")

        doc.save(ARTIFACTS_DIR + "BaseConversions.txt_to_docx.docx")
        #ExEnd:TxtToDocx

    def test_pdf_to_jpeg(self):

        #ExStart:PdfToJpeg
        #GistId:f9e5cde75221f622f636297c5fcc7297
        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        doc.save(ARTIFACTS_DIR + "BaseConversions.pdf_to_jpeg.jpeg")
        #ExEnd:PdfToJpeg

    def test_pdf_to_docx(self):

        #ExStart:PdfToDocx
        #GistId:1cd02caea10d62b6238a3177a70dd81d
        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        doc.save(ARTIFACTS_DIR + "BaseConversions.pdf_to_docx.docx")
        #ExEnd:PdfToDocx

    def test_pdf_to_xlsx(self):
        #ExStart:PdfToXlsx
        #GistId:b2e1027992a4ccbf53b6a983a808ba20
        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        doc.save(ARTIFACTS_DIR + "BaseConversions.pdf_to_xlsx.xlsx")
        #ExEnd:PdfToXlsx

    def test_find_replace_xlsx(self):

        #ExStart:FindReplaceXlsx
        #GistId:b2e1027992a4ccbf53b6a983a808ba20
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Ruby bought a ruby necklace.")

        # We can use a "FindReplaceOptions" object to modify the find - and -replace process.

        options = FindReplaceOptions()

        # Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
        # Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
        options.match_case = True

        doc.range.replace("Ruby", "Jade", options)

        doc.save(ARTIFACTS_DIR + "BaseConversions.find_replace_xlsx.xlsx")
        #ExEnd:FindReplaceXlsx

    def test_compress_xlsx(self):

        #ExStart:CompressXlsx
        #GistId:b2e1027992a4ccbf53b6a983a808ba20
        doc = aw.Document(MY_DIR + "Document.docx")
        saveOptions = XlsxSaveOptions()
        saveOptions.compression_level = CompressionLevel.MAXIMUM

        doc.save(ARTIFACTS_DIR + "BaseConversions.compress_xlsx.xlsx", saveOptions)
        #ExEnd:CompressXlsx
    
    @unittest.skipUnless(sys.platform.startswith("win"), "requires Windows")
    def test_images_to_pdf(self):

        #ExStart:ImageToPdf
        #GistId:36a49a29062268dc5e6d3134163f8d99
        self.convert_image_to_pdf(IMAGES_DIR + "Logo.jpg",
                                  ARTIFACTS_DIR + "BaseConversions.JpgToPdf.pdf")
        self.convert_image_to_pdf(IMAGES_DIR + "Transparent background logo.png",
                                  ARTIFACTS_DIR + "BaseConversions.PngToPdf.pdf")
        self.convert_image_to_pdf(IMAGES_DIR + "Windows MetaFile.wmf",
                                  ARTIFACTS_DIR + "BaseConversions.WmfToPdf.pdf")
        self.convert_image_to_pdf(IMAGES_DIR + "Tagged Image File Format.tiff",
                                  ARTIFACTS_DIR + "BaseConversions.TiffToPdf.pdf")
        self.convert_image_to_pdf(IMAGES_DIR + "Graphics Interchange Format.gif",
                                  ARTIFACTS_DIR + "BaseConversions.GifToPdf.pdf")
        #ExEnd:ImageToPdf

    #ExStart:ConvertImageToPdf
    #GistId:36a49a29062268dc5e6d3134163f8d99
    @staticmethod
    def convert_image_to_pdf(input_file_name: str, output_file_name: str):
        """Converts an image to PDF using Aspose.Words for .NET.

       :param input_file_name: File name of input image file.
       :param output_file_name: Output PDF file name.
        """
        print("Converting " + input_file_name + " to PDF...")
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Read the image from file
        with drawing.Image.from_file(input_file_name) as image:
            # Find which dimension the frames in this image represent. For example
            # the frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension".
            dimension = drawing.imaging.FrameDimension(image.frame_dimensions_list[0])

            frames_count = image.get_frame_count(dimension)

            for frame_idx in range(frames_count):
                # Insert a section break before each new page, in case of a multi-frame TIFF.
                if frame_idx != 0:
                    builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

                image.select_active_frame(dimension, frame_idx)

                frame_stream = io.BytesIO()
                image.save(frame_stream, drawing.imaging.ImageFormat.png)

                # We want the size of the page to be the same as the size of the image.
                # Convert pixels to points to size the page to the actual image size.
                page_setup = builder.page_setup
                page_setup.page_width = aw.ConvertUtil.pixel_to_point(image.width, image.horizontal_resolution)
                page_setup.page_height = aw.ConvertUtil.pixel_to_point(image.height, image.vertical_resolution)

                # Insert the image into the document and position it at the top left corner of the page.
                builder.insert_image(
                    frame_stream,
                    aw.drawing.RelativeHorizontalPosition.PAGE,
                    0,
                    aw.drawing.RelativeVerticalPosition.PAGE,
                    0,
                    page_setup.page_width,
                    page_setup.page_height,
                    aw.drawing.WrapType.NONE)

        doc.save(output_file_name)
        #ExEnd:ConvertImageToPdf
