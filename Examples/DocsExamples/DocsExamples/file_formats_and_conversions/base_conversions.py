import io
import unittest

import aspose.words as aw
import aspose.pydrawing as drawing

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, IMAGES_DIR

class BaseConversions(DocsExamplesBase):

    def test_doc_to_docx(self):

        #ExStart:LoadAndSave
        #ExStart:OpenDocument
        doc = aw.Document(MY_DIR + "Document.doc")
        #ExEnd:OpenDocument

        doc.save(ARTIFACTS_DIR + "BaseConversions.doc_to_docx.docx")
        #ExEnd:LoadAndSave

    def test_docx_to_rtf(self):

        #ExStart:LoadAndSaveToStream
        #ExStart:OpeningFromStream
        # Read only access is enough for Aspose.words to load a document.
        stream = io.FileIO(MY_DIR + "Document.docx")

        doc = aw.Document(stream)
        # You can close the stream now, it is no longer needed because the document is in memory.
        stream.close()
        #ExEnd:OpeningFromStream

        # ... do something with the document.

        # Convert the document to a different format and save to stream.
        dst_stream = io.BytesIO()
        doc.save(dst_stream, aw.SaveFormat.RTF)
        #ExEnd:LoadAndSaveToStream

        with open(ARTIFACTS_DIR + "BaseConversions.docx_to_rtf.rtf", "wb") as output:
            output.write(dst_stream.getbuffer())

    def test_docx_to_pdf(self):

        #ExStart:Doc2Pdf
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_pdf.pdf")
        #ExEnd:Doc2Pdf

    def test_docx_to_byte(self):

        #ExStart:DocxToByte
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

    @unittest.skip("Aspose.Email is required. Will do later.")
    def test_docx_to_mhtml_and_sending_email(self):
        print("not supported yet")
#        #ExStart:DocxToMhtmlAndSendingEmail
#        doc = aw.Document(MY_DIR + "Document.docx")
#
#        Stream stream = new MemoryStream()
#        doc.save(stream, SaveFormat.mhtml)
#
#        # Rewind the stream to the beginning so Aspose.email can read it.
#        stream.position = 0
#
#        # Create an Aspose.network MIME email message from the stream.
#        MailMessage message = MailMessage.load(stream, new MhtmlLoadOptions())
#        message.from = "your_from@email.com"
#        message.to = "your_to@email.com"
#        message.subject = "Aspose.words + Aspose.email MHTML Test Message"
#
#        # Send the message using Aspose.email.
#        SmtpClient client = new SmtpClient()
#        client.host = "your_smtp.com"
#        client.send(message)
#        #ExEnd:DocxToMhtmlAndSendingEmail

    def test_docx_to_markdown(self):

        #ExStart:SaveToMarkdownDocument
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Some text!")

        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_markdown.md")
        #ExEnd:SaveToMarkdownDocument

    def test_docx_to_txt(self):

        #ExStart:DocxToTxt
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.save(ARTIFACTS_DIR + "BaseConversions.docx_to_txt.txt")
        #ExEnd:DocxToTxt

    def test_txt_to_docx(self):

        #ExStart:TxtToDocx
        # The encoding of the text file is automatically detected.
        doc = aw.Document(MY_DIR + "English text.txt")

        doc.save(ARTIFACTS_DIR + "BaseConversions.txt_to_docx.docx")
        #ExEnd:TxtToDocx

    def test_pdf_to_jpeg(self):

        #ExStart:PdfToJpeg
        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        doc.save(ARTIFACTS_DIR + "BaseConversions.pdf_to_jpeg.jpeg")
        #ExEnd:PdfToJpeg

    def test_pdf_to_docx(self):

        #ExStart:PdfToDocx
        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        doc.save(ARTIFACTS_DIR + "BaseConversions.pdf_to_docx.docx")
        #ExEnd:PdfToDocx

    def test_images_to_pdf(self):

        #ExStart:ImageToPdf
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

    @staticmethod
    def convert_image_to_pdf(input_file_name: str, output_file_name: str):
        """Converts an image to PDF using Aspose.Words for .NET.

       :param input_file_name: File name of input image file.
       :param output_file_name: Output PDF file name.
        """
        print(f"Converting {input_file_name} to PDF ....")

        #ExStart:ConvertImageToPdf
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

                # We want the size of the page to be the same as the size of the image.
                # Convert pixels to points to size the page to the actual image size.
                page_setup = builder.page_setup
                page_setup.page_width = aw.ConvertUtil.pixel_to_point(image.width, image.horizontal_resolution)
                page_setup.page_height = aw.ConvertUtil.pixel_to_point(image.height, image.vertical_resolution)

                # Insert the image into the document and position it at the top left corner of the page.
                builder.insert_image(
                    image,
                    aw.drawing.RelativeHorizontalPosition.PAGE,
                    0,
                    aw.drawing.RelativeVerticalPosition.PAGE,
                    0,
                    page_setup.page_width,
                    page_setup.page_height,
                    aw.drawing.WrapType.NONE)

        doc.save(output_file_name)
        #ExEnd:ConvertImageToPdf
