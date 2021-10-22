import unittest
import os
import sys
import io

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithOleObjectsAndActiveX(docs_base.DocsExamplesBase):

    def test_insert_ole_object(self):

        #ExStart:DocumentBuilderInsertOleObject
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_ole_object("http://www.aspose.com", "htmlfile", True, True, None)

        doc.save(docs_base.artifacts_dir + "WorkingWithOleObjectsAndActiveX.insert_ole_object.docx")
        #ExEnd:DocumentBuilderInsertOleObject


    def test_insert_ole_object_with_ole_package(self):

        #ExStart:InsertOleObjectwithOlePackage
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        stream = io.FileIO(docs_base.my_dir + "Zip file.zip")

        shape = builder.insert_ole_object(stream, "Package", True, None)
        ole_package = shape.ole_format.ole_package
        ole_package.file_name = "filename.zip"
        ole_package.display_name = "displayname.zip"

        doc.save(docs_base.artifacts_dir + "WorkingWithOleObjectsAndActiveX.insert_ole_object_with_ole_package.docx")

        stream.close()
        #ExEnd:InsertOleObjectwithOlePackage

        #ExStart:GetAccessToOLEObjectRawData
        ole_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        ole_raw_data = ole_shape.ole_format.get_raw_data()
        #ExEnd:GetAccessToOLEObjectRawData


    def test_insert_ole_object_as_icon(self):

        #ExStart:InsertOLEObjectAsIcon
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_ole_object_as_icon(docs_base.my_dir + "Presentation.pptx", False, docs_base.images_dir + "Logo icon.ico", "My embedded file")

        doc.save(docs_base.artifacts_dir + "WorkingWithOleObjectsAndActiveX.insert_ole_object_as_icon.docx")
        #ExEnd:InsertOLEObjectAsIcon


    def test_insert_ole_object_as_icon_using_stream(self):

        #ExStart:InsertOLEObjectAsIconUsingStream
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        stream = io.FileIO(docs_base.my_dir + "Presentation.pptx")
        builder.insert_ole_object_as_icon(stream, "Package", docs_base.images_dir + "Logo icon.ico", "My embedded file")

        doc.save(docs_base.artifacts_dir + "WorkingWithOleObjectsAndActiveX.insert_ole_object_as_icon_using_stream.docx")

        stream.close()
        #ExEnd:InsertOLEObjectAsIconUsingStream


    def test_read_active_x_control_properties(self):

        doc = aw.Document(docs_base.my_dir + "ActiveX controls.docx")

        properties = ""
        for shape in doc.get_child_nodes(aw.NodeType.SHAPE, True):

            shape = shape.as_shape()

            if shape.ole_format is None:
                break

            ole_control = shape.ole_format.ole_control
            if ole_control.is_forms2_ole_control:

                check_box =  ole_control.as_forms2_ole_control()
                properties = properties + "\nCaption: " + check_box.caption
                properties = properties + "\nValue: " + check_box.value
                properties = properties + "\nEnabled: " + str(check_box.enabled)
                properties = properties + "\nType: " + str(check_box.type)

                if check_box.child_nodes is not None:
                    properties = properties + "\nChildNodes: " + check_box.child_nodes

                properties += "\n"

        properties = properties + "\nTotal ActiveX Controls found: " + str(doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
        print("\n" + properties)

    def test_insert_online_video(self):

        #ExStart:InsertOnlineVideo
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Pass direct url from youtu.be.
        url = "https://youtu.be/t_1LYZ102RA"

        width = 360
        height = 270

        shape = builder.insert_online_video(url, width, height)

        doc.save(docs_base.artifacts_dir + "WorkingWithOleObjectsAndActiveX.insert_online_video.docx")
        #ExEnd:InsertOnlineVideo

    def test_insert_online_video_with_embed_html(self):

        #ExStart:InsertOnlineVideoWithEmbedHtml
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Shape width/height.
        width = 360
        height = 270

        # Poster frame image.
        image = open(docs_base.images_dir + "Logo.jpg", "rb")
        image_bytes = image.read()
        image.close()

        # Visible url
        vimeo_video_url = "https://vimeo.com/52477838"

        # Embed Html code.
        vimeo_embed_code = ""

        builder.insert_online_video(vimeo_video_url, vimeo_embed_code, image_bytes, width, height)

        doc.save(docs_base.artifacts_dir + "WorkingWithOleObjectsAndActiveX.insert_online_video_with_embed_html.docx")
        #ExEnd:InsertOnlineVideoWithEmbedHtml


if __name__ == '__main__':
    unittest.main()
