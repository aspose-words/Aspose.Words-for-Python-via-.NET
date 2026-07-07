import io

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, IMAGES_DIR

class WorkingWithOleObjectsAndActiveX(DocsExamplesBase):

    def test_insert_ole_object(self):

        #ExStart:InsertOleObject
        #GistId:2c7d2bb472db1a698568d980c01b0735
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_ole_object("http://www.aspose.com", "htmlfile", True, True, None)

        doc.save(ARTIFACTS_DIR + "WorkingWithOleObjectsAndActiveX.insert_ole_object.docx")
        #ExEnd:InsertOleObject

    def test_insert_ole_object_with_ole_package(self):

        #ExStart:InsertOleObjectwithOlePackage
        #GistId:2c7d2bb472db1a698568d980c01b0735
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        with io.FileIO(MY_DIR + "Zip file.zip") as stream:
            shape = builder.insert_ole_object(stream, "Package", True, None)

        ole_package = shape.ole_format.ole_package
        ole_package.file_name = "filename.zip"
        ole_package.display_name = "displayname.zip"

        doc.save(ARTIFACTS_DIR + "WorkingWithOleObjectsAndActiveX.insert_ole_object_with_ole_package.docx")
        #ExEnd:InsertOleObjectwithOlePackage

        #ExStart:GetAccessToOleObjectRawData
        #GistId:2c7d2bb472db1a698568d980c01b0735
        ole_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        ole_raw_data = ole_shape.ole_format.get_raw_data()
        #ExEnd:GetAccessToOleObjectRawData

    def test_insert_ole_object_as_icon(self):

        #ExStart:InsertOleObjectAsIcon
        #GistId:2c7d2bb472db1a698568d980c01b0735
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_ole_object_as_icon(MY_DIR + "Presentation.pptx", False, IMAGES_DIR + "Logo icon.ico", "My embedded file")

        doc.save(ARTIFACTS_DIR + "WorkingWithOleObjectsAndActiveX.insert_ole_object_as_icon.docx")
        #ExEnd:InsertOleObjectAsIcon

    def test_insert_ole_object_as_icon_using_stream(self):

        #ExStart:InsertOleObjectAsIconUsingStream
        #GistId:2c7d2bb472db1a698568d980c01b0735
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        with io.FileIO(MY_DIR + "Presentation.pptx") as stream:
            builder.insert_ole_object_as_icon(stream, "Package", IMAGES_DIR + "Logo icon.ico", "My embedded file")

        doc.save(ARTIFACTS_DIR + "WorkingWithOleObjectsAndActiveX.insert_ole_object_as_icon_using_stream.docx")
        #ExEnd:InsertOleObjectAsIconUsingStream

    def test_read_active_x_control_properties(self):

        doc = aw.Document(MY_DIR + "ActiveX controls.docx")

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
        #GistId:2c7d2bb472db1a698568d980c01b0735
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        url = "https://youtu.be/t_1LYZ102RA"
        width = 360
        height = 270

        builder.insert_online_video(url, width, height)

        doc.save(ARTIFACTS_DIR + "WorkingWithOleObjectsAndActiveX.insert_online_video.docx")
        #ExEnd:InsertOnlineVideo

    def test_insert_online_video_with_embed_html(self):

        #ExStart:InsertOnlineVideoWithEmbedHtml
        #GistId:2c7d2bb472db1a698568d980c01b0735
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        width = 360
        height = 270

        vimeo_video_url = "https://vimeo.com/52477838"
        vimeo_embed_code = "<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" " \
        "title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>"

        image_bytes = open(IMAGES_DIR + "Logo.jpg", "rb").read()
        
        builder.insert_online_video(vimeo_video_url, vimeo_embed_code, image_bytes, width, height)

        doc.save(ARTIFACTS_DIR + "WorkingWithOleObjectsAndActiveX.insert_online_video_with_embed_html.docx")
        #ExEnd:InsertOnlineVideoWithEmbedHtml
