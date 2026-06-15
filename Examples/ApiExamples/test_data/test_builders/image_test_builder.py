# -*- coding: utf-8 -*-

from api_example_base import ApiExampleBase


class ImageTestBuilder(ApiExampleBase):
    def __init__(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import io
        import aspose.words as aw
        
        self.m_image_stream = None
        self.m_image = None
        m_image = aw.drawing.Image.from_file(IMAGE_DIR + "Transparent background logo.png")
        m_image_stream = io.BytesIO()
        self.m_image_bytes = []
        self.m_image_string = ""

    def with_image(self, image_path):
            mImage = aw.Image.from_file(imagePath)
            
            return self

    def with_image_stream(self, image_stream):
        self.m_image_stream = image_stream
        return self

    def with_image_bytes(self, image_bytes):
        self.m_image_bytes = image_bytes
        return self

    def with_image_string(self, image_string):
        self.m_image_string = image_string
        return self

    def build(self):
        return ImageTestClass(self.m_image,self.m_image_stream,self.m_image_bytes,self.m_image_string)
