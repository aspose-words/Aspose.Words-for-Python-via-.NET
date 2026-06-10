# -*- coding: utf-8 -*-

import aspose.pydrawing


class ColorItemTestBuilder:
    def __init__(self):
        self.name = "DefaultName"
        self.color = aspose.pydrawing.Color.black
        self.color_code = aspose.pydrawing.Color.black.to_argb()
        self.value1 = 1
        self.value2 = 1
        self.value3 = 1

    def with_color(self, name, color):
        self.name = name
        self.color = color
        return self

    def with_color_code(self, name, color_code):
        self.name = name
        self.color_code = color_code
        return self

    def with_color_and_values(self, name, color, value1, value2, value3):
        self.name = name
        self.color = color
        self.value1 = value1
        self.value2 = value2
        self.value3 = value3
        return self

    def with_color_code_and_values(self, name, color_code, value1, value2, value3):
        self.name = name
        self.color_code = color_code
        self.value1 = value1
        self.value2 = value2
        self.value3 = value3
        return self

    def build(self):
        return ColorItemTestClass(self.name,self.color,self.color_code,self.value1,self.value2,self.value3)
