# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR

class ExUtilityClasses(ApiExampleBase):

    def test_points_and_inches(self):
        #ExStart
        #ExFor:ConvertUtil
        #ExFor:ConvertUtil.point_to_inch
        #ExFor:ConvertUtil.inch_to_point
        #ExSummary:Shows how to specify page properties in inches.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # A section's "Page Setup" defines the size of the page margins in points.
        # We can also use the "ConvertUtil" class to use a more familiar measurement unit,
        # such as inches when defining boundaries.
        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.inch_to_point(1)
        page_setup.bottom_margin = aw.ConvertUtil.inch_to_point(2)
        page_setup.left_margin = aw.ConvertUtil.inch_to_point(2.5)
        page_setup.right_margin = aw.ConvertUtil.inch_to_point(1.5)
        # An inch is 72 points.
        self.assertEqual(72, aw.ConvertUtil.inch_to_point(1))
        self.assertEqual(1, aw.ConvertUtil.point_to_inch(72))
        # Add content to demonstrate the new margins.
        builder.writeln(f'This Text is {page_setup.left_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.left_margin)} inches from the left, ' + f'{page_setup.right_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.right_margin)} inches from the right, ' + f'{page_setup.top_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.top_margin)} inches from the top, ' + f'and {page_setup.bottom_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.bottom_margin)} inches from the bottom of the page.')
        doc.save(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndInches.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndInches.docx')
        page_setup = doc.first_section.page_setup
        self.assertAlmostEqual(72, page_setup.top_margin, delta=0.01)
        self.assertAlmostEqual(1, aw.ConvertUtil.point_to_inch(page_setup.top_margin), delta=0.01)
        self.assertAlmostEqual(144, page_setup.bottom_margin, delta=0.01)
        self.assertAlmostEqual(2, aw.ConvertUtil.point_to_inch(page_setup.bottom_margin), delta=0.01)
        self.assertAlmostEqual(180, page_setup.left_margin, delta=0.01)
        self.assertAlmostEqual(2.5, aw.ConvertUtil.point_to_inch(page_setup.left_margin), delta=0.01)
        self.assertAlmostEqual(108, page_setup.right_margin, delta=0.01)
        self.assertAlmostEqual(1.5, aw.ConvertUtil.point_to_inch(page_setup.right_margin), delta=0.01)

    def test_points_and_millimeters(self):
        #ExStart
        #ExFor:ConvertUtil.millimeter_to_point
        #ExSummary:Shows how to specify page properties in millimeters.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # A section's "Page Setup" defines the size of the page margins in points.
        # We can also use the "ConvertUtil" class to use a more familiar measurement unit,
        # such as millimeters when defining boundaries.
        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.millimeter_to_point(30)
        page_setup.bottom_margin = aw.ConvertUtil.millimeter_to_point(50)
        page_setup.left_margin = aw.ConvertUtil.millimeter_to_point(80)
        page_setup.right_margin = aw.ConvertUtil.millimeter_to_point(40)
        # A centimeter is approximately 28.3 points.
        self.assertAlmostEqual(28.34, aw.ConvertUtil.millimeter_to_point(10), delta=0.01)
        # Add content to demonstrate the new margins.
        builder.writeln(f'This Text is {page_setup.left_margin} points from the left, ' + f'{page_setup.right_margin} points from the right, ' + f'{page_setup.top_margin} points from the top, ' + f'and {page_setup.bottom_margin} points from the bottom of the page.')
        doc.save(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndMillimeters.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndMillimeters.docx')
        page_setup = doc.first_section.page_setup
        self.assertAlmostEqual(85.05, page_setup.top_margin, delta=0.01)
        self.assertAlmostEqual(141.75, page_setup.bottom_margin, delta=0.01)
        self.assertAlmostEqual(226.75, page_setup.left_margin, delta=0.01)
        self.assertAlmostEqual(113.4, page_setup.right_margin, delta=0.01)

    def test_points_and_pixels(self):
        #ExStart
        #ExFor:ConvertUtil.pixel_to_point(float)
        #ExFor:ConvertUtil.point_to_pixel(float)
        #ExSummary:Shows how to specify page properties in pixels.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # A section's "Page Setup" defines the size of the page margins in points.
        # We can also use the "ConvertUtil" class to use a different measurement unit,
        # such as pixels when defining boundaries.
        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.pixel_to_point(pixels=100)
        page_setup.bottom_margin = aw.ConvertUtil.pixel_to_point(pixels=200)
        page_setup.left_margin = aw.ConvertUtil.pixel_to_point(pixels=225)
        page_setup.right_margin = aw.ConvertUtil.pixel_to_point(pixels=125)
        # A pixel is 0.75 points.
        self.assertEqual(0.75, aw.ConvertUtil.pixel_to_point(pixels=1))
        self.assertEqual(1, aw.ConvertUtil.point_to_pixel(points=0.75))
        # The default DPI value used is 96.
        self.assertEqual(0.75, aw.ConvertUtil.pixel_to_point(pixels=1, resolution=96))
        # Add content to demonstrate the new margins.
        builder.writeln(f'This Text is {page_setup.left_margin} points/{aw.ConvertUtil.point_to_pixel(points=page_setup.left_margin)} pixels from the left, ' + f'{page_setup.right_margin} points/{aw.ConvertUtil.point_to_pixel(points=page_setup.right_margin)} pixels from the right, ' + f'{page_setup.top_margin} points/{aw.ConvertUtil.point_to_pixel(points=page_setup.top_margin)} pixels from the top, ' + f'and {page_setup.bottom_margin} points/{aw.ConvertUtil.point_to_pixel(points=page_setup.bottom_margin)} pixels from the bottom of the page.')
        doc.save(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndPixels.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndPixels.docx')
        page_setup = doc.first_section.page_setup
        self.assertAlmostEqual(75, page_setup.top_margin, delta=0.01)
        self.assertAlmostEqual(100, aw.ConvertUtil.point_to_pixel(points=page_setup.top_margin), delta=0.01)
        self.assertAlmostEqual(150, page_setup.bottom_margin, delta=0.01)
        self.assertAlmostEqual(200, aw.ConvertUtil.point_to_pixel(points=page_setup.bottom_margin), delta=0.01)
        self.assertAlmostEqual(168.75, page_setup.left_margin, delta=0.01)
        self.assertAlmostEqual(225, aw.ConvertUtil.point_to_pixel(points=page_setup.left_margin), delta=0.01)
        self.assertAlmostEqual(93.75, page_setup.right_margin, delta=0.01)
        self.assertAlmostEqual(125, aw.ConvertUtil.point_to_pixel(points=page_setup.right_margin), delta=0.01)

    def test_points_and_pixels_dpi(self):
        #ExStart
        #ExFor:ConvertUtil.pixel_to_new_dpi
        #ExFor:ConvertUtil.pixel_to_point(float,float)
        #ExFor:ConvertUtil.point_to_pixel(float,float)
        #ExSummary:Shows how to use convert points to pixels with default and custom resolution.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Define the size of the top margin of this section in pixels, according to a custom DPI.
        my_dpi = 192
        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.pixel_to_point(pixels=100, resolution=my_dpi)
        self.assertAlmostEqual(37.5, page_setup.top_margin, delta=0.01)
        # At the default DPI of 96, a pixel is 0.75 points.
        self.assertEqual(0.75, aw.ConvertUtil.pixel_to_point(pixels=1))
        builder.writeln(f'This Text is {page_setup.top_margin} points/{aw.ConvertUtil.point_to_pixel(points=page_setup.top_margin, resolution=my_dpi)} ' + f'pixels (at a DPI of {my_dpi}) from the top of the page.')
        # Set a new DPI and adjust the top margin value accordingly.
        new_dpi = 300
        page_setup.top_margin = aw.ConvertUtil.pixel_to_new_dpi(page_setup.top_margin, my_dpi, new_dpi)
        self.assertAlmostEqual(59, page_setup.top_margin, delta=0.01)
        builder.writeln(f'At a DPI of {new_dpi}, the text is now {page_setup.top_margin} points/{aw.ConvertUtil.point_to_pixel(points=page_setup.top_margin, resolution=my_dpi)} ' + 'pixels from the top of the page.')
        doc.save(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndPixelsDpi.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'UtilityClasses.PointsAndPixelsDpi.docx')
        page_setup = doc.first_section.page_setup
        self.assertAlmostEqual(59, page_setup.top_margin, delta=0.01)
        self.assertAlmostEqual(78.66, aw.ConvertUtil.point_to_pixel(points=page_setup.top_margin), delta=0.01)
        self.assertAlmostEqual(157.33, aw.ConvertUtil.point_to_pixel(points=page_setup.top_margin, resolution=my_dpi), delta=0.01)
        self.assertAlmostEqual(133.33, aw.ConvertUtil.point_to_pixel(points=100), delta=0.01)
        self.assertAlmostEqual(266.66, aw.ConvertUtil.point_to_pixel(points=100, resolution=my_dpi), delta=0.01)