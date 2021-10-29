import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExUtilityClasses(ApiExampleBase):

    def test_points_and_inches(self):

        #ExStart
        #ExFor:aw.ConvertUtil
        #ExFor:aw.ConvertUtil.PointToInch
        #ExFor:aw.ConvertUtil.InchToPoint
        #ExSummary:Shows how to specify page properties in inches.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A section's "Page Setup" defines the size of the page margins in points.
        # We can also use the "aw.ConvertUtil" class to use a more familiar measurement unit,
        # such as inches when defining boundaries.
        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.inch_to_point(1.0)
        page_setup.bottom_margin = aw.ConvertUtil.inch_to_point(2.0)
        page_setup.left_margin = aw.ConvertUtil.inch_to_point(2.5)
        page_setup.right_margin = aw.ConvertUtil.inch_to_point(1.5)

        # An inch is 72 points.
        self.assertEqual(72.0, aw.ConvertUtil.inch_to_point(1))
        self.assertEqual(1.0, aw.ConvertUtil.point_to_inch(72))

        # Add content to demonstrate the new margins.
        builder.writeln(
            f"This Text is {page_setup.left_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.left_margin)} inches from the left, " +
            f"{page_setup.right_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.right_margin)} inches from the right, " +
            f"{page_setup.top_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.top_margin)} inches from the top, " +
            f"and {page_setup.bottom_margin} points/{aw.ConvertUtil.point_to_inch(page_setup.bottom_margin)} inches from the bottom of the page.")

        doc.save(ARTIFACTS_DIR + "UtilityClasses.PointsAndInches.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "UtilityClasses.PointsAndInches.docx")
        page_setup = doc.first_section.page_setup

        self.assertAlmostEqual(72.0, page_setup.top_margin, 2)
        self.assertAlmostEqual(1.0, aw.ConvertUtil.point_to_inch(page_setup.top_margin), 2)
        self.assertAlmostEqual(144.0, page_setup.bottom_margin, 2)
        self.assertAlmostEqual(2.0, aw.ConvertUtil.point_to_inch(page_setup.bottom_margin), 2)
        self.assertAlmostEqual(180.0, page_setup.left_margin, 2)
        self.assertAlmostEqual(2.5, aw.ConvertUtil.point_to_inch(page_setup.left_margin), 2)
        self.assertAlmostEqual(108.0, page_setup.right_margin, 2)
        self.assertAlmostEqual(1.5, aw.ConvertUtil.point_to_inch(page_setup.right_margin), 2)

    def test_points_and_millimeters(self):

        #ExStart
        #ExFor:aw.ConvertUtil.MillimeterToPoint
        #ExSummary:Shows how to specify page properties in millimeters.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A section's "Page Setup" defines the size of the page margins in points.
        # We can also use the "ConvertUtil" class to use a more familiar measurement unit,
        # such as millimeters when defining boundaries.
        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.millimeter_to_point(30)
        page_setup.bottom_margin = aw.ConvertUtil.millimeter_to_point(50)
        page_setup.left_margin = aw.ConvertUtil.millimeter_to_point(80)
        page_setup.right_margin = aw.ConvertUtil.millimeter_to_point(40)

        # A centimeter is approximately 28.3 points.
        self.assertAlmostEqual(28.34, aw.ConvertUtil.millimeter_to_point(10), 1)

        # Add content to demonstrate the new margins.
        builder.writeln(
            f"This Text is {page_setup.left_margin} points from the left, " +
            f"{page_setup.right_margin} points from the right, " +
            f"{page_setup.top_margin} points from the top, " +
            f"and {page_setup.bottom_margin} points from the bottom of the page.")

        doc.save(ARTIFACTS_DIR + "UtilityClasses.PointsAndMillimeters.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "UtilityClasses.PointsAndMillimeters.docx")
        page_setup = doc.first_section.page_setup

        self.assertAlmostEqual(85.05, page_setup.top_margin, 2)
        self.assertAlmostEqual(141.75, page_setup.bottom_margin, 2)
        self.assertAlmostEqual(226.75, page_setup.left_margin, 2)
        self.assertAlmostEqual(113.4, page_setup.right_margin, 2)

    def test_points_and_pixels(self):

        #ExStart
        #ExFor:aw.ConvertUtil.PixelToPoint(double)
        #ExFor:aw.ConvertUtil.PointToPixel(double)
        #ExSummary:Shows how to specify page properties in pixels.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A section's "Page Setup" defines the size of the page margins in points.
        # We can also use the "ConvertUtil" class to use a different measurement unit,
        # such as pixels when defining boundaries.
        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.pixel_to_point(100)
        page_setup.bottom_margin = aw.ConvertUtil.pixel_to_point(200)
        page_setup.left_margin = aw.ConvertUtil.pixel_to_point(225)
        page_setup.right_margin = aw.ConvertUtil.pixel_to_point(125)

        # A pixel is 0.75 points.
        self.assertEqual(0.75, aw.ConvertUtil.pixel_to_point(1))
        self.assertEqual(1.0, aw.ConvertUtil.point_to_pixel(0.75))

        # The default DPI value used is 96.
        self.assertEqual(0.75, aw.ConvertUtil.pixel_to_point(1, 96))

        # Add content to demonstrate the new margins.
        builder.writeln(
            f"This Text is {page_setup.left_margin} points/{aw.ConvertUtil.point_to_pixel(page_setup.left_margin)} pixels from the left, " +
            f"{page_setup.right_margin} points/{aw.ConvertUtil.point_to_pixel(page_setup.right_margin)} pixels from the right, " +
            f"{page_setup.top_margin} points/{aw.ConvertUtil.point_to_pixel(page_setup.top_margin)} pixels from the top, " +
            f"and {page_setup.bottom_margin} points/{aw.ConvertUtil.point_to_pixel(page_setup.bottom_margin)} pixels from the bottom of the page.")

        doc.save(ARTIFACTS_DIR + "UtilityClasses.PointsAndPixels.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "UtilityClasses.PointsAndPixels.docx")
        page_setup = doc.first_section.page_setup

        self.assertAlmostEqual(75.0, page_setup.top_margin, 2)
        self.assertAlmostEqual(100.0, aw.ConvertUtil.point_to_pixel(page_setup.top_margin), 2)
        self.assertAlmostEqual(150.0, page_setup.bottom_margin, 2)
        self.assertAlmostEqual(200.0, aw.ConvertUtil.point_to_pixel(page_setup.bottom_margin), 2)
        self.assertAlmostEqual(168.75, page_setup.left_margin, 2)
        self.assertAlmostEqual(225.0, aw.ConvertUtil.point_to_pixel(page_setup.left_margin), 2)
        self.assertAlmostEqual(93.75, page_setup.right_margin, 2)
        self.assertAlmostEqual(125.0, aw.ConvertUtil.point_to_pixel(page_setup.right_margin), 2)

    def test_points_and_pixels_dpi(self):

        #ExStart
        #ExFor:aw.ConvertUtil.PixelToNewDpi
        #ExFor:aw.ConvertUtil.PixelToPoint(double, double)
        #ExFor:aw.ConvertUtil.PointToPixel(double, double)
        #ExSummary:Shows how to use convert points to pixels with default and custom resolution.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Define the size of the top margin of this section in pixels, according to a custom DPI.
        MY_DPI = 192

        page_setup = builder.page_setup
        page_setup.top_margin = aw.ConvertUtil.pixel_to_point(100, MY_DPI)

        self.assertAlmostEqual(37.5, page_setup.top_margin, 2)

        # At the default DPI of 96, a pixel is 0.75 points.
        self.assertEqual(0.75, aw.ConvertUtil.pixel_to_point(1))

        builder.writeln(
            f"This Text is {page_setup.top_margin} points/{aw.ConvertUtil.point_to_pixel(page_setup.top_margin, MY_DPI)} " +
            f"pixels (at a DPI of {MY_DPI}) from the top of the page.")

        # Set a new DPI and adjust the top margin value accordingly.
        NEW_DPI = 300
        page_setup.top_margin = aw.ConvertUtil.pixel_to_new_dpi(page_setup.top_margin, MY_DPI, NEW_DPI)
        self.assertEqual(59.0, page_setup.top_margin, 0.01)

        builder.writeln(
            f"At a DPI of {NEW_DPI}, the text is now {page_setup.top_margin} points/{aw.ConvertUtil.point_to_pixel(page_setup.top_margin, MY_DPI)} " +
            "pixels from the top of the page.")

        doc.save(ARTIFACTS_DIR + "UtilityClasses.PointsAndPixelsDpi.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "UtilityClasses.PointsAndPixelsDpi.docx")
        page_setup = doc.first_section.page_setup

        self.assertAlmostEqual(59.0, page_setup.top_margin, 2)
        self.assertAlmostEqual(78.66, aw.ConvertUtil.point_to_pixel(page_setup.top_margin), 1)
        self.assertAlmostEqual(157.33, aw.ConvertUtil.point_to_pixel(page_setup.top_margin, MY_DPI), 2)
        self.assertAlmostEqual(133.33, aw.ConvertUtil.point_to_pixel(100), 2)
        self.assertAlmostEqual(266.66, aw.ConvertUtil.point_to_pixel(100, MY_DPI), 1)
