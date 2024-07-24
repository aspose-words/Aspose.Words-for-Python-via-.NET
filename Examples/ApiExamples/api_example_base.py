# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import urllib.request
import os
import io
import zipfile
from datetime import date, time, datetime, timedelta
from typing import Optional
import platform
if not platform.python_version().startswith("3.7") and not platform.python_version().startswith("3.6"):
    from PIL import Image

import aspose.words as aw
import aspose.pydrawing as drawing


root_dir = os.getenv("ROOT_DIR")
if root_dir is not None:
    ROOT_DIR = root_dir
else:
    ROOT_DIR = os.path.abspath(os.curdir) + "/"
    ROOT_DIR = ROOT_DIR[:ROOT_DIR.find("Aspose.Words-for-Python-via-.NET")]

API_EXAMPLES_ROOT = ROOT_DIR + "Aspose.Words-for-Python-via-.NET/Examples/"
LICENSE_PATH = os.getenv("ASPOSE_WORDS_PYTHON_LICENSE", "Aspose.Words.Python.NET.lic")

# Path to the documents used by the code examples.
MY_DIR = API_EXAMPLES_ROOT + "Data/"

# Path to the documents used by the code examples.
ARTIFACTS_DIR = MY_DIR + "Artifacts/"

# Path to the documents used by the code examples.
GOLDS_DIR = MY_DIR + "Golds/"

# Path to the temporary directory.
TEMP_DIR = MY_DIR + "Temp/"

# Path to the images used by the code examples.
IMAGE_DIR = MY_DIR + "Images/"

# Path to the free fonts.
FONTS_DIR = MY_DIR + "MyFonts/"

# Path to the demo database directory.
DATABASE_DIR = MY_DIR + "Database/"

# URL of the test image.
IMAGE_URL = "https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png"


class ApiExampleBase(unittest.TestCase):

    def setUp(self):
        if os.path.exists(LICENSE_PATH):
            lic = aw.License()
            lic.set_license(LICENSE_PATH)
        if not os.path.exists(ARTIFACTS_DIR):
            os.makedirs(ARTIFACTS_DIR)

    def assertIn(self, member, container, msg=None):
        if member in container:
            return

        if isinstance(container, str) and len(container) > 256:
            container = container[:256] + '...'

        if isinstance(container, bytes) and len(container) > 256:
            container = container[:256] + b'...'

        unittest.TestCase.assertIn(self, member, container, msg=msg)

    def verify_image(self, expected_width: int, expected_height: int, filename: Optional[str] = None, image_stream: Optional[io.BytesIO] = None):
        """Checks whether a file or a stream contains a valid image with specified dimensions.

        Serves to check that an image file is valid and nonempty without looking up its file size.

        :param expected_width: Expected width of the image, in pixels.
        :param expected_height: Expected height of the image, in pixels.
        :param filename: Local file system filename of the image file.
        :param image_stream: Stream that contains the image."""

        assert filename is None or image_stream is None
        assert filename is not None or image_stream is not None

        if not platform.python_version().startswith("3.7") and not platform.python_version().startswith("3.6"):
            if filename is not None:
                with open(filename, 'rb') as stream:
                    image = Image.open(stream)
                    self.assertEqual(expected_width, image.width)
                    self.assertEqual(expected_height, image.height)
            else:
                image = Image.open(image_stream)
                self.assertEqual(expected_width, image.width)
                self.assertEqual(expected_height, image.height)

            return

        if filename is not None:
            with open(filename, 'rb') as stream:
                with drawing.Image.from_stream(stream) as image:
                    self.assertEqual(expected_width, image.width)
                    self.assertEqual(expected_height, image.height)
        else:
            with drawing.Image.from_stream(image_stream) as image:
                self.assertEqual(expected_width, image.width)
                self.assertEqual(expected_height, image.height)

    def verify_image_contains_transparency(self, filename: str):
        """Checks whether an image from the local file system contains any transparency.

        :param filename: Local file system filename of the image file."""

        with drawing.Image.from_file(filename) as image:
            for x in range(image.width):
                for y in range(image.height):
                    if image.get_pixel(x, y).a != 255:
                        return

        raise Exception("The image from \"" + filename + "\" does not contain any transparency.")

    def verify_web_response_status_code(self, expected_http_status_code: int, web_address: str):
        """Checks whether an HTTP request sent to the specified address produces an expected web response.

        Serves as a notification of any URLs used in code examples becoming unusable in the future.

        :param expected_http_status_code: Expected result status code of a request HTTP "HEAD" method performed on the web address.
        :param web_address: URL where the request will be sent."""

        req = urllib.request.Request(web_address, method="HEAD")
        response = urllib.request.urlopen(req)

        self.assertEqual(expected_http_status_code, response.getcode())

    def verify_doc_package_file_contains_string(self, expected: str, doc_filename: str, doc_part_filename: str):
        """Checks whether a file inside a document's OOXML package contains a string.

        If an output document does not have a testable value that can be found as a property in its object when loaded,
        the value can sometimes be found in the document's OOXML package.

        :param expected: The string we are looking for.
        :param doc_filename: Local file system filename of the document.
        :param doc_part_filename: Name of the file within the document opened as a .zip that is expected to contain the string."""

        with zipfile.ZipFile(doc_filename) as archive:
            with archive.open(doc_part_filename) as stream:
                self.assertIn(expected.encode('utf-8'), stream.read())

    def verify_field(self,
                     expected_type: aw.fields.FieldType,
                     expected_field_code: str,
                     expected_result: str,
                     field: aw.fields.Field):
        """Checks whether values of properties of a field with a type not related to date/time are equal to expected values.

        Best used when there are many fields closely being tested and should be avoided if a field has a long field code/result.

        :param expected_type: The FieldType that we expect the field to have.
        :param expected_field_code: The expected output value of GetFieldCode() being called on the field.
        :param expected_result: The field's expected result, which will be the value displayed by it in the document.
        :param field: The field that's being tested.
        """

        self.assertEqual(expected_type, field.type)
        self.assertEqual(expected_field_code, field.get_field_code(True))
        self.assertEqual(expected_result, field.result)

    def verify_datetime_field(self,
                              expected_type: aw.fields.FieldType,
                              expected_field_code: str,
                              expected_result: datetime,
                              field: aw.fields.Field,
                              delta: timedelta):
        """Checks whether values of properties of a field with a type related to date/time are equal to expected values.

        Used when comparing DateTime instances to Field.Result values parsed to DateTime, which may differ slightly.
        Give a delta value that's generous enough for any lower end system to pass, also a delta of zero is allowed.

        :param expected_type: The FieldType that we expect the field to have.
        :param expected_field_code: The expected output value of GetFieldCode() being called on the field.
        :param expected_result: The date/time that the field's result is expected to represent.
        :param field: The field that's being tested.
        :param delta: Margin of error for expected_result."""

        self.assertEqual(expected_type, field.type)
        self.assertEqual(expected_field_code, field.get_field_code(True))

        if field.type == aw.fields.FieldType.FIELD_TIME:
            actual = datetime.strptime(field.result, "%H:%M:%S")
            expected = datetime.combine(date(1900, 1, 1), expected_result.time())
            self.assertAlmostEqual(expected, actual, delta=delta)
        else:
            actual = datetime.strptime(field.result, "%d/%m/%Y")
            expected = datetime.combine(expected_result.date(), time())
            self.assertAlmostEqual(expected, actual, delta=delta)

    def verify_image_in_shape(self,
                              expected_width: int,
                              expected_height: int,
                              expected_image_type: aw.drawing.ImageType,
                              image_shape: aw.drawing.Shape):
        """Checks whether a shape contains a valid image with specified dimensions.

        Serves to check that an image file is valid and nonempty without looking up its data length.

        :param expected_width: Expected width of the image, in pixels.
        :param expected_height: Expected height of the image, in pixels.
        :param expected_image_type: Expected format of the image.
        :param image_shape: Shape that contains the image.
        """

        self.assertTrue(image_shape.has_image)
        self.assertEqual(expected_image_type, image_shape.image_data.image_type)
        self.assertEqual(expected_width, image_shape.image_data.image_size.width_pixels)
        self.assertEqual(expected_height, image_shape.image_data.image_size.height_pixels)

    def verify_footnote(self,
                        expected_footnote_type: aw.notes.FootnoteType,
                        expected_is_auto: bool,
                        expected_reference_mark: str,
                        expected_contents: str,
                        footnote: aw.notes.Footnote):
        """Checks whether values of a footnote's properties are equal to their expected values.

        :param expected_footnote_type: Expected type of the footnote/endnote.</param>
        :param expected_is_auto: Expected auto-numbered status of this footnote.</param>
        :param expected_reference_mark: If "is_auto" is False, then the footnote is expected to display this string instead of a number after referenced text.</param>
        :param expected_contents: Expected side comment provided by the footnote.</param>
        :param footnote: Footnote node in question.</param>"""

        self.assertEqual(expected_footnote_type, footnote.footnote_type)
        self.assertEqual(expected_is_auto, footnote.is_auto)
        self.assertEqual(expected_reference_mark, footnote.reference_mark)
        self.assertEqual(expected_contents, footnote.to_string(aw.SaveFormat.TEXT).strip())

    def verify_list_level(self,
                          expected_list_format: str,
                          expected_number_position: float,
                          expected_number_style: aw.NumberStyle,
                          list_level: aw.lists.ListLevel):
        """Checks whether values of a list level's properties are equal to their expected values.

        Only necessary for list levels that have been explicitly created by the user.

        :param expected_list_format: Expected format for the list symbol.
        :param expected_number_position: Expected indent for this level, usually growing larger with each level.
        :param expected_number_style:
        :param list_level: List level in question."""

        self.assertEqual(expected_list_format, list_level.number_format)
        self.assertEqual(expected_number_position, list_level.number_position)
        self.assertEqual(expected_number_style, list_level.number_style)

    @staticmethod
    def copy_stream(src_stream: io.BytesIO, dst_stream: io.BytesIO):
        """Copies from the current position in src stream till the end.
        Copies into the current position in dst stream."""

        assert src_stream is not None
        assert dst_stream is not None

        dst_stream.write(src_stream.read())

    @staticmethod
    def dump_array(data: bytes, start: int, count: int) -> str:
        """Dumps byte array into a string."""

        if data is None:
            return "Null"

        result = ""
        while count > 0:
            result += "{:02X} ".format(data[start])
            start += 1
            count -= 1

        return result

    def verify_tab_stop(self,
                        expected_position: float,
                        expected_tab_alignment: aw.TabAlignment,
                        expected_tab_leader: aw.TabLeader,
                        is_clear: bool, tab_stop: aw.TabStop):
        """Checks whether values of a tab stop's properties are equal to their expected values.

        :param expected_position: Expected position on the tab stop ruler, in points.
        :param expected_tab_alignment: Expected position where the position is measured from.
        :param expected_tab_leader: Expected characters that pad the space between the start and end of the tab whitespace.
        :param is_clear: Whether or no this tab stop clears any tab stops.
        :param tab_stop: Tab stop that's being tested."""

        self.assertEqual(expected_position, tab_stop.position)
        self.assertEqual(expected_tab_alignment, tab_stop.alignment)
        self.assertEqual(expected_tab_leader, tab_stop.leader)
        self.assertEqual(is_clear, tab_stop.is_clear)

    def verify_shape(self, expected_shape_type: aw.drawing.ShapeType, expected_name: str, expected_width: float, expected_height: float, expected_top: float, expected_left: float, shape: aw.drawing.Shape):
        """Checks whether values of a shape's properties are equal to their expected values.

        All dimension measurements are in points."""

        self.assertEqual(expected_shape_type, shape.shape_type)
        self.assertEqual(expected_name, shape.name)
        self.assertEqual(expected_width, shape.width)
        self.assertEqual(expected_height, shape.height)
        self.assertEqual(expected_top, shape.top)
        self.assertEqual(expected_left, shape.left)

    def verify_text_box(self, expected_layout_flow: aw.drawing.LayoutFlow, expected_fit_shape_to_text: bool, expected_text_box_wrap_mode: aw.drawing.TextBoxWrapMode, margin_top: float, margin_bottom: float, margin_left: float, margin_right: float, text_box: aw.drawing.TextBox):
        """Checks whether values of properties of a textbox are equal to their expected values.

        All dimension measurements are in points."""

        self.assertEqual(expected_layout_flow, text_box.layout_flow)
        self.assertEqual(expected_fit_shape_to_text, text_box.fit_shape_to_text)
        self.assertEqual(expected_text_box_wrap_mode, text_box.text_box_wrap_mode)
        self.assertEqual(margin_top, text_box.internal_margin_top)
        self.assertEqual(margin_bottom, text_box.internal_margin_bottom)
        self.assertEqual(margin_left, text_box.internal_margin_left)
        self.assertEqual(margin_right, text_box.internal_margin_right)

    def verify_editable_range(self,
                              expected_id: int,
                              expected_editor_user: str,
                              expected_editor_group: aw.EditorType,
                              editable_range: aw.EditableRange):
        """Checks whether values of properties of an editable range are equal to their expected values."""

        self.assertEqual(expected_id, editable_range.id)
        self.assertEqual(expected_editor_user, editable_range.single_user)
        self.assertEqual(expected_editor_group, editable_range.editor_group)

    def verify_date(self, expected: datetime, actual: datetime, delta: timedelta):
        """Checks whether a DateTime matches an expected value, with a margin of error.

        :param expected: The date/time that we expect the result to be.</param>
        :param actual: The DateTime object being tested.</param>
        :param delta: Margin of error for expectedResult.</param>"""

        self.assertAlmostEqual(expected, actual, delta=delta)

    def fields_are_nested(self, inner_field, outer_field):
        """Checks whether a field contains another complete field as a sibling within its nodes.

        If two fields have the same immediate parent node and therefore their nodes are siblings,
        the "field_start" of the outer field appears before the "field_start" of the inner node,
        and the "field_end" of the outer node appears after the "field_end" of the inner node,
        then the inner field is considered to be nested within the outer field. 

        :param inner_field: The field that we expect to be fully within "outer_field".
        :param outer_field: The field that we to contain "inner_field.
        """
        inner_field_parent = inner_field.start.parent_node.as_composite_node()

        self.assertEqual(inner_field_parent, outer_field.start.parent_node)
        self.assertGreater(
            inner_field_parent.child_nodes.index_of(inner_field.start),
            inner_field_parent.child_nodes.index_of(outer_field.start))
        self.assertLess(
            inner_field_parent.child_nodes.index_of(inner_field.end),
            inner_field_parent.child_nodes.index_of(outer_field.end))

    def image_to_byte_array(self, image_path: str) -> bytes:
        """Converts an image to a byte array."""

        with open(image_path, 'rb') as stream:
            buf = io.BytesIO(stream.read())
            return bytes(buf.getbuffer())
