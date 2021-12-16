# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io
import zipfile
from datetime import datetime, date, time, timedelta
from typing import Optional

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class TestUtil(ApiExampleBase):

    @staticmethod
    def verify_image(testcase: unittest.TestCase, expected_width: int, expected_height: int, filename: Optional[str] = None, image_stream: Optional[io.BytesIO] = None):
        """Checks whether a file or a stream contains a valid image with specified dimensions.
            
        Serves to check that an image file is valid and nonempty without looking up its file size.
            
        :param expected_width: Expected width of the image, in pixels.
        :param expected_height: Expected height of the image, in pixels.
        :param filename: Local file system filename of the image file.
        :param image_stream: Stream that contains the image."""

        assert filename is None or image_stream is None
        assert filename is not None or image_stream is not None

        if filename is not None:
            with open(filename, 'rb') as image_stream:
                with drawing.Image.from_stream(image_stream) as image:
                    testcase.assertEqual(expected_width, image.width)
                    testcase.assertEqual(expected_height, image.height)
        else:
            with drawing.Image.from_stream(image_stream) as image:
                testcase.assertEqual(expected_width, image.width)
                testcase.assertEqual(expected_height, image.height)
            
    def image_contains_transparency(testcase: unittest.TestCase, filename: str):
        """Checks whether an image from the local file system contains any transparency.
        
        :param filename: Local file system filename of the image file."""

        with drawing.Image.from_file(filename) as image:
            for x in range(image.width):
                for y in range(image.height):
                    if image.get_pixel(x, y).a != 255:
                        return

        raise Exception(f"The image from \"{filename}\" does not contain any transparency.")

    ## <summary>
    ## Checks whether an HTTP request sent to the specified address produces an expected web response.
    ## </summary>
    ## <remarks>
    ## Serves as a notification of any URLs used in code examples becoming unusable in the future.
    ## </remarks>
    ## <param name="expectedHttpStatusCode">Expected result status code of a request HTTP "HEAD" method performed on the web address.</param>
    ## <param name="webAddress">URL where the request will be sent.</param>
    #internal static void VerifyWebResponseStatusCode(HttpStatusCode expectedHttpStatusCode, string webAddress)

    #    HttpWebRequest request = (HttpWebRequest)WebRequest.create(webAddress)
    #    request.Method = "HEAD"

    #    self.assertEqual(expectedHttpStatusCode, ((HttpWebResponse)request.get_response()).StatusCode)

    ## <summary>
    ## Checks whether an SQL query performed on a database file stored in the local file system
    ## produces a result that resembles the contents of an Aspose.Words table.
    ## </summary>
    ## <param name="expected_result">Expected result of the SQL query in the form of an Aspose.Words table.</param>
    ## <param name="dbFilename">Local system filename of a database file.</param>
    ## <param name="sqlQuery">Microsoft.Jet.OLEDB.4.0-compliant SQL query.</param>
    #internal static void TableMatchesQueryResult(Table expected_result, string dbFilename, string sqlQuery)

    ##if NET48 || NET5_0 || JAVA
    #    (OleDbConnection connection = OleDbConnection())

    #        connection.ConnectionString = f"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbFilename};"
    #        connection.open()

    #        OleDbCommand command = connection.create_command()
    #        command.CommandText = sqlQuery
    #        OleDbDataReader reader = command.execute_reader(CommandBehavior.CloseConnection)

    #        myDataTable = DataTable()
    #        myDataTable.load(reader)

    #        self.assertEqual(expected_result.Rows.Count, myDataTable.Rows.Count)
    #        self.assertEqual(expected_result.Rows[0].Cells.Count, myDataTable.Columns.Count)

    #        for (int i = 0; i < myDataTable.Rows.Count; i++)
    #            for (int j = 0; j < myDataTable.Columns.Count; j++)
    #                self.assertEqual(expected_result.Rows[i].Cells[j].get_text().replace(ControlChar.Cell, ""),
    #                    myDataTable.Rows[i][j].to_string())

    ##endif

    ## <summary>
    ## Checks whether a document produced during a mail merge contains every element of every table produced by a list of consecutive SQL queries on a database.
    ## </summary>
    ## <remarks>
    ## Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
    ## </remarks>
    ## <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
    ## <param name="sqlQueries">List of SQL queries performed on the database all of whose results we expect to find in the document.</param>
    ## <param name="doc">Document created during a mail merge.</param>
    ## <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    #internal static void MailMergeMatchesQueryResultMultiple(string dbFilename, string[] sqlQueries, Document doc, bool onePagePerRow)

    #    foreach (string query in sqlQueries)
    #        MailMergeMatchesQueryResult(dbFilename, query, doc, onePagePerRow)

    ## <summary>
    ## Checks whether a document produced during a mail merge contains every element of a table produced by an SQL query on a database.
    ## </summary>
    ## <remarks>
    ## Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
    ## </remarks>
    ## <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
    ## <param name="sqlQuery">SQL query performed on the database all of whose results we expect to find in the document.</param>
    ## <param name="doc">Document created during a mail merge.</param>
    ## <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    #internal static void MailMergeMatchesQueryResult(string dbFilename, string sqlQuery, Document doc, bool onePagePerRow)

    ##if NET48 || JAVA
    #    expectedStrings = List<string[]>()
    #    string connectionString = @"Driver={Microsoft Access Driver (*.mdb)};Dbq=" + dbFilename

    #    (OdbcConnection connection = OdbcConnection())

    #        connection.ConnectionString = connectionString
    #        connection.open()

    #        OdbcCommand command = connection.create_command()
    #        command.CommandText = sqlQuery

    #        using (OdbcDataReader reader = command.execute_reader(CommandBehavior.CloseConnection))

    #            while (reader.read())

    #                row = string[reader.FieldCount]

    #                for (int i = 0; i < reader.FieldCount; i++)
    #                    switch (reader[i])

    #                        case decimal d:
    #                            row[i] = d.to_string("G29")
    #                            break
    #                        case string s:
    #                            row[i] = s.strip().replace("\n", "")
    #                            break
    #                        default:
    #                            row[i] = ""
    #                            break

    #                expectedStrings.add(row)

    #    MailMergeMatchesArray(expectedStrings.to_array(), doc, onePagePerRow)
    ##endif

    ## <summary>
    ## Checks whether a document produced during a mail merge contains every element of every DataTable in a DataSet.
    ## </summary>
    ## <param name="expected_result">DataSet containing DataTables which contain values that we expect the document to contain.</param>
    ## <param name="doc">Document created during a mail merge.</param>
    ## <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    #internal static void MailMergeMatchesDataSet(DataSet dataSet, Document doc, bool onePagePerRow)

    #    foreach (DataTable table in dataSet.Tables)
    #        MailMergeMatchesDataTable(table, doc, onePagePerRow)

    ## <summary>
    ## Checks whether a document produced during a mail merge contains every element of a DataTable.
    ## </summary>
    ## <param name="expected_result">Values from the mail merge data source that we expect the document to contain.</param>
    ## <param name="doc">Document created during a mail merge.</param>
    ## <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    #internal static void MailMergeMatchesDataTable(DataTable expected_result, Document doc, bool onePagePerRow)

    #    expectedStrings = string[expected_result.Rows.Count][]

    #    for (int i = 0; i < expected_result.Rows.Count; i++)
    #        expectedStrings[i] = Array.convert_all(expected_result.Rows[i].ItemArray, x => x.to_string())

    #    MailMergeMatchesArray(expectedStrings, doc, onePagePerRow)

    ## <summary>
    ## Checks whether a document produced during a mail merge contains every element of an array of arrays of strings.
    ## </summary>
    ## <remarks>
    ## Only suitable for rectangular arrays.
    ## </remarks>
    ## <param name="expected_result">Values from the mail merge data source that we expect the document to contain.</param>
    ## <param name="doc">Document created during a mail merge.</param>
    ## <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    #internal static void MailMergeMatchesArray(string[][] expected_result, Document doc, bool onePagePerRow)

    #    try

    #        if onePagePerRow:

    #            string[] docTextByPages = doc.get_text().strip().split(new[] { ControlChar.PageBreak }, StringSplitOptions.RemoveEmptyEntries)

    #            for (int i = 0; i < expected_result.Length; i++)
    #                for (int j = 0; j < expected_result[0].Length; j++)
    #                    if (!docTextByPages[i].contains(expected_result[i][j])) throw new ArgumentException(expected_result[i][j])

    #        else:

    #            string docText = doc.get_text()

    #            for (int i = 0; i < expected_result.Length; i++)
    #                for (int j = 0; j < expected_result[0].Length; j++)
    #                    if (!docText.contains(expected_result[i][j])) throw new ArgumentException(expected_result[i][j])

    #    catch (ArgumentException e)

    #        Assert.fail(f"String \"{e.Message}\" not found in {(doc.OriginalFileName == null ? "a virtual document" : doc.OriginalFileName.split('\\').last())}.")

    @staticmethod
    def doc_package_file_contains_string(expected: str, doc_filename: str, doc_part_filename: str):
        """Checks whether a file inside a document's OOXML package contains a string.
        
        If an output document does not have a testable value that can be found as a property in its object when loaded,
        the value can sometimes be found in the document's OOXML package.
        
        :param expected: The string we are looking for.
        :param doc_filename: Local file system filename of the document.
        :param doc_part_filename: Name of the file within the document opened as a .zip that is expected to contain the string."""
        
        with zipfile.ZipFile(doc_filename) as archive:
            with archive.open(doc_part_filename) as stream:
                TestUtil.stream_contains_string(expected, stream)

    ## <summary>
    ## Checks whether a file in the local file system contains a string in its raw data.
    ## </summary>
    ## <param name="expected">The string we are looking for.</param>
    ## <param name="filename">Local system filename of a file which, when read from the beginning, should contain the string.</param>
    #internal static void FileContainsString(string expected, string filename)

    #    if !IsRunningOnMono():

    #        (Stream stream = FileStream(filename, FileMode.Open))

    #            stream_contains_string(expected, stream)
    
    @staticmethod
    def stream_contains_string(expected: str, stream: io.BytesIO):
        """Checks whether a stream contains a string.
        
        :param expected: The string we are looking for.
        :param stream: The stream which, when read from the beginning, should contain the string."""

        expected_sequence = expected.encode('utf-8')

        sequence_match_length = 0
        while sequence_match_length < len(expected_sequence):
            actual = stream.read(1)
            if not actual:
                raise Exception(f'String "{expected}" not found in the provided source.')
            
            if actual[0] == expected_sequence[sequence_match_length]:
                sequence_match_length += 1
            else:
                sequence_match_length = 0

    @staticmethod
    def verify_field(testcase: unittest.TestCase,
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

        testcase.assertEqual(expected_type, field.type)
        testcase.assertEqual(expected_field_code, field.get_field_code(True))
        testcase.assertEqual(expected_result, field.result)

    @staticmethod
    def verify_datetime_field(testcase: unittest.TestCase,
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

        testcase.assertEqual(expected_type, field.type)
        testcase.assertEqual(expected_field_code, field.get_field_code(True))
        
        if field.type == aw.fields.FieldType.FIELD_TIME:
            actual = datetime.strptime(field.result, "%H:%M:%S")
            expected = datetime.combine(date(1900, 1, 1), expected_result.time())
            testcase.assertAlmostEqual(expected, actual, delta=delta)
        else:
            actual = datetime.strptime(field.result, "%d/%m/%Y")
            expected = datetime.combine(expected_result.date(), time())
            testcase.assertAlmostEqual(expected, actual, delta=delta)

    ## <summary>
    ## Checks whether a field contains another complete field as a sibling within its nodes.
    ## </summary>
    ## <remarks>
    ## If two fields have the same immediate parent node and therefore their nodes are siblings,
    ## the FieldStart of the outer field appears before the FieldStart of the inner node,
    ## and the FieldEnd of the outer node appears after the FieldEnd of the inner node,
    ## then the inner field is considered to be nested within the outer field.
    ## </remarks>
    ## <param name="innerField">The field that we expect to be fully within outerField.</param>
    ## <param name="outerField">The field that we to contain innerField.</param>
    #internal static void FieldsAreNested(Field innerField, Field outerField)

    #    CompositeNode innerFieldParent = innerField.Start.ParentNode

    #    self.assertTrue(innerFieldParent == outerField.Start.ParentNode)
    #    self.assertTrue(innerFieldParent.ChildNodes.index_of(innerField.Start) > innerFieldParent.ChildNodes.index_of(outerField.Start))
    #    self.assertTrue(innerFieldParent.ChildNodes.index_of(innerField.End) < innerFieldParent.ChildNodes.index_of(outerField.End))

    @staticmethod
    def verify_image_in_shape(testcase: unittest.TestCase,
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

        testcase.assertTrue(image_shape.has_image)
        testcase.assertEqual(expected_image_type, image_shape.image_data.image_type)
        testcase.assertEqual(expected_width, image_shape.image_data.image_size.width_pixels)
        testcase.assertEqual(expected_height, image_shape.image_data.image_size.height_pixels)

    # <summary>
    # 
    # </summary>
    
    @staticmethod
    def verify_footnote(testcase: unittest.TestCase,
                        expected_footnote_type: aw.notes.FootnoteType,
                        expected_is_auto: bool,
                        expected_reference_mark: str,
                        expected_contents: str,
                        footnote: aw.notes.Footnote):
        """Checks whether values of a footnote's properties are equal to their expected values.
        
        :param expected_footnote_type: Expected type of the footnote/endnote.</param>
        :param expected_is_auto: Expected auto-numbered status of this footnote.</param>
        :param expected_reference_mark: If "is_auto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
        :param expected_contents: Expected side comment provided by the footnote.</param>
        :param footnote: Footnote node in question.</param>"""

        testcase.assertEqual(expected_footnote_type, footnote.footnote_type)
        testcase.assertEqual(expected_is_auto, footnote.is_auto)
        testcase.assertEqual(expected_reference_mark, footnote.reference_mark)
        testcase.assertEqual(expected_contents, footnote.to_string(aw.SaveFormat.TEXT).strip())

    @staticmethod
    def verify_list_level(testcase: unittest.TestCase,
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

        testcase.assertEqual(expected_list_format, list_level.number_format)
        testcase.assertEqual(expected_number_position, list_level.number_position)
        testcase.assertEqual(expected_number_style, list_level.number_style)

    @staticmethod
    def copy_stream(src_stream: io.BytesIO, dst_stream: io.BytesIO):
        """Copies from the current position in src stream till the end.
        Copies into the current position in dst stream."""

        assert src_stream is not None
        assert dst_stream is not None

        dst_stream.write(src_stream.read())

    def dump_array(data: bytes, start: int, count: int) -> str:
        """Dumps byte array into a string."""

        if data is None:
            return "Null"

        result = ""
        while count > 0:
            result += f"{data[start]:02X} "
            start += 1
            count -= 1

        return result

    def verify_tab_stop(testcase: unittest.TestCase, 
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

        testcase.assertEqual(expected_position, tab_stop.position)
        testcase.assertEqual(expected_tab_alignment, tab_stop.alignment)
        testcase.assertEqual(expected_tab_leader, tab_stop.leader)
        testcase.assertEqual(is_clear, tab_stop.is_clear)

    @staticmethod
    def verify_shape(testcase: unittest.TestCase, expected_shape_type: aw.drawing.ShapeType, expected_name: str, expected_width: float, expected_height: float, expected_top: float, expected_left: float, shape: aw.drawing.Shape):
        """Checks whether values of a shape's properties are equal to their expected values.
        
        All dimension measurements are in points."""

        testcase.assertEqual(expected_shape_type, shape.shape_type)
        testcase.assertEqual(expected_name, shape.name)
        testcase.assertEqual(expected_width, shape.width)
        testcase.assertEqual(expected_height, shape.height)
        testcase.assertEqual(expected_top, shape.top)
        testcase.assertEqual(expected_left, shape.left)

    @staticmethod
    def verify_text_box(testcase: unittest.TestCase, expected_layout_flow: aw.drawing.LayoutFlow, expected_fit_shape_to_text: bool, expected_text_box_wrap_mode: aw.drawing.TextBoxWrapMode, margin_top: float, margin_bottom: float, margin_left: float, margin_right: float, text_box: aw.drawing.TextBox):
        """Checks whether values of properties of a textbox are equal to their expected values.
        
        All dimension measurements are in points."""

        testcase.assertEqual(expected_layout_flow, text_box.layout_flow)
        testcase.assertEqual(expected_fit_shape_to_text, text_box.fit_shape_to_text)
        testcase.assertEqual(expected_text_box_wrap_mode, text_box.text_box_wrap_mode)
        testcase.assertEqual(margin_top, text_box.internal_margin_top)
        testcase.assertEqual(margin_bottom, text_box.internal_margin_bottom)
        testcase.assertEqual(margin_left, text_box.internal_margin_left)
        testcase.assertEqual(margin_right, text_box.internal_margin_right)

    @staticmethod
    def verify_editable_range(testcase: unittest.TestCase,
                              expected_id: int,
                              expected_editor_user: str,
                              expected_editor_group: aw.EditorType,
                              editable_range: aw.EditableRange):
        """Checks whether values of properties of an editable range are equal to their expected values."""

        testcase.assertEqual(expected_id, editable_range.id)
        testcase.assertEqual(expected_editor_user, editable_range.single_user)
        testcase.assertEqual(expected_editor_group, editable_range.editor_group)

    @staticmethod
    def verify_date(testcase: unittest.TestCase, expected: datetime, actual: datetime, delta: timedelta):
        """Checks whether a DateTime matches an expected value, with a margin of error.
            
        :param expected: The date/time that we expect the result to be.</param>
        :param actual: The DateTime object being tested.</param>
        :param delta: Margin of error for expectedResult.</param>"""
        
        testcase.assertAlmostEqual(expected, actual, delta=delta)
