![PyPI](https://img.shields.io/pypi/v/aspose-words.svg?label=PyPI) ![PyPI](https://img.shields.io/pypi/dm/aspose-words.svg?label=PyPI%20downloads) ![GitHub](https://img.shields.io/github/license/aspose-words/Aspose.Words-for-Python-via-.NET)

# Word Processing API for Python

[Aspose.Words for Python](https://products.aspose.com/words/python-net/) is a powerful on-premise class library that can be used for numerous document processing tasks. It enables developers to enhance their own applications with features such as generating, modifying, converting and rendering documents, without relying on third-party applications, for example, Microsoft Word, or Office Automation.

This repository contains a collection of Python examples that help you learn and explore the API features.

## Word API Features

The following are some popular features of Aspose.Words for Python:

- Aspose.Words can be used to develop applications for a vast range of operating systems such as Windows or Linux.
- Comprehensive [document import and export](https://docs.aspose.com/words/python-net/loading-saving-and-converting/) with [35+ supported file formats](https://docs.aspose.com/words/python-net/supported-document-formats/). This allows users to [convert documents](https://docs.aspose.com/words/python-net/convert-a-document/) from one popular format to another, for example, from DOCX into PDF or Markdown, or from PDF into various Word formats.
- Programmatic access to the formatting properties of all document elements. For example, using Aspose.Words users can [split a document](https://docs.aspose.com/words/python-net/split-a-document/) into parts or [compare two documents](https://docs.aspose.com/words/python-net/compare-documents/).
- [High fidelity rendering](https://docs.aspose.com/words/python-net/rendering/) of document pages. For example, if it is needed to render a document as in Microsoft Word, Aspose.Words will successfully cope with this task.
- [Generate reports with Mail Merge](https://docs.aspose.com/words/python-net/mail-merge-and-reporting/), which allows filling in merge templates with data from various sources to create merged documents.
- LINQ Reporting Engine to fetch data from databases, XML, JSON, OData, external documents, and much more.

Try our [free online Apps](https://products.aspose.app/words/family) demonstrating some of the most popular Aspose.Words functionality.

## Supported Document Formats
Aspose.Words for Python supports [a wide range of formats for loading and saving documents](https://docs.aspose.com/words/python-net/supported-document-formats/), some of them are listed below:

**Microsoft Word:** DOC, DOT, DOCX, DOTX, DOTM, FlatOpc, FlatOpcMacroEnabled, FlatOpcTemplate, FlatOpcTemplateMacroEnabled, RTF, WordML, DocPreWord60\
**OpenDocument:** ODT, OTT\
**Web:** HTML, MHTML\
**Markdown:** MD\
**Markup:** XamlFixed, HtmlFixed, XamlFlow, XamlFlowPack\
**Fixed Layout:** PDF, XPS, OpenXps\
**Image:** SVG, TIFF, PNG, BMP, JPEG, GIF\
**Metafile:** EMF\
**Printer:** PCL, PS\
**Text:** TXT\
**eBook:** MOBI, CHM, EPUB

## Platform Independence

Aspose.Words for Python can be used to develop applications for a vast range of operating systems, such as Windows and Linux, where Python 3.5 or later is installed. You can build both 32-bit and 64-bit Python applications.

## Get Started

Ready to give Aspose.Words for Python a try?

Simply run ```pip install aspose-words``` from the Console to fetch the package.
If you already have Aspose.Words for Python and want to upgrade the version, please run ```pip install --upgrade aspose-words``` to get the latest version.

You can run the following snippets in your environment to see how Aspose.Words works, or check out the [Examples](https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET/tree/master/Examples/DocsExamples/DocsExamples) or [Aspose.Words for Python Documentation](https://docs.aspose.com/words/python-net/) for other common use cases.

## Using Python to Create a DOCX File from Scratch

Aspose.Words for Python allows you to create a new blank document and add content to this document.

```python
import aspose.words as aw

# Create a blank document.
doc = aw.Document()

# Use a document builder to add content to the document.
builder = aw.DocumentBuilder(doc)
# Write a new paragraph in the document with the text "Hello World!".
builder.writeln("Hello World!")

# Save the document in DOCX format. Save format is automatically determined from the file extension.
doc.save("output.docx")
```

## Using Python to Convert a Word Document to HTML

Aspose.Words for Python also allows you to convert Microsoft Word formats to PDF, XPS, Markdown, HTML, JPEG, TIFF, and other file formats. The following snippet demonstrates the conversion from DOCX to HTML:

```python
import aspose.words as aw

# Load the document from the disc.
doc = aw.Document("TestDocument.docx")

# Save the document to HTML format.
doc.save("output.html")
```

## Using Python to Import PDF and Save as a DOCX File

In addition, you can import a PDF document into your Python application and export it as a DOCX format file without the need to install Microsoft Word:

```python
import aspose.words as aw

# Load the PDF document from the disc.
doc = aw.Document("TestDocument.pdf")

# Save the document to DOCX format.
doc.save("output.docx")
```
## Docker File
The Dockerfile includes commands for configuring Linux images to enable the use of Aspose.Words for Python via .Net. 

[Product Page](https://products.aspose.com/words/python-net/) | [Docs](https://docs.aspose.com/words/python-net/) | [Demos](https://products.aspose.app/words/family) | [Examples](https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET/tree/master/Examples) | [Blog](https://blog.aspose.com/category/words/) | [Search](https://search.aspose.com/) | [Free Support](https://forum.aspose.com/c/words) | [Temporary License](https://purchase.aspose.com/temporary-license)
