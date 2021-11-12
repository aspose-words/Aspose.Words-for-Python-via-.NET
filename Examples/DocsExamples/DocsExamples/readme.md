To run examples from command line, chnage the directory to `\Aspose.Words-for-Python-via-.NET\Examples\DocsExamples\DocsExamples`

Then run one of the following commands:

```
python -m unittest file_formats_and_conversions.base_conversions
```

To run examples from the particular class:

```
python -m unittest file_formats_and_conversions.base_conversions.BaseConversions
```

To run one particular example:
```
python -m unittest file_formats_and_conversions.base_conversions.BaseConversions.test_doc_to_docx
```

To Run all examples:
```
python -m unittest discover -p *.py
```

To run on Linux docker:
```
docker build -t aw_py_examples .
```
then:
```
docker run --mount type=bind,source=D:\Aspose.Words-for-Python-via-.NET\Examples\Data,target=/usr/src/app/Aspose.Words-for-Python-via-.NET/Examples/Data --mount type=bind,source=D:\Licenses,target=/usr/src/app/Licenses --rm aw_py_examples -m unittest discover -p *.py
```
Where `D:\Aspose.Words-for-Python-via-.NET\Examples\Data` path to Data folder in this repo and `D:\Licenses` path to folder with `Aspose.Words.Python.NET.lic`, path can be modified in the Dockerfile.