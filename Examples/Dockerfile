FROM python:3.7-buster
WORKDIR /usr/src/app

# Copy and install Aspose.Words wheel.
COPY Bin/Aspose.Words-20.6-py3-none-manylinux1_x86_64.whl .
RUN pip install Aspose.Words-20.6-py3-none-manylinux1_x86_64.whl

# Copy api examples python files.
COPY ApiExamples/*.py .
CMD ["ExDocument.py"]
ENTRYPOINT ["python3"]