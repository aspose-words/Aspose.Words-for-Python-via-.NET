FROM python:3.7-buster
WORKDIR /usr/src/app

# Install aspose-words.
RUN pip install -i https://test.pypi.org/simple/ aspose-words

ENV ASPOSE_WORDS_PYTHON_LICENSE=/usr/src/app/Licenses/Aspose.Total.Python.NET.lic

# Copy docs examples python files.
COPY . .
ENTRYPOINT ["python3"]