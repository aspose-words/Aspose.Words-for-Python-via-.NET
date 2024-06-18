import sys
import aspose.words as aw
import requests
from io import BytesIO

def insert_image_after_phrase(doc_path, phrase_match, image_url):
    lic = aw.License()
    lic.set_license("X:\GitHub\Examples\Aspose.Words-for-Python-via-.NET\Examples\Data\License\Aspose.Words.Python.NET.lic")

    # Create Aspose Words document object
    document = aw.Document(doc_path)
    builder = aw.DocumentBuilder(document)
    # Download the image from the URL
    response = requests.get(image_url)
    image_bytes = BytesIO(response.content)

    for r in document.get_child_nodes(aw.NodeType.RUN, True):
        run = r.as_run()
        if phrase_match in run.text:
            builder.move_to(run.next_sibling)
            builder.insert_image(image_bytes)
    # Save the document
    document.save(doc_path)

# Usage example:
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python script.py <doc_path> <phrase_match> <image_url>")
        sys.exit(1)

    doc_path = sys.argv[1]
    phrase_match = sys.argv[2]
    image_url = sys.argv[3]

    insert_image_after_phrase(doc_path, phrase_match, image_url)