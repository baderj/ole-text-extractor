import os

import pytest
from olefile.olefile import OleDirectoryEntry
from src.oletextextract import OLETextExtract


@pytest.mark.parametrize(
    "filename",
    [("doc_english.doc"), ("doc_russian.doc"), ("doc_german.doc"), ("doc_mixed.doc")],
)
def test_eval(filename):
    path = os.path.join("test", "testfiles", filename)
    result = os.path.join("test", "testfiles", filename + ".sol")

    ote = OLETextExtract()
    text = ote.extract(path)

    with open(result, "rb") as r:
        solution = r.read().decode("utf-8")

    assert solution == text
