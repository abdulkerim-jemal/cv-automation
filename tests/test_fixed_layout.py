from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).parents[1]))


def test_templates_exist():
    root=Path(__file__).parents[1]
    assert (root/'Asail'/'Experianced.docx').exists()
    assert (root/'Asail'/'Non Experianced.docx').exists()
    assert (root/'Al Zaid'/'Experianced.docx').exists()
    assert (root/'Al Zaid'/'Non_Experianced.docx').exists()
