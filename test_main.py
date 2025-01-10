import main

def test_file_exist():
    test_files = main.Files(r"old/Inhome-parser/inhome.xlsx")
    assert test_files.file_exist() == True
