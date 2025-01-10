import main
import pytest


# Test Files Class
def test_file_exist():
    test_files_1 = main.Files(r"old/Inhome-parser/inhome.xlsx")
    test_files_2 = main.Files(r"old/Inhome-parser/inhome1.xlsx")
    assert test_files_1.file_exist() == True
    assert test_files_2.file_exist() == False


# Test Reader Class
def test_R_file_exist():
    test_files_1 = main.Reader(r"old/Inhome-parser/inhome.xlsx")
    test_files_2 = main.Reader(r"old/Inhome-parser/inhome1.xlsx")
    assert test_files_1.file_exist() == True
    test_list = test_files_1.get_list()
    assert test_files_1.get_list() == test_list
    assert test_files_2.file_exist() == False
    assert test_files_2.get_list() == []


def test_data_1():
    pass


def test_get_list():
    pass
