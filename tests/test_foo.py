
from json2excel import Excel2JSON
from exceptions import MissingRequiredError, InvalidNumericValueError

import os
import json
import pytest

def test_標準的な変換処理():

    src = "tests/resource/001/test.xlsx"
    config = "tests/resource/001/convert.yaml"
    with open("tests/resource/001/expected.json") as f:
        expected = json.load(f)

    ret = Excel2JSON.xls_with_yaml2json(src, config)

    assert ret == expected

def test_table変換する際の列名にブランク列が途中に含まれる場合():

    src = "tests/resource/002/test.xlsx"
    config = "tests/resource/002/convert.yaml"
    with open("tests/resource/002/expected.json") as f:
        expected = json.load(f)

    ret = Excel2JSON.xls_with_yaml2json(src, config)

    assert ret == expected


def test_table変換する際の列名に改行コードが含まれる場合():

    src = "tests/resource/003/test.xlsx"
    config = "tests/resource/003/convert.yaml"
    with open("tests/resource/003/expected.json") as f:
        expected = json.load(f)

    ret = Excel2JSON.xls_with_yaml2json(src, config)

    assert ret == expected


def test_table変換する際の列名を正規表現で一致させて置換する場合():

    src = "tests/resource/004/test.xlsx"
    config = "tests/resource/004/convert.yaml"
    with open("tests/resource/004/expected.json") as f:
        expected = json.load(f)

    ret = Excel2JSON.xls_with_yaml2json(src, config)

    assert ret == expected

def test_必須の項目が存在しない場合_type_list():

    src = "tests/resource/005/test.xlsx"
    config = "tests/resource/005/convert.yaml"

    with pytest.raises(MissingRequiredError) as e:
        ret = Excel2JSON.xls_with_yaml2json(src, config)

    assert str(e.value) == "Required 'code' is missing."

def test_必須の項目が存在しない場合_type_cell():

    src = "tests/resource/006/test.xlsx"
    config = "tests/resource/006/convert.yaml"

    with pytest.raises(MissingRequiredError) as e:
        ret = Excel2JSON.xls_with_yaml2json(src, config)

    assert str(e.value) == "Required 'hoge' is missing."

def test_数値のみ許容する場合_type_cell():

    src = "tests/resource/007/test.xlsx"
    config = "tests/resource/007/convert.yaml"

    with pytest.raises(InvalidNumericValueError) as e:
        ret = Excel2JSON.xls_with_yaml2json(src, config)

    assert str(e.value) == "'hoge' is Invalid Numeric Value."




