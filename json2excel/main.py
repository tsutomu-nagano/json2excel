# from svglib.svglib import svg2rlg
# from reportlab.graphics import renderPDF, renderPM

from typing import List, Optional
from fastapi import Depends, FastAPI, File, UploadFile, Form, HTTPException, Query
from pydantic import BaseModel
from fastapi.responses import StreamingResponse, PlainTextResponse, FileResponse, JSONResponse
from fastapi.encoders import jsonable_encoder

from pathlib import Path

import pandas as pd
import json
import datetime
import io
import sys
import os
import re
import yaml
import tempfile
import shutil
import requests
import zipfile

from openpyxl import load_workbook

from pandas import json_normalize

import json2excel
import excel_decorator 
import exceptions


tags_metadata = [
    {
        "name": "convert",
        "description": "変換する"
    },

]


app = FastAPI(
    title="json 2 excel",
    description="JSONファイルをEXCELに変換するAPI",
    version="0.0.1",
    # contact={
    #     "name": "Deadpoolio the Amazing",
    #     "url": "http://x-force.example.com/contact/",
    #     "email": "dp@x-force.example.com",
    # },
    # license_info={
    #     "name": "che 2.0",
    #     "url": "https://www.apache.org/licenses/LICENSE-2.0.html",
    # },
    openapi_tags=tags_metadata
)




@app.post(
        "/json2excel",
        summary="変換定義（YAML）ファイルを基にJSONファイルをExcelファイルに変換します",
        tags = ["convert"])
def json_to_excel(
    source_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
    conversion_config_file: UploadFile = File(...),
    decorate_config_file: UploadFile = File(None)
    ):

    with tempfile.NamedTemporaryFile(delete=False, suffix = Path(source_file.filename).suffix) as src,\
         tempfile.NamedTemporaryFile(delete=False, suffix = Path(template_file.filename).suffix) as temp, \
         tempfile.NamedTemporaryFile(delete=False, suffix = Path(template_file.filename).suffix) as dest, \
         tempfile.NamedTemporaryFile(delete=False, suffix = Path(conversion_config_file.filename).suffix) as conv:

        shutil.copyfileobj(source_file.file, src)
        shutil.copyfileobj(template_file.file, temp)
        shutil.copyfileobj(template_file.file, dest)
        shutil.copyfileobj(conversion_config_file.file, conv)

        src.seek(0)
        conv.seek(0)

        json2excel.JSON2Excel.json_with_yaml2xls(src.name, conv.name, temp.name, dest.name)

        # decorator があった場合は処理する
        if decorate_config_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix = Path(decorate_config_file.filename).suffix) as deco:

                shutil.copyfileobj(decorate_config_file.file, deco)

                deco.seek(0)

                excel_decorator.ExcelDecorator.decoration(dest = dest.name,config = deco.name)


        headers = {'Content-Disposition': 'attachment; filename="dest.xlsx"'}
        return FileResponse(dest.name, headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.post(
        "/decoration",
        summary="変換定義（YAML）ファイルを基にExcelファイルにハイパーリンクなどの設定を追加します",
        tags = ["convert"])
def excel_decoration(
    source_file: UploadFile = File(...),
    config_file: UploadFile = File(...),
    ):

    with tempfile.NamedTemporaryFile(delete=False, suffix = Path(source_file.filename).suffix) as src,\
         tempfile.NamedTemporaryFile(delete=False, suffix = Path(config_file.filename).suffix) as conf:

        shutil.copyfileobj(source_file.file, src)
        shutil.copyfileobj(config_file.file, conf)

        conf.seek(0)


        excel_decorator.ExcelDecorator.decoration(dest = src.name,config = conf.name)

        headers = {'Content-Disposition': 'attachment; filename="dest.xlsx"'}
        return FileResponse(src.name, headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get(
        "/sample",
        summary="処理に必要な変換定義（YAML）ファイルと簡単なExcelファイルを取得します",
        tags = ["convert"])
def get_sample():

    with tempfile.NamedTemporaryFile(delete=False, dir =".", suffix = ".zip") as t1:

        with zipfile.ZipFile(t1.name, 'w') as myzip:
            myzip.write("sample/convert.yaml")
            myzip.write('sample/test.xlsx')

        return FileResponse(path=t1.name, filename="sample.zip")


