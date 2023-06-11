import fitz
import os
from openpyxl import Workbook, load_workbook
import sys
from PyPDF2 import PdfMerger, PdfFileReader, PdfFileWriter
from sys import platform
from datetime import date, datetime, timedelta
import time
import glob
import shutil
import easyocr
from pathlib import Path
from pdf2image import convert_from_path
import numpy
import pandas as pd
from random import randint
import json
import warnings

warnings.filterwarnings("ignore", category=Warning)

def generate_xls_from_pdf(fileinput):
    print("Generate new XLS file from PDF File...", end=" ", flush=True)
    pdfFileObject = open(fileinput, 'rb')
    pdfReader = PdfFileReader(pdfFileObject)
    reader = easyocr.Reader(['en'], gpu=False, verbose=False)
    print(pdfReader.numPages)
    for i in range(0, pdfReader.numPages):
        if os.path.exists(Path("pdftmp.pdf")):
            os.remove(Path("pdftmp.pdf"))
        pdfWriter = PdfFileWriter()
        pdf = pdfReader.getPage(i)
        pdfWriter.addPage(pdf)
        with Path("pdftmp.pdf").open(mode="wb") as output_file:
            pdfWriter.write(output_file)
        images = convert_from_path(Path('pdftmp.pdf'))
        imgcrop = images[0].crop(box = (180,750,750,900))
        res = reader.readtext(numpy.array(imgcrop)  , detail = 0)
        tracking_id = res[1].strip().replace('TRACKING #','').replace(' ','').replace(":","")
        # print(tracking_id)
        lines = pdf.extract_text(space_width=200).split("\n")
        # print(lines)
        if platform == "win32":
            weight = lines[1].split('-')[1].strip().split("lb")[0]
        else:
            weight = lines[0].split('-')[1].strip().replace("lb","")    
            
        to = lines[7]
        address = lines[8]
        shipper = lines[2]
        code = lines[12]
        box = lines[-2].replace(":", " ")
        input(tracking_id)
    print("Finished")


generate_xls_from_pdf(r"C:\synergy-data-tester\shipment_creation_2023-05-30\package-FBA176VGRJ9Y.pdf")